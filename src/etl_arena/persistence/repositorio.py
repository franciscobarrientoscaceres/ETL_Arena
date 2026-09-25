"""``RepositorioCorridas``: persistencia append-only por ``IdCorrida`` (R11, R12; ADR-08).

* ``iniciar`` inserta ``etl_run`` (Estado = running) en su propia transacción.
* ``guardar_corrida`` inserta **todas** las tablas de la corrida en **una** transacción, por lotes
  con ``fast_executemany``; ante cualquier error hace rollback completo y relanza.
* ``finalizar`` solo actualiza las columnas de estado que ``etl_writer`` puede modificar.
Nunca ``DELETE`` ni ``UPDATE`` de filas de otras corridas.
"""

from __future__ import annotations

import json
import time
from calendar import monthrange
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import datetime
from typing import Literal

from sqlalchemy.engine import Engine

from etl_arena.config import ConfiguracionCalculo
from etl_arena.model import FilaReconciliacion, KpiMensual, ReferenciaExcel
from etl_arena.persistence.paquete import LOTE, MetadatosCorrida, PaqueteCorrida, Tabla, _dt, _f, fila_etl_run

EstadoFinal = Literal["success", "failed", "parity_failed"]


@dataclass
class ResultadoGuardado:
    filas_por_tabla: dict[str, int]
    segundos: float


class RepositorioCorridas:
    def __init__(self, engine: Engine):
        self.engine = engine

    # ------------------------------------------------------------------ utilidades
    def _conexion(self):
        conn = self.engine.raw_connection()
        conn.driver_connection.autocommit = False  # el proxy del pool no propaga el atributo
        return conn

    @staticmethod
    def _insertar(cur, tabla: Tabla) -> None:
        if not tabla.filas:
            return
        sql = tabla.sql_insert()
        cur.fast_executemany = True
        for i in range(0, len(tabla.filas), LOTE):
            cur.executemany(sql, tabla.filas[i : i + LOTE])

    # ------------------------------------------------------------------ catálogo
    def cargar_catalogo(self) -> dict[str, int]:
        """``CodigoFalla → IdTipoDetencion`` (una sola consulta; sin N+1)."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute("SELECT CodigoFalla, IdTipoDetencion FROM dbo.tipo_detencion")
            return {codigo: id_tipo for codigo, id_tipo in cur.fetchall()}
        finally:
            conn.close()

    def kpi_mensuales_vigentes(self, id_proyecto: int, anio: int) -> list[KpiMensual]:
        """Filas vigentes de ``v_monthly_kpi_vigente`` (corridas oficiales + ``excel_manual``) para el anual (R10)."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "SELECT Anio, Mes, DiasMes, BloquesMuestreo, BloquesRacksIndisponibles, Origen, IdCorrida, "
                "DisponibilidadContractual FROM dbo.v_monthly_kpi_vigente WHERE IdProyecto = ? AND Anio = ?",
                id_proyecto,
                anio,
            )
            return [
                KpiMensual(int(a), int(m), float(d), float(b), float(f), o, str(i).lower() if i else None, float(k))
                for a, m, d, b, f, o, i, k in cur.fetchall()
            ]
        finally:
            conn.close()

    # ------------------------------------------------------------------ ciclo de vida de la corrida
    def iniciar(self, cfg: ConfiguracionCalculo, meta: MetadatosCorrida) -> None:
        columnas, valores = fila_etl_run(cfg, meta)
        tabla = Tabla("etl_run", columnas, [valores])
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(tabla.sql_insert(), valores)
            conn.commit()
        finally:
            conn.close()

    def guardar_corrida(self, paquete: PaqueteCorrida) -> ResultadoGuardado:
        """Una transacción para todas las tablas (R11.3). La corrida debe existir y seguir en
        ``running``: nunca se agregan datos a una corrida ya cerrada (append-only, ADR-08)."""
        return self._guardar_tablas(paquete, verificar_running=True)

    def _guardar_tablas(self, paquete: PaqueteCorrida, verificar_running: bool = False) -> ResultadoGuardado:
        inicio = time.perf_counter()
        conn = self._conexion()
        try:
            cur = conn.cursor()
            if verificar_running:
                cur.execute(
                    "SELECT Estado FROM dbo.etl_run WITH (UPDLOCK, HOLDLOCK) WHERE IdCorrida = ?", paquete.id_corrida
                )
                fila = cur.fetchone()
                if fila is None or fila[0] != "running":
                    raise RuntimeError(
                        f"etl_run {paquete.id_corrida} no existe o no está en 'running' "
                        f"({fila[0] if fila else 'inexistente'}): no se agregan datos"
                    )
            for tabla in paquete.tablas:
                self._insertar(cur, tabla)
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
        return ResultadoGuardado({t.nombre: len(t.filas) for t in paquete.tablas}, time.perf_counter() - inicio)

    def finalizar(
        self,
        id_corrida: str,
        estado: EstadoFinal,
        resumen_calidad: dict | None = None,
        mensaje_error: str | None = None,
        minutos_muestreo_derivado: float | None = None,
        id_referencia_excel: str | None = None,
    ) -> None:
        """Transición de estado del propio ``etl_run`` (transacción aparte, R11.2–R11.3)."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "UPDATE dbo.etl_run SET Estado = ?, FinalizadoEn = ?, ResumenCalidad = ?, MensajeError = ?, "
                "MinutosMuestreoDerivado = COALESCE(?, MinutosMuestreoDerivado), "
                "IdReferenciaExcel = COALESCE(?, IdReferenciaExcel) WHERE IdCorrida = ? AND Estado = 'running'",
                estado,
                datetime.now().replace(microsecond=0),
                json.dumps(resumen_calidad, ensure_ascii=False) if resumen_calidad is not None else None,
                mensaje_error,
                _f(minutos_muestreo_derivado),
                id_referencia_excel,
                id_corrida,
            )
            cur.execute("SELECT @@ROWCOUNT")  # no depende de SET NOCOUNT de la sesión reutilizada
            if cur.fetchone()[0] != 1:
                raise RuntimeError(f"etl_run {id_corrida} no existe o ya no está en 'running'")
            conn.commit()
        finally:
            conn.close()

    # ------------------------------------------------------------------ referencia Excel y reconciliación
    def guardar_referencia_excel(self, ref: ReferenciaExcel) -> str:
        idr = ref.id_referencia
        tablas = [
            Tabla(
                "excel_reference_run",
                (
                    "IdReferencia",
                    "Corte",
                    "ArchivoLibro",
                    "HashLibro",
                    "ExtraidoEn",
                    "C5",
                    "C7",
                    "C21",
                    "C31",
                    "L2",
                    "L4",
                    "L14",
                    "DailyD5",
                    "C12",
                    "C14",
                    "C16",
                    "C19",
                    "C23",
                    "L10",
                ),
                [
                    (
                        idr,
                        ref.corte,
                        ref.archivo_libro[:500],
                        ref.hash_libro,
                        _dt(ref.extraido_en),
                        _dt(ref.c5),
                        _dt(ref.c7),
                        ref.c21,
                        ref.c31,
                        _dt(ref.l2),
                        _dt(ref.l4),
                        ref.l14,
                        _dt(ref.daily_d5),
                        ref.c12,
                        _f(ref.c14),
                        _f(ref.c16),
                        _f(ref.c19),
                        _f(ref.c23),
                        _f(ref.l10),
                    )
                ],
            ),
            Tabla(
                "excel_reference_sample",
                ("IdReferencia", "FilaResultado", "SerialFecha", "NumeroPCS", "Valor"),
                [(idr, f, s, pcs, v) for f, s, pcs, v in ref.tabla],
            ),
            Tabla(
                "excel_reference_fault_event",
                ("IdReferencia", "OrdenExcel", "B", "C", "D", "E", "F", "G", "H", "I"),
                [
                    (idr, e.orden_excel, _f(e.b), _f(e.c), _f(e.d), _f(e.e), e.f, e.g, _f(e.h), _f(e.i))
                    for e in ref.eventos
                ],
            ),
            Tabla(
                "excel_reference_daily",
                ("IdReferencia", "DiaN", "Dia", "D", "E", "F", "G"),
                [(idr, d.dia_n, _dt(d.dia), _f(d.d), _f(d.e), _f(d.f), _f(d.g)) for d in ref.diario],
            ),
        ]
        self._guardar_tablas(PaqueteCorrida(idr, tablas))
        return idr

    def guardar_reconciliacion(
        self, id_corrida: str, id_referencia: str | None, filas: Iterable[FilaReconciliacion]
    ) -> int:
        tabla = Tabla(
            "reconciliation_result",
            (
                "IdCorrida",
                "IdReferencia",
                "Nivel",
                "Metrica",
                "Clave",
                "ValorPython",
                "ValorExcel",
                "Delta",
                "Tolerancia",
                "Aprobado",
            ),
            [
                (
                    id_corrida,
                    id_referencia,
                    f.nivel,
                    f.metrica[:100],
                    (f.clave or None) and f.clave[:200],
                    (f.valor_python or None) and f.valor_python[:100],
                    (f.valor_excel or None) and f.valor_excel[:100],
                    _f(f.delta),
                    _f(f.tolerancia),
                    f.aprobado,
                )
                for f in filas
            ],
        )
        self._guardar_tablas(PaqueteCorrida(id_corrida, [tabla]))
        return len(tabla.filas)

    # ------------------------------------------------------------------ cargas mensuales y revisión
    def registrar_carga_exclusion(self, id_proyecto: int, anio: int, mes: int, archivo: str, sha256: str) -> int:
        """Entrega mensual de la ``Exclusion_Matrix`` (D-17)."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO dbo.exclusion_matrix_carga (IdProyecto, Anio, Mes, ArchivoOrigen, Sha256Archivo) "
                "OUTPUT INSERTED.IdCarga VALUES (?, ?, ?, ?, ?)",
                id_proyecto,
                anio,
                mes,
                archivo[:500],
                sha256,
            )
            id_carga = int(cur.fetchone()[0])
            conn.commit()
            return id_carga
        finally:
            conn.close()

    def exclusiones_cargadas(self, id_proyecto: int, anio: int, mes: int) -> bool:
        """``EstadoExclusiones`` de una corrida del mes: ¿ya se cargó la matriz de Alex? (R19.5)."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "SELECT COUNT(*) FROM dbo.exclusion_matrix_carga WHERE IdProyecto = ? AND Anio = ? AND Mes = ?",
                id_proyecto,
                anio,
                mes,
            )
            return cur.fetchone()[0] > 0
        finally:
            conn.close()

    def mes_cerrado(self, id_proyecto: int, anio: int, mes: int) -> bool:
        """¿El mes ya tiene cierre oficial (``cierre_mensual`` success del mes completo) o KPI
        ``excel_manual``? Si no, el orquestador lo encola (R19.3)."""
        inicio = datetime(anio, mes, 1)
        fin = datetime(anio, mes, monthrange(anio, mes)[1])
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "SELECT CASE WHEN EXISTS (SELECT 1 FROM dbo.etl_run WHERE IdProyecto = ? "
                "AND TipoCorrida = N'cierre_mensual' AND EsOficial = 1 AND Estado = N'success' "
                "AND InicioPeriodo = ? AND FinPeriodo = ?) OR EXISTS (SELECT 1 FROM dbo.monthly_official_kpi "
                "WHERE IdProyecto = ? AND Anio = ? AND Mes = ? AND Origen = N'excel_manual') THEN 1 ELSE 0 END",
                id_proyecto,
                inicio,
                fin,
                id_proyecto,
                anio,
                mes,
            )
            return cur.fetchone()[0] == 1
        finally:
            conn.close()

    def registrar_revision(
        self,
        id_proyecto: int,
        numero_pcs: int,
        fecha_inicio: datetime,
        estado: str,
        revisado_por: str,
        observacion: str | None = None,
    ) -> None:
        """Revisión de una detención (rol ``revisor``; historial append-only, R12.5)."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO dbo.detencion_revision (IdProyecto, NumeroPCS, FechaInicio, EstadoRevision, Observacion, "
                "RevisadoPor) VALUES (?, ?, ?, ?, ?, ?)",
                id_proyecto,
                numero_pcs,
                _dt(fecha_inicio),
                estado,
                observacion,
                revisado_por,
            )
            conn.commit()
        finally:
            conn.close()

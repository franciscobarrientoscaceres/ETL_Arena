"""``RepositorioCorridas``: registro de ejecuciones + publicación del estado vigente por mes (R11 rev. 3, ADR-12).

* ``iniciar`` inserta ``etl_run`` (Estado = running) en su propia transacción y devuelve ``NumCorrida``.
* ``publicar_mes`` reemplaza, en **una** transacción con bloqueo por proyecto, todo lo vigente del mes:
  protecciones (D-21, D-23) → correcciones (D-24) → borrar + insertar muestras, día a día y calidad →
  ``MERGE`` de detenciones conservando ``IdDetencion`` (D-25) → fila de ``disponibilidad_mensual``.
  Ante cualquier error (o una protección) hace rollback: el estado anterior queda intacto.
* ``finalizar`` cierra el ``etl_run`` (solo las columnas de estado que ``etl_writer`` puede modificar).
Las tablas de registro (``etl_run``, ``correccion_dato``, ``excel_reference_run``, ``reconciliation_result``,
``exclusion_matrix_carga``, ``detencion_revision``) son solo-inserción.
"""

from __future__ import annotations

import contextlib
import json
import time
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import datetime
from typing import Literal

from sqlalchemy.engine import Engine

from etl_arena.config import ConfiguracionCalculo
from etl_arena.excel_semantics import serial_a_datetime
from etl_arena.model import FilaReconciliacion, KpiMensual, ReferenciaExcel
from etl_arena.persistence.paquete import (
    COLUMNAS_DETENCION,
    LOTE,
    MetadatosCorrida,
    PaqueteMes,
    Tabla,
    _dt,
    _f,
    fila_etl_run,
)

EstadoFinal = Literal["success", "failed", "parity_failed"]
TABLAS_REEMPLAZO_MENSUAL = ("muestra_pcs", "muestra_planta", "disponibilidad_diaria", "calidad_dato")
CAMPOS_CRUDOS_PCS = ("MODULES", "FAULT", "STATUS", "WARNING")  # orden de PaqueteMes.crudos_pcs[2:]


class ErrorPublicacion(RuntimeError):
    """La publicación del mes se detuvo por una protección (D-21, D-23) o un bloqueo; nada se escribió."""


@dataclass(frozen=True)
class OpcionesPublicacion:
    permitir_recorte: bool = False  # D-21
    reemplazar_manual: bool = False  # D-23: meses excel_manual (jul/ago 2026)
    forzar_sin_exclusiones: bool = False  # D-23: bajar un mes "Con Exclusiones" a "Sin Exclusiones"


@dataclass
class ResumenPublicacion:
    anio: int
    mes: int
    filas_borradas: dict[str, int] = field(default_factory=dict)
    filas_insertadas: dict[str, int] = field(default_factory=dict)
    detenciones: dict[str, int] = field(default_factory=dict)  # nuevas | actualizadas | eliminadas
    correcciones: int = 0
    segundos: float = 0.0

    @property
    def mes_texto(self) -> str:
        return f"{self.anio}-{self.mes:02d}"

    def a_dict(self) -> dict:
        return {
            "mes": self.mes_texto,
            "filas_borradas": self.filas_borradas,
            "filas_insertadas": self.filas_insertadas,
            "detenciones": self.detenciones,
            "correcciones": self.correcciones,
            "segundos": round(self.segundos, 1),
        }


def _texto_valor(v) -> str | None:
    if v is None:
        return None
    return repr(v) if isinstance(v, float) else str(v)[:255]


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

    def _insertar_tablas(self, tablas: list[Tabla]) -> None:
        conn = self._conexion()
        try:
            cur = conn.cursor()
            for t in tablas:
                self._insertar(cur, t)
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    @staticmethod
    def _filas_afectadas(cur) -> int:
        cur.execute("SELECT @@ROWCOUNT")  # no depende de SET NOCOUNT de la sesión reutilizada
        return int(cur.fetchone()[0])

    # ------------------------------------------------------------------ catálogo y estado
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
        """Meses vigentes del año en ``disponibilidad_mensual`` (corridas + ``excel_manual``), para el anual (R10)."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "SELECT Anio, Mes, DiasMes, BloquesMuestreo, BloquesRacksIndisponibles, Origen, NumCorrida, "
                "DisponibilidadContractual FROM dbo.disponibilidad_mensual WHERE IdProyecto = ? AND Anio = ?",
                id_proyecto,
                anio,
            )
            return [
                KpiMensual(int(a), int(m), float(d), float(b), float(f), o, str(n) if n else None, float(k))
                for a, m, d, b, f, o, n, k in cur.fetchall()
            ]
        finally:
            conn.close()

    def estado_mes(self, id_proyecto: int, anio: int, mes: int) -> dict | None:
        """Fila vigente del mes en ``disponibilidad_mensual`` (``None`` si nunca se publicó)."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "SELECT UltimoDato, MesCompleto, EstadoExclusiones, TipoCorrida, Origen, NumCorrida "
                "FROM dbo.disponibilidad_mensual WHERE IdProyecto = ? AND Anio = ? AND Mes = ?",
                id_proyecto,
                anio,
                mes,
            )
            f = cur.fetchone()
            if f is None:
                return None
            claves = ("UltimoDato", "MesCompleto", "EstadoExclusiones", "TipoCorrida", "Origen", "NumCorrida")
            return dict(zip(claves, f, strict=True))
        finally:
            conn.close()

    # ------------------------------------------------------------------ ciclo de vida de la corrida
    def iniciar(self, cfg: ConfiguracionCalculo, meta: MetadatosCorrida) -> int:
        """Registra la corrida en ``running`` y devuelve su ``NumCorrida`` (correlativo para personas)."""
        columnas, valores = fila_etl_run(cfg, meta)
        sql = (
            Tabla("etl_run", columnas, [valores])
            .sql_insert()
            .replace(" VALUES ", " OUTPUT INSERTED.NumCorrida VALUES ", 1)
        )
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(sql, valores)
            num = int(cur.fetchone()[0])
            conn.commit()
            return num
        finally:
            conn.close()

    def publicar_mes(
        self,
        id_corrida: str,
        num_corrida: int,
        paquete: PaqueteMes,
        opciones: OpcionesPublicacion | None = None,
        archivo_origen: str = "",
        sha256_origen: str = "0" * 64,
    ) -> ResumenPublicacion:
        """Reemplaza el estado vigente del mes en una transacción (R11.3–R11.6, D-20…D-25)."""
        op = opciones or OpcionesPublicacion()
        inicio = time.perf_counter()
        idp, anio, mes = paquete.id_proyecto, paquete.anio, paquete.mes
        resumen = ResumenPublicacion(anio, mes)
        conn = self._conexion()
        try:
            cur = conn.cursor()
            self._bloquear(cur, idp)
            self._proteger(cur, paquete, op)
            resumen.correcciones = self._registrar_correcciones(
                cur, id_corrida, paquete, archivo_origen[:260], sha256_origen
            )
            for nombre in TABLAS_REEMPLAZO_MENSUAL:
                cur.execute(f"DELETE FROM dbo.{nombre} WHERE IdProyecto = ? AND Anio = ? AND Mes = ?", idp, anio, mes)
                resumen.filas_borradas[nombre] = self._filas_afectadas(cur)
                tabla = paquete.tabla(nombre)
                con_num = Tabla(nombre, (*tabla.columnas, "NumCorrida"), [(*f, num_corrida) for f in tabla.filas])
                self._insertar(cur, con_num)
                resumen.filas_insertadas[nombre] = len(con_num.filas)
            resumen.detenciones = self._fusionar_detenciones(cur, paquete, num_corrida)
            self._reemplazar_mensual(cur, paquete, num_corrida)
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            with contextlib.suppress(Exception):  # NOCOUNT es de la sesión: no dejarlo puesto en el pool
                conn.cursor().execute("SET NOCOUNT OFF")
            conn.close()
        resumen.segundos = time.perf_counter() - inicio
        return resumen

    @staticmethod
    def _bloquear(cur, id_proyecto: int) -> None:
        """Una sola publicación a la vez por proyecto (dos PC o dos procesos sobre la misma base)."""
        cur.execute(
            "SET NOCOUNT ON; DECLARE @r INT; "
            "EXEC @r = sp_getapplock @Resource = ?, @LockMode = 'Exclusive', @LockOwner = 'Transaction', "
            "@LockTimeout = 120000; SELECT @r",
            f"etl_arena:proyecto:{id_proyecto}",
        )
        if int(cur.fetchone()[0]) < 0:
            raise ErrorPublicacion(f"otra carga está publicando el proyecto {id_proyecto}: reintentar en unos minutos")

    @staticmethod
    def _proteger(cur, paquete: PaqueteMes, op: OpcionesPublicacion) -> None:
        cur.execute(
            "SELECT UltimoDato, EstadoExclusiones, Origen FROM dbo.disponibilidad_mensual WITH (UPDLOCK, HOLDLOCK) "
            "WHERE IdProyecto = ? AND Anio = ? AND Mes = ?",
            paquete.id_proyecto,
            paquete.anio,
            paquete.mes,
        )
        vigente = cur.fetchone()
        if vigente is None:
            return
        ultimo, exclusiones, origen = vigente
        mes = f"{paquete.anio}-{paquete.mes:02d}"
        if origen == "excel_manual" and not op.reemplazar_manual:
            raise ErrorPublicacion(
                f"{mes} es un valor oficial importado del Excel (excel_manual): se reemplaza solo con "
                "--reemplazar-manual (D-23, D-11)"
            )
        baja_a_sin = exclusiones == "con_exclusiones" and paquete.estado_exclusiones == "sin_exclusiones"
        if baja_a_sin and not op.forzar_sin_exclusiones:
            raise ErrorPublicacion(
                f"{mes} está publicado 'Con Exclusiones' y esta carga es 'Sin Exclusiones': se reemplaza solo con "
                "--forzar-sin-exclusiones (D-23)"
            )
        recorta = ultimo is not None and paquete.ultimo_dato is not None and paquete.ultimo_dato < ultimo
        if recorta and not op.permitir_recorte:
            raise ErrorPublicacion(
                f"{mes}: la carga llega hasta {paquete.ultimo_dato:%Y-%m-%d %H:%M} y lo vigente hasta "
                f"{ultimo:%Y-%m-%d %H:%M}; se borrarían datos: usar --permitir-recorte si es intencional (D-21)"
            )

    @staticmethod
    def _registrar_correcciones(cur, id_corrida: str, paquete: PaqueteMes, archivo: str, sha: str) -> int:
        """Cada dato crudo de SCADA (RawData-PCS) o de PlantActivity que cambió respecto de lo vigente (D-24).

        Solo se comparan los bloques que existen en ambas cargas; los de la ``Exclusion_Matrix`` los registra la
        etapa ``load-exclusion-matrix`` celda a celda (con sus comentarios)."""
        idp, anio, mes = paquete.id_proyecto, paquete.anio, paquete.mes
        filas: list[tuple] = []
        cur.execute(
            "SELECT SerialFecha, Ocurrencia, NumeroPCS, ModulosRaw, FallaRaw, EstadoRaw, AdvertenciaRaw "
            "FROM dbo.muestra_pcs WHERE IdProyecto = ? AND Anio = ? AND Mes = ?",
            idp,
            anio,
            mes,
        )
        for serial, occ, pcs, *antes in cur.fetchall():
            nuevo = paquete.crudos_pcs.get((float(serial), int(occ), int(pcs)))
            if nuevo is None:
                continue
            fila, marca, *despues = nuevo
            for campo, a, d in zip(CAMPOS_CRUDOS_PCS, antes, despues, strict=True):
                if a != d:
                    filas.append(
                        (
                            id_corrida,
                            "recarga",
                            "RawData-PCS",
                            fila,
                            float(serial),
                            marca or serial_a_datetime(float(serial)),
                            int(pcs),
                            campo,
                            _texto_valor(a),
                            _texto_valor(d),
                            archivo,
                            sha,
                        )
                    )
        cur.execute(
            "SELECT SerialFecha, Ocurrencia, FactorOperacionalRaw FROM dbo.muestra_planta "
            "WHERE IdProyecto = ? AND Anio = ? AND Mes = ?",
            idp,
            anio,
            mes,
        )
        for serial, occ, antes in cur.fetchall():
            nuevo = paquete.crudos_planta.get((float(serial), int(occ)))
            if nuevo is None or nuevo[2] == antes:
                continue
            fila, marca, despues = nuevo
            filas.append(
                (
                    id_corrida,
                    "plant_activity",
                    "PlantActivity",
                    fila,
                    float(serial),
                    marca or serial_a_datetime(float(serial)),
                    None,
                    "C",
                    _texto_valor(antes),
                    _texto_valor(despues),
                    archivo,
                    sha,
                )
            )
        tabla = Tabla("correccion_dato", ("IdCorrida", *RepositorioCorridas.COLUMNAS_CORRECCION), filas)
        RepositorioCorridas._insertar(cur, tabla)
        return len(filas)

    @staticmethod
    def _fusionar_detenciones(cur, paquete: PaqueteMes, num_corrida: int) -> dict[str, int]:
        """``MERGE`` por clave de negocio dentro del mes: conserva ``IdDetencion`` (D-25)."""
        idp, anio, mes = paquete.id_proyecto, paquete.anio, paquete.mes
        cur.execute(
            "SELECT IdDetencion, NumeroPCS, FechaInicio, Ocurrencia FROM dbo.detencion WITH (UPDLOCK, HOLDLOCK) "
            "WHERE IdProyecto = ? AND Anio = ? AND Mes = ?",
            idp,
            anio,
            mes,
        )
        existentes = {(int(p), _dt(f), int(o)): int(i) for i, p, f, o in cur.fetchall()}
        cols = COLUMNAS_DETENCION
        k_pcs, k_ini, k_occ = cols.index("NumeroPCS"), cols.index("FechaInicio"), cols.index("Ocurrencia")
        claves = {"IdProyecto", "Anio", "Mes", "NumeroPCS", "FechaInicio", "Ocurrencia"}
        mutables = [c for c in cols if c not in claves]
        idx_mut = [cols.index(c) for c in mutables]
        nuevas, actualizar, vistos = [], [], set()
        for fila in paquete.detenciones.filas:
            clave = (fila[k_pcs], fila[k_ini], fila[k_occ])
            if clave in existentes:
                vistos.add(existentes[clave])
                actualizar.append((*(fila[i] for i in idx_mut), num_corrida, existentes[clave]))
            else:
                nuevas.append((*fila, num_corrida))
        eliminar = [(i,) for i in existentes.values() if i not in vistos]
        cur.fast_executemany = True
        if actualizar:
            sets = ", ".join(f"{c} = ?" for c in (*mutables, "NumCorrida"))
            cur.executemany(f"UPDATE dbo.detencion SET {sets} WHERE IdDetencion = ?", actualizar)
        if eliminar:
            cur.executemany("DELETE FROM dbo.detencion WHERE IdDetencion = ?", eliminar)
        RepositorioCorridas._insertar(cur, Tabla("detencion", (*cols, "NumCorrida"), nuevas))
        return {"nuevas": len(nuevas), "actualizadas": len(actualizar), "eliminadas": len(eliminar)}

    @staticmethod
    def _reemplazar_mensual(cur, paquete: PaqueteMes, num_corrida: int) -> None:
        cur.execute(
            "DELETE FROM dbo.disponibilidad_mensual WHERE IdProyecto = ? AND Anio = ? AND Mes = ?",
            paquete.id_proyecto,
            paquete.anio,
            paquete.mes,
        )
        datos = {
            "IdProyecto": paquete.id_proyecto,
            "Anio": paquete.anio,
            "Mes": paquete.mes,
            **paquete.mensual,
            "NumCorrida": num_corrida,
            "ActualizadoEn": datetime.now().replace(microsecond=0),
        }
        RepositorioCorridas._insertar(cur, Tabla("disponibilidad_mensual", tuple(datos), [tuple(datos.values())]))

    def finalizar(
        self,
        id_corrida: str,
        estado: EstadoFinal,
        resumen_calidad: dict | None = None,
        mensaje_error: str | None = None,
        minutos_muestreo_derivado: float | None = None,
        id_referencia_excel: str | None = None,
        publicada: bool = False,
        meses_publicados: Iterable[str] = (),
    ) -> None:
        """Transición de estado del propio ``etl_run`` (transacción aparte, R11.7)."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "UPDATE dbo.etl_run SET Estado = ?, FinalizadoEn = ?, ResumenCalidad = ?, MensajeError = ?, "
                "MinutosMuestreoDerivado = COALESCE(?, MinutosMuestreoDerivado), "
                "IdReferenciaExcel = COALESCE(?, IdReferenciaExcel), Publicada = ?, MesesPublicados = ? "
                "WHERE IdCorrida = ? AND Estado = 'running'",
                estado,
                datetime.now().replace(microsecond=0),
                json.dumps(resumen_calidad, ensure_ascii=False, default=str) if resumen_calidad is not None else None,
                mensaje_error,
                _f(minutos_muestreo_derivado),
                id_referencia_excel,
                publicada,
                ", ".join(meses_publicados) or None,
                id_corrida,
            )
            if self._filas_afectadas(cur) != 1:
                raise RuntimeError(f"etl_run {id_corrida} no existe o ya no está en 'running'")
            conn.commit()
        finally:
            conn.close()

    # ------------------------------------------------------------------ referencia Excel y reconciliación
    def guardar_referencia_excel(self, ref: ReferenciaExcel) -> str:
        """Solo parámetros y KPI; el detalle queda en ``referencia_excel.json`` del corte (D-24)."""
        tabla = Tabla(
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
                    ref.id_referencia,
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
        )
        self._insertar_tablas([tabla])
        return ref.id_referencia

    def guardar_reconciliacion(
        self, id_corrida: str, id_referencia: str | None, filas: Iterable[FilaReconciliacion]
    ) -> int:
        """Resumen por nivel + diferencias fuera de tolerancia (``reconciliation.a_filas``)."""
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
        self._insertar_tablas([tabla])
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

    COLUMNAS_CORRECCION = (
        "TipoCorreccion",
        "Hoja",
        "NumeroFilaOrigen",
        "SerialFechaExcelOrigen",
        "MarcaTiempoLocalOrigen",
        "NumeroPCS",
        "Campo",
        "ValorAnterior",
        "ValorNuevo",
        "ArchivoOrigen",
        "Sha256Archivo",
    )

    def guardar_correcciones(self, id_corrida: str, filas: Iterable[tuple]) -> int:
        """Celdas cambiadas por una carga de la matriz (D-17), ligadas a la corrida que las aplicó.

        Cada fila sigue ``COLUMNAS_CORRECCION``; los valores se guardan como texto (≤ 255). Una sola
        transacción; devuelve cuántas filas insertó."""
        tabla = Tabla("correccion_dato", ("IdCorrida", *self.COLUMNAS_CORRECCION))
        for f in filas:
            *inicio, anterior, nuevo, archivo, sha = f
            texto = [None if v is None else str(v)[:255] for v in (anterior, nuevo)]
            tabla.filas.append((id_corrida, *inicio, *texto, str(archivo)[:260], sha))
        self._insertar_tablas([tabla])
        return len(tabla.filas)

    def vincular_carga_con_cierre(self, id_carga: int, id_corrida: str) -> bool:
        """Anota en ``exclusion_matrix_carga`` qué cierre mensual usó la entrega (``IdCorridaCierre``).

        Única actualización permitida sobre la tabla (06_roles): solo si aún estaba vacía; devuelve si la anotó."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "UPDATE dbo.exclusion_matrix_carga SET IdCorridaCierre = ? "
                "WHERE IdCarga = ? AND IdCorridaCierre IS NULL",
                id_corrida,
                id_carga,
            )
            n = self._filas_afectadas(cur)
            conn.commit()
            return n == 1
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
        """¿El mes ya tiene cierre publicado (``cierre_mensual`` del mes completo) o KPI ``excel_manual``?
        Si no, el orquestador lo encola (R19.3)."""
        conn = self._conexion()
        try:
            cur = conn.cursor()
            cur.execute(
                "SELECT CASE WHEN EXISTS (SELECT 1 FROM dbo.disponibilidad_mensual WHERE IdProyecto = ? AND Anio = ? "
                "AND Mes = ? AND ((MesCompleto = 1 AND TipoCorrida = N'cierre_mensual') OR Origen = N'excel_manual')) "
                "THEN 1 ELSE 0 END",
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


RepositorioEstado = RepositorioCorridas  # nombre del diseño rev. 3

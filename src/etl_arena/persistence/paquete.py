"""Filas por tabla de una corrida, listas para insertar (R11, R12.4, R15; design.md §Data Models).

Funciones puras (sin base de datos): convierten ``ResultadoCalculo`` en tuplas en el orden de
columnas del DDL. ``NaN`` → ``NULL`` (SQL Server no admite NaN en FLOAT); marcas de tiempo al
segundo (``DATETIME``).
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Literal

import numpy as np

from etl_arena.aggregation.mensual_anual import FilaAnual
from etl_arena.excel_semantics import datetime_a_serial, serial_a_datetime, texto_excel
from etl_arena.model import Anomalia, KpiMensual
from etl_arena.normalization.esquema import ORDEN_CAMPOS, columna
from etl_arena.pipeline import ResultadoCalculo

EstadoExclusiones = Literal["sin_exclusiones", "con_exclusiones"]
LOTE = 50_000


@dataclass(frozen=True)
class MetadatosCorrida:
    """Datos de la corrida que no salen del cálculo (R11.4)."""

    hash_archivo_origen: str
    estado_exclusiones: EstadoExclusiones = "sin_exclusiones"
    sistema_origen: str = "scada_export"
    filas_nuevas_export: frozenset[int] = frozenset()  # filas aportadas por el export incremental (D-07)
    iniciado_en: datetime = field(default_factory=lambda: datetime.now().replace(microsecond=0))


@dataclass
class Tabla:
    nombre: str
    columnas: tuple[str, ...]
    filas: list[tuple] = field(default_factory=list)

    def sql_insert(self) -> str:
        cols = ", ".join(self.columnas)
        marcas = ", ".join("?" for _ in self.columnas)
        return f"INSERT INTO dbo.{self.nombre} ({cols}) VALUES ({marcas})"


@dataclass
class PaqueteCorrida:
    id_corrida: str
    tablas: list[Tabla]
    anomalias_persistencia: list[Anomalia] = field(default_factory=list)

    def tabla(self, nombre: str) -> Tabla:
        return next(t for t in self.tablas if t.nombre == nombre)

    def total_filas(self) -> int:
        return sum(len(t.filas) for t in self.tablas)


# ------------------------------------------------------------------ conversiones
def _f(x) -> float | None:
    if x is None:
        return None
    x = float(x)
    return None if math.isnan(x) else x


def _dt(v) -> datetime | None:
    if v is None:
        return None
    if isinstance(v, np.datetime64):
        if np.isnat(v):
            return None
        v = v.astype("datetime64[s]").astype(datetime)
    if isinstance(v, datetime):
        return v.replace(microsecond=0)
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    raise TypeError(f"fecha no soportada: {v!r}")


def _txt(v, largo: int = 255) -> str | None:
    if v is None:
        return None
    s = texto_excel(v)
    return s[:largo]


def _dt_serial(s: float) -> datetime | None:
    return None if s == 0.0 or math.isnan(s) else serial_a_datetime(float(s))


# ------------------------------------------------------------------ selección de filas de staging
def filas_staging(r: ResultadoCalculo) -> np.ndarray:
    """Índices de la matriz a persistir en raw/plant_activity/exclusion (D-09): filas del período
    KPI o de eventos, más la fila anterior (la usa el fallback de descripción, F-04)."""
    cfg, s = r.cfg, r.matriz.serial
    ini = min(datetime_a_serial(cfg.inicio_periodo), datetime_a_serial(cfg.inicio_periodo_eventos))
    fin1 = max(datetime_a_serial(cfg.fin_periodo), datetime_a_serial(cfg.fin_periodo_eventos)) + 1
    idx = np.flatnonzero((s >= ini) & (s < fin1))
    if len(idx) and idx[0] > 0:
        idx = np.concatenate(([idx[0] - 1], idx))
    return idx


# ------------------------------------------------------------------ constructores por tabla
def construir_paquete(
    r: ResultadoCalculo,
    meta: MetadatosCorrida,
    catalogo: dict[str, int],
    kpi_mensual: KpiMensual | None = None,
    anual: list[FilaAnual] | None = None,
) -> PaqueteCorrida:
    """Todas las filas de la corrida excepto ``etl_run`` (que se inserta al iniciar)."""
    cfg, m, disp, ev = r.cfg, r.matriz, r.disponibilidad, r.eventos
    idc = cfg.id_corrida
    p = cfg.total_pcs
    anomalias_extra: list[Anomalia] = []

    colmap = Tabla("raw_pcs_column_map", ("IdCorrida", "NumeroPCS", "Campo", "NumeroColumna", "NombreColumna"))
    for pcs in m.pcs[:p]:
        for campo in ORDEN_CAMPOS:
            col = columna(pcs, campo)
            colmap.filas.append((idc, pcs, campo, col, _txt(r.libro.encabezado.get(col, ""), 200) or ""))

    sel = filas_staging(r)
    modulos_disp = m.modulos_disponibles(cfg.baterias_por_pcs)
    raw = Tabla(
        "raw_pcs_sample",
        (
            "IdCorrida",
            "NumeroFilaOrigen",
            "NumeroPCS",
            "SerialFechaExcelOrigen",
            "MarcaTiempoLocalOrigen",
            "FallaRaw",
            "FallaRawEsNumero",
            "EstadoRaw",
            "AdvertenciaRaw",
            "ModulosRaw",
            "ModulosDisponibles",
            "ModulosDisponiblesNulo",
            "EsFilaNuevaDelExport",
        ),
    )
    for i in sel.tolist():
        fila = int(m.numero_fila[i])
        marca = _dt(m.marca_tiempo[i])
        nueva = fila in meta.filas_nuevas_export
        for j in range(p):
            falla = m.falla[i, j]
            raw.filas.append(
                (
                    idc,
                    fila,
                    m.pcs[j],
                    float(m.serial[i]),
                    marca,
                    _txt(falla),
                    isinstance(falla, float),
                    _txt(m.estado[i, j]),
                    _txt(m.advertencia[i, j]),
                    _f(m.modulos[i, j]),
                    float(modulos_disp[i, j]),
                    bool(m.modulos_nulo[i, j]),
                    nueva,
                )
            )

    pa = Tabla(
        "plant_activity_sample",
        (
            "IdCorrida",
            "NumeroFilaOrigen",
            "SerialFecha",
            "MarcaTiempoMuestra",
            "FactorOperacionalRaw",
            "EventoExcusadoRaw",
            "SetpointPotenciaActivaKW",
            "DroopSobrefrecuenciaHabilitado",
            "DroopBajafrecuenciaHabilitado",
            "PotenciaActivaPOIKW",
            "PorcentajeSOC",
        ),
    )
    for i in sel.tolist():
        fila = int(m.numero_fila[i])
        c = r.libro.actividad.get(fila, {})
        num = [c.get(k) if isinstance(c.get(k), float) else None for k in range(2, 10)]
        pa.filas.append((idc, fila, num[0], _dt_serial(num[0]) if num[0] is not None else None, *num[1:]))

    exc = r.exclusion
    ems = Tabla(
        "exclusion_matrix_sample",
        (
            "IdCorrida",
            "NumeroFilaOrigen",
            "NumeroPCS",
            "SerialFecha",
            "MarcaTiempoMuestra",
            "ValorExclusion",
            "BateriasPrevias",
            "EventoExcusadoFila",
            "Comentario",
        ),
    )
    for i in sel.tolist():
        for j in np.flatnonzero(exc.valor[i, :p]).tolist():
            s_em = _f(exc.serial_em[i])
            ems.filas.append(
                (
                    idc,
                    int(m.numero_fila[i]),
                    m.pcs[j],
                    s_em,
                    _dt_serial(s_em) if s_em is not None else None,
                    int(exc.valor[i, j]),
                    _f(exc.baterias_previas[i, j]),
                    _f(exc.evento_excusado[i]),
                    _txt(exc.comentario[i], 500),
                )
            )

    asr = Tabla(
        "availability_sample_result",
        (
            "IdCorrida",
            "NumeroFilaOrigen",
            "NumeroPCS",
            "SerialFecha",
            "MarcaTiempoMuestra",
            "ModulosDisponibles",
            "ModulosDisponiblesNulo",
            "BateriasIndisponibles",
            "ValorExclusion",
            "FactorOperacional",
            "BateriasIndisponiblesPonderadas",
            "ImpactoRackPonderado",
        ),
    )
    for k, i in enumerate(disp.filas_procesadas.tolist()):
        marca = _dt(m.marca_tiempo[i])
        fila = int(disp.numero_fila[k])
        for j in range(p):
            b = disp.baterias[k, j]
            asr.filas.append(
                (
                    idc,
                    fila,
                    disp.pcs[j],
                    float(disp.serial[k]),
                    marca,
                    float(modulos_disp[i, j]),
                    bool(m.modulos_nulo[i, j]),
                    0.0 if np.isnan(b) else float(b),
                    int(disp.exclusion[k, j]),
                    float(disp.factor_operacional[k]),
                    0.0 if np.isnan(b) else float(disp.ponderadas[k, j]),
                    float(disp.impacto[k, j]),
                )
            )

    c23 = disp.minutos_muestreo_derivado
    dias_periodo = (cfg.fin_periodo - cfg.inicio_periodo).days + 1
    arr = Tabla(
        "availability_run_result",
        (
            "IdCorrida",
            "BloquesMuestreo",
            "TotalRacks",
            "BloquesRacksIndisponibles",
            "DisponibilidadPeriodo",
            "DisponibilidadAnualAcumulada",
            "MinutosMuestreoDerivado",
            "HorasRackEventos",
            "BloquesMuestreoCalendario",
        ),
        [
            (
                idc,
                disp.bloques_muestreo,
                cfg.total_racks,
                float(disp.bloques_racks_indisponibles),
                _f(disp.disponibilidad_periodo),
                _f(disp.disponibilidad_anual_acumulada),
                _f(c23),
                float(ev.horas_rack_totales),
                dias_periodo * 24 * 60 / c23 if c23 else None,
            )
        ],
    )

    eventos = ev.eventos()
    fe = Tabla(
        "fault_event",
        (
            "IdCorrida",
            "OrdenExcel",
            "NumeroPCS",
            "SerialInicio",
            "SerialFin",
            "MarcaTiempoInicio",
            "MarcaTiempoFin",
            "DuracionHoras",
            "CodigoFalla",
            "DescripcionFalla",
            "DescripcionFallaFallback",
            "NumeroBloques",
            "SumaBloques",
            "PromedioBateriasInvolucradas",
            "HorasRackIndisponibles",
            "EventoArrastradoExcel",
            "ExcelHabriaFallado",
            "CerradoPorModulosNulo",
            "TieneExclusion",
        ),
    )
    det = Tabla(
        "detencion",
        (
            "IdProyecto",
            "IdCorrida",
            "OrdenExcel",
            "NumeroPCS",
            "FechaInicio",
            "FechaTermino",
            "DuracionSegundos",
            "DuracionHoras",
            "IdTipoDetencion",
            "CodigoFalla",
            "DescripcionFalla",
            "DescripcionFallaFallback",
            "PromedioBateriasInvolucradas",
            "HorasRackIndisponibles",
            "CerradoPorModulosNulo",
            "EsExcusable",
            "EventoArrastradoExcel",
        ),
    )
    for e in eventos:
        fe.filas.append(
            (
                idc,
                e.orden_excel,
                e.numero_pcs,
                e.serial_inicio,
                e.serial_fin,
                _dt(e.marca_tiempo_inicio),
                _dt(e.marca_tiempo_fin),
                e.duracion_horas,
                e.codigo_falla[:300],
                e.descripcion_falla[:255],
                e.descripcion_falla_fallback,
                e.numero_bloques,
                e.suma_bloques,
                e.promedio_baterias,
                e.horas_rack_indisponibles,
                e.arrastrado_excel,
                e.excel_habria_fallado,
                e.cerrado_por_nulo,
                e.tiene_exclusion,
            )
        )
        id_tipo = catalogo.get(e.codigo_falla)
        if id_tipo is None:
            anomalias_extra.append(
                Anomalia(
                    "codigo_sin_catalogo",
                    "advertencia",
                    numero_pcs=e.numero_pcs,
                    serial=e.serial_inicio,
                    detalle=f"evento {e.orden_excel}: código {e.codigo_falla!r} no está en tipo_detencion",
                )
            )
        det.filas.append(
            (
                cfg.id_proyecto,
                idc,
                e.orden_excel,
                e.numero_pcs,
                _dt(e.marca_tiempo_inicio),
                _dt(e.marca_tiempo_fin),
                round(e.duracion_horas * 3600),
                e.duracion_horas,  # mismo valor que fault_event.DuracionHoras (ListOfFaults!E)
                id_tipo,
                e.codigo_falla[:300],
                e.descripcion_falla[:255],
                e.descripcion_falla_fallback,
                e.promedio_baterias,
                e.horas_rack_indisponibles,
                e.cerrado_por_nulo,
                e.tiene_exclusion,
                e.arrastrado_excel,
            )
        )

    fcs = Tabla(
        "fault_code_summary",
        ("IdCorrida", "Ranking", "CodigoFalla", "DescripcionFallaPE", "HorasRackIndisponibles", "Porcentaje"),
        [
            (idc, k, f.codigo[:10], f.descripcion_pe[:150] or None, f.horas_rack, _f(f.porcentaje))
            for k, f in enumerate(ev.resumen, start=1)
        ],
    )

    daily = Tabla(
        "daily_availability",
        (
            "IdCorrida",
            "DiaN",
            "Dia",
            "BloquesRacksIndisponiblesDiarios",
            "BloquesRacksIndisponiblesAcumulados",
            "Disponibilidad",
            "Variacion",
        ),
        [
            (idc, d.numero_dia, _dt(d.dia), d.diario, d.acumulado, _f(d.disponibilidad), _f(d.variacion))
            for d in r.diario
        ],
    )

    mensual = Tabla(
        "monthly_official_kpi",
        (
            "IdProyecto",
            "Anio",
            "Mes",
            "DiasMes",
            "BloquesMuestreo",
            "BloquesRacksIndisponibles",
            "DisponibilidadContractual",
            "Origen",
            "IdCorrida",
        ),
    )
    if kpi_mensual is not None:
        k = kpi_mensual
        mensual.filas.append(
            (
                cfg.id_proyecto,
                k.anio,
                k.mes,
                k.dias_mes,
                k.bloques_muestreo,
                k.bloques_racks_indisponibles,
                k.disponibilidad_contractual,
                k.origen,
                idc,
            )
        )

    anual_t = Tabla(
        "annual_availability",
        (
            "IdCorrida",
            "Anio",
            "Mes",
            "DiasMes",
            "BloquesMuestreo",
            "BloquesRacksIndisponibles",
            "DisponibilidadMensual",
            "BloquesMuestreoAcumulados",
            "BloquesIndisponiblesAcumulados",
            "DisponibilidadAcumulada",
            "DisponibilidadContractual",
        ),
        [
            (
                idc,
                a.anio,
                a.mes,
                a.dias_mes,
                a.bloques_muestreo,
                _f(a.bloques_racks_indisponibles),
                _f(a.disponibilidad_mensual),
                a.bloques_muestreo_acumulados,
                a.bloques_indisponibles_acumulados,
                _f(a.disponibilidad_acumulada),
                _f(a.disponibilidad_contractual),
            )
            for a in (anual or [])
        ],
    )

    dqi = Tabla(
        "data_quality_issue",
        ("IdCorrida", "Tipo", "Severidad", "NumeroFilaOrigen", "NumeroPCS", "SerialFecha", "Detalle"),
    )
    for a in [*r.anomalias, *anomalias_extra]:
        dqi.filas.append(
            (idc, a.tipo[:50], a.severidad, a.numero_fila, a.numero_pcs, _f(a.serial), (a.detalle or "")[:1000] or None)
        )

    tablas = [colmap, raw, pa, ems, asr, arr, fe, fcs, daily, mensual, anual_t, det, dqi]
    return PaqueteCorrida(idc, tablas, anomalias_extra)


def fila_etl_run(r_cfg, meta: MetadatosCorrida) -> tuple[tuple[str, ...], tuple]:
    """Columnas y valores de ``etl_run`` al iniciar (Estado = running)."""
    c = r_cfg
    columnas = (
        "IdCorrida",
        "IdProyecto",
        "TipoCorrida",
        "EsOficial",
        "EstadoExclusiones",
        "ArchivoOrigen",
        "HashArchivoOrigen",
        "SistemaOrigen",
        "IniciadoEn",
        "InicioPeriodo",
        "FinPeriodo",
        "InicioPeriodoEventos",
        "FinPeriodoEventos",
        "FinDiario",
        "SoloTiempoOperacional",
        "AplicarEventoExcusable",
        "AplicarEventoExcusableEventos",
        "ModoHuecos",
        "TotalPCS",
        "BateriasPorPCS",
        "RacksPorPCS",
        "TotalRacks",
        "MinutosMuestreo",
        "Estado",
        "VersionAlgoritmo",
    )
    valores = (
        c.id_corrida,
        c.id_proyecto,
        c.tipo_corrida,
        c.es_oficial,
        meta.estado_exclusiones,
        c.archivo_origen[:500],
        meta.hash_archivo_origen,
        meta.sistema_origen,
        meta.iniciado_en,
        _dt(c.inicio_periodo),
        _dt(c.fin_periodo),
        _dt(c.inicio_periodo_eventos),
        _dt(c.fin_periodo_eventos),
        _dt(c.fin_diario),
        c.solo_tiempo_operacional,
        c.aplicar_evento_excusable,
        c.aplicar_evento_excusable_eventos,
        c.modo_huecos,
        c.total_pcs,
        c.baterias_por_pcs,
        c.racks_por_pcs,
        c.total_racks,
        c.minutos_muestreo,
        "running",
        c.version_algoritmo,
    )
    return columnas, valores

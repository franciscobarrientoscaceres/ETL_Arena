"""Filas de las tablas de estado de un mes, listas para publicar (R11 rev. 3, ADR-12; design.md §Revisión 3).

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

from etl_arena.excel_semantics import serial_a_datetime, texto_excel
from etl_arena.model import Anomalia, KpiMensual
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


# ------------------------------------------------------------------ paquete de un mes (estado vigente, ADR-12)
@dataclass
class PaqueteMes:
    """Todo lo que se publica de un mes: filas de las tablas de estado + crudos para detectar correcciones.

    Las filas no llevan ``NumCorrida`` (se agrega al publicar). ``detenciones`` va por clave de negocio
    ``(NumeroPCS, FechaInicio, Ocurrencia)`` dentro del mes."""

    id_proyecto: int
    anio: int
    mes: int
    ultimo_dato: datetime | None
    mes_completo: bool
    estado_exclusiones: EstadoExclusiones
    muestras_pcs: Tabla
    muestras_planta: Tabla
    detenciones: Tabla
    diario: Tabla
    calidad: Tabla
    mensual: dict
    crudos_pcs: dict  # (SerialFecha, Ocurrencia, PCS) -> (fila, marca, MODULES, FAULT, STATUS, WARNING)
    crudos_planta: dict  # (SerialFecha, Ocurrencia) -> (fila, marca, C)
    anomalias_persistencia: list[Anomalia] = field(default_factory=list)

    @property
    def tablas(self) -> list[Tabla]:
        return [self.muestras_pcs, self.muestras_planta, self.detenciones, self.diario, self.calidad]

    def tabla(self, nombre: str) -> Tabla:
        return next(t for t in self.tablas if t.nombre == nombre)


COLUMNAS_MUESTRA_PCS = (
    "IdProyecto",
    "SerialFecha",
    "Ocurrencia",
    "NumeroPCS",
    "Anio",
    "Mes",
    "MarcaTiempo",
    "NumeroFilaOrigen",
    "FallaRaw",
    "FallaRawEsNumero",
    "EstadoRaw",
    "AdvertenciaRaw",
    "ModulosRaw",
    "ModulosDisponibles",
    "ModulosDisponiblesNulo",
    "ValorExclusion",
    "BateriasPrevias",
    "BateriasIndisponibles",
    "FactorOperacional",
    "BateriasIndisponiblesPonderadas",
    "ImpactoRackPonderado",
    "EsFilaNuevaDelExport",
)
COLUMNAS_MUESTRA_PLANTA = (
    "IdProyecto",
    "SerialFecha",
    "Ocurrencia",
    "Anio",
    "Mes",
    "MarcaTiempo",
    "NumeroFilaOrigen",
    "FactorOperacionalRaw",
    "EventoExcusadoRaw",
    "SetpointPotenciaActivaKW",
    "DroopSobrefrecuenciaHabilitado",
    "DroopBajafrecuenciaHabilitado",
    "PotenciaActivaPOIKW",
    "PorcentajeSOC",
    "EventoExcusadoFila",
    "ComentarioExclusion",
)
COLUMNAS_DETENCION = (
    "IdProyecto",
    "Anio",
    "Mes",
    "NumeroPCS",
    "FechaInicio",
    "Ocurrencia",
    "FechaTermino",
    "DuracionSegundos",
    "DuracionHoras",
    "IdTipoDetencion",
    "CodigoFalla",
    "DescripcionFalla",
    "DescripcionFallaFallback",
    "SerialInicio",
    "SerialFin",
    "NumeroBloques",
    "SumaBloques",
    "PromedioBateriasInvolucradas",
    "HorasRackIndisponibles",
    "EventoArrastradoExcel",
    "ExcelHabriaFallado",
    "CerradoPorModulosNulo",
    "EsExcusable",
    "OrdenExcel",
)
COLUMNAS_DIARIO = (
    "IdProyecto",
    "Dia",
    "Anio",
    "Mes",
    "DiaN",
    "BloquesRacksIndisponiblesDiarios",
    "BloquesRacksIndisponiblesAcumulados",
    "Disponibilidad",
    "Variacion",
    "EstadoExclusiones",
)
COLUMNAS_CALIDAD = (
    "IdProyecto",
    "Anio",
    "Mes",
    "Tipo",
    "Severidad",
    "NumeroFilaOrigen",
    "NumeroPCS",
    "SerialFecha",
    "Detalle",
)


class ErrorPeriodoMensual(ValueError):
    """El período no es publicable: debe empezar el día 1 y no salir del mes (D-20)."""


def validar_periodo_mensual(cfg) -> None:
    ini, fin = cfg.inicio_periodo, cfg.fin_periodo
    if ini.day != 1 or (fin.year, fin.month) != (ini.year, ini.month) or fin < ini:
        raise ErrorPeriodoMensual(
            f"período {ini}..{fin}: para publicar debe empezar el día 1 y terminar dentro del mismo mes (D-20)"
        )


def _ocurrencias(claves) -> list[int]:
    """1, 2… para claves repetidas en orden (timestamps repetidos por el cambio de hora, F-32)."""
    vistas: dict = {}
    salida = []
    for c in claves:
        vistas[c] = vistas.get(c, 0) + 1
        salida.append(vistas[c])
    return salida


def construir_paquete_mes(
    r: ResultadoCalculo,
    meta: MetadatosCorrida,
    catalogo: dict[str, int],
    kpi_mensual: KpiMensual | None = None,
) -> PaqueteMes:
    """Filas de las tablas de estado para el mes del período (que debe ser mensual, D-20)."""
    from calendar import monthrange

    from etl_arena.orquestacion import mes_cubierto

    cfg, m, disp, ev = r.cfg, r.matriz, r.disponibilidad, r.eventos
    validar_periodo_mensual(cfg)
    p, idp = cfg.total_pcs, cfg.id_proyecto
    anio, mes = cfg.inicio_periodo.year, cfg.inicio_periodo.month
    anomalias_extra: list[Anomalia] = []
    modulos_disp = m.modulos_disponibles(cfg.baterias_por_pcs)
    exc = r.exclusion

    filas = disp.filas_procesadas.tolist()
    occ = _ocurrencias(float(disp.serial[k]) for k in range(len(filas)))
    mp = Tabla("muestra_pcs", COLUMNAS_MUESTRA_PCS)
    mpl = Tabla("muestra_planta", COLUMNAS_MUESTRA_PLANTA)
    crudos_pcs: dict = {}
    crudos_planta: dict = {}
    for k, i in enumerate(filas):
        serial, o = float(disp.serial[k]), occ[k]
        fila = int(m.numero_fila[i])
        marca = _dt(m.marca_tiempo[i])
        nueva = fila in meta.filas_nuevas_export
        for j in range(p):
            falla = m.falla[i, j]
            b = disp.baterias[k, j]
            crudo = (_f(m.modulos[i, j]), _txt(falla), _txt(m.estado[i, j]), _txt(m.advertencia[i, j]))
            crudos_pcs[(serial, o, m.pcs[j])] = (fila, marca, *crudo)
            mp.filas.append(
                (
                    idp,
                    serial,
                    o,
                    m.pcs[j],
                    anio,
                    mes,
                    marca,
                    fila,
                    crudo[1],
                    isinstance(falla, float),
                    crudo[2],
                    crudo[3],
                    crudo[0],
                    float(modulos_disp[i, j]),
                    bool(m.modulos_nulo[i, j]),
                    int(disp.exclusion[k, j]),
                    _f(exc.baterias_previas[i, j]),
                    0.0 if np.isnan(b) else float(b),
                    float(disp.factor_operacional[k]),
                    0.0 if np.isnan(b) else float(disp.ponderadas[k, j]),
                    float(disp.impacto[k, j]),
                    nueva,
                )
            )
        c = r.libro.actividad.get(fila, {})
        num = [c.get(col) if isinstance(c.get(col), float) else None for col in range(3, 10)]
        crudos_planta[(serial, o)] = (fila, marca, num[0])
        mpl.filas.append(
            (idp, serial, o, anio, mes, marca, fila, *num, _f(exc.evento_excusado[i]), _txt(exc.comentario[i], 500))
        )

    eventos = ev.eventos()
    occ_ev = _ocurrencias((e.numero_pcs, _dt(e.marca_tiempo_inicio)) for e in eventos)
    det = Tabla("detencion", COLUMNAS_DETENCION)
    for e, o in zip(eventos, occ_ev, strict=True):
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
                idp,
                anio,
                mes,
                e.numero_pcs,
                _dt(e.marca_tiempo_inicio),
                o,
                _dt(e.marca_tiempo_fin),
                round(e.duracion_horas * 3600),
                e.duracion_horas,  # ListOfFaults!E
                id_tipo,
                e.codigo_falla[:300],
                e.descripcion_falla[:255],
                e.descripcion_falla_fallback,
                e.serial_inicio,
                e.serial_fin,
                e.numero_bloques,
                e.suma_bloques,
                e.promedio_baterias,
                e.horas_rack_indisponibles,
                e.arrastrado_excel,
                e.excel_habria_fallado,
                e.cerrado_por_nulo,
                e.tiene_exclusion,
                e.orden_excel,
            )
        )

    diario = Tabla(
        "disponibilidad_diaria",
        COLUMNAS_DIARIO,
        [
            (
                idp,
                _dt(d.dia),
                d.dia.year,
                d.dia.month,
                d.numero_dia,
                d.diario,
                d.acumulado,
                _f(d.disponibilidad),
                _f(d.variacion),
                meta.estado_exclusiones,
            )
            for d in r.diario
        ],
    )

    calidad = Tabla("calidad_dato", COLUMNAS_CALIDAD)
    for a in [*r.anomalias, *anomalias_extra]:
        calidad.filas.append(
            (
                idp,
                anio,
                mes,
                a.tipo[:50],
                a.severidad,
                a.numero_fila,
                a.numero_pcs,
                _f(a.serial),
                (a.detalle or "")[:1000] or None,
            )
        )

    c23 = disp.minutos_muestreo_derivado
    dias_periodo = (cfg.fin_periodo - cfg.inicio_periodo).days + 1
    ultimo_serial = float(disp.serial[-1]) if filas else None
    completo = bool(
        cfg.fin_periodo.day == monthrange(anio, mes)[1]
        and ultimo_serial is not None
        and mes_cubierto(ultimo_serial, anio, mes, cfg.minutos_muestreo)
    )
    mensual = {
        "InicioPeriodo": _dt(cfg.inicio_periodo),
        "UltimoDato": _dt_serial(ultimo_serial) if ultimo_serial is not None else None,
        "MesCompleto": completo,
        "DiasMes": kpi_mensual.dias_mes if kpi_mensual is not None else float(dias_periodo),
        "BloquesMuestreo": float(disp.bloques_muestreo),
        "BloquesMuestreoCalendario": dias_periodo * 24 * 60 / c23 if c23 else None,
        "TotalRacks": cfg.total_racks,
        "BloquesRacksIndisponibles": float(disp.bloques_racks_indisponibles),
        "DisponibilidadMensual": _f(disp.disponibilidad_periodo),
        "DisponibilidadAnualAcumulada": _f(disp.disponibilidad_anual_acumulada),
        "HorasRackEventos": float(ev.horas_rack_totales),
        "MinutosMuestreoDerivado": _f(c23),
        "DisponibilidadContractual": kpi_mensual.disponibilidad_contractual if kpi_mensual is not None else 0.98,
        "AcumulaEnAnual": mes >= cfg.mes_inicio_acumulado(anio),
        "EstadoExclusiones": meta.estado_exclusiones,
        "TipoCorrida": cfg.tipo_corrida,
        "Origen": "corrida",
        "VersionAlgoritmo": cfg.version_algoritmo,
    }
    return PaqueteMes(
        idp,
        anio,
        mes,
        mensual["UltimoDato"],
        completo,
        meta.estado_exclusiones,
        mp,
        mpl,
        det,
        diario,
        calidad,
        mensual,
        crudos_pcs,
        crudos_planta,
        anomalias_extra,
    )


def fila_etl_run(r_cfg, meta: MetadatosCorrida) -> tuple[tuple[str, ...], tuple]:
    """Columnas y valores de ``etl_run`` al iniciar (Estado = running)."""
    c = r_cfg
    columnas = (
        "IdCorrida",
        "IdProyecto",
        "TipoCorrida",
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

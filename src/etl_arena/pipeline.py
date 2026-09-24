"""Cálculo completo sobre un libro de trabajo, sin persistencia (fase [E] de design.md).

La tarea 3.1 agrega persistencia SQL y reporte de calidad sobre este resultado.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from etl_arena.aggregation import calcular_diaria
from etl_arena.availability import ResultadoDisponibilidad, calcular
from etl_arena.config import ConfiguracionCalculo
from etl_arena.enrichment import asociar_actividad
from etl_arena.fault_events import ResultadoEventos, detectar_eventos
from etl_arena.ingestion import LibroCrudo, leer_libro
from etl_arena.ingestion.catalogos import leer_codigos_resumen
from etl_arena.model import Anomalia, DatosActividad, DiaDisponibilidad, ErrorParidad, MatrizPCS
from etl_arena.normalization import a_matriz, validar_esquema


@dataclass
class ResultadoCalculo:
    cfg: ConfiguracionCalculo
    libro: LibroCrudo
    matriz: MatrizPCS
    actividad: DatosActividad
    disponibilidad: ResultadoDisponibilidad
    eventos: ResultadoEventos
    diario: list[DiaDisponibilidad]
    anomalias: list[Anomalia] = field(default_factory=list)


def calcular_libro(
    ruta: str | Path, cfg: ConfiguracionCalculo, codigos_resumen: list[tuple[str, str]] | None = None
) -> ResultadoCalculo:
    """Ingesta → normalización → enriquecimiento → motores → Daily.

    ``codigos_resumen``: códigos de ``ListOfFaults!N`` para el resumen N:Q; si se omite, se
    leen del mismo libro.
    """
    libro = leer_libro(ruta, cfg)
    mapa, anomalias_esquema = validar_esquema(libro.encabezado, cfg)
    errores = [a for a in anomalias_esquema if a.severidad == "error"]
    if errores:
        raise ErrorParidad(f"encabezado de RawData-PCS inválido ({len(errores)} errores)", errores)
    m = a_matriz(libro, mapa, cfg)
    act, anomalias_pa = asociar_actividad(m, libro.actividad)
    disp = calcular(m, act, cfg)
    if codigos_resumen is None:
        codigos_resumen = leer_codigos_resumen(ruta)
    eventos = detectar_eventos(m, act, cfg, disp.minutos_muestreo_derivado, codigos_resumen)
    diario = calcular_diaria(disp, cfg)
    anomalias = [*libro.anomalias, *anomalias_esquema, *anomalias_pa, *eventos.anomalias]
    if disp.minutos_muestreo_derivado is not None and disp.minutos_muestreo_derivado != cfg.minutos_muestreo:
        anomalias.append(
            Anomalia(
                "c23_distinto",
                "advertencia",
                detalle=f"C23 derivado = {disp.minutos_muestreo_derivado} ≠ {cfg.minutos_muestreo} (F-13)",
            )
        )
    return ResultadoCalculo(cfg, libro, m, act, disp, eventos, diario, anomalias)

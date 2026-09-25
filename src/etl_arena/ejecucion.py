"""Corrida completa (fase [E] + [R] de design.md; tarea 3.1).

config → ingesta → normalización → enriquecimiento → disponibilidad → eventos → agregaciones →
calidad → persistencia → reconciliación. Con repositorio, el ``etl_run`` se registra antes de
calcular, así que toda falla queda en SQL con ``Estado = failed`` y su mensaje (R11.2–R11.3).
"""

from __future__ import annotations

import logging
import time
import traceback
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from etl_arena.aggregation import ErrorMensual, calcular_anual, registrar_mes_oficial
from etl_arena.aggregation.mensual_anual import FilaAnual
from etl_arena.config import ConfiguracionCalculo
from etl_arena.model import KpiMensual, ReferenciaExcel
from etl_arena.pipeline import ResultadoCalculo, calcular_libro
from etl_arena.reconciliation import ReporteReconciliacion, reconciliar
from etl_arena.reconciliation.servicio import a_filas, resumen
from etl_arena.reporting import generar_resumen

log = logging.getLogger("etl_arena.corrida")


@dataclass
class ResultadoEjecucion:
    id_corrida: str
    estado: str  # success | parity_failed | failed
    resultado: ResultadoCalculo | None = None
    reconciliacion: ReporteReconciliacion | None = None
    resumen_calidad: dict = field(default_factory=dict)
    kpi_mensual: KpiMensual | None = None
    anual: list[FilaAnual] = field(default_factory=list)
    filas_guardadas: dict[str, int] = field(default_factory=dict)
    segundos: dict[str, float] = field(default_factory=dict)
    error: str | None = None


def _kpi_mensual(r: ResultadoCalculo) -> KpiMensual | None:
    """Solo corridas oficiales cuyo período es un mes desde el día 1 (R10.2, D-07)."""
    if not r.cfg.es_oficial:
        return None
    try:
        return registrar_mes_oficial(r.disponibilidad, r.cfg)
    except ErrorMensual:
        return None


def _anual(cfg: ConfiguracionCalculo, mes: KpiMensual, repo) -> list[FilaAnual]:
    """Annual_AVA con los meses vigentes en SQL, reemplazando el de esta corrida (R10.3, D-06)."""
    otros = [k for k in (repo.kpi_mensuales_vigentes(cfg.id_proyecto, mes.anio) if repo else []) if k.mes != mes.mes]
    return calcular_anual([*otros, mes], cfg.total_racks, mes.anio, cfg.mes_inicio_acumulado(mes.anio))


def ejecutar_corrida(
    libro: str | Path,
    cfg: ConfiguracionCalculo,
    repo=None,
    *,
    hash_archivo: str,
    estado_exclusiones: str | None = None,
    filas_nuevas_export: frozenset[int] = frozenset(),
    referencia: ReferenciaExcel | None = None,
    codigos_resumen: list[tuple[str, str]] | None = None,
) -> ResultadoEjecucion:
    """Ejecuta y (si hay ``repo``) persiste una corrida. ``estado_exclusiones`` por defecto: según si
    la ``Exclusion_Matrix`` del mes de ``fin_periodo`` ya se cargó (D-17, R19.5)."""
    from etl_arena.persistence import MetadatosCorrida, construir_paquete  # evita import circular

    extra = {"id_corrida": cfg.id_corrida}
    t0 = time.perf_counter()
    segundos: dict[str, float] = {}
    if estado_exclusiones is None:
        cargada = repo is not None and repo.exclusiones_cargadas(
            cfg.id_proyecto, cfg.fin_periodo.year, cfg.fin_periodo.month
        )
        estado_exclusiones = "con_exclusiones" if cargada else "sin_exclusiones"
    meta = MetadatosCorrida(
        hash_archivo, estado_exclusiones=estado_exclusiones, filas_nuevas_export=filas_nuevas_export
    )
    salida = ResultadoEjecucion(cfg.id_corrida, "running")
    if repo is not None:
        repo.iniciar(cfg, meta)
    log.info(
        "corrida iniciada: %s %s..%s (%s)",
        cfg.tipo_corrida,
        cfg.inicio_periodo,
        cfg.fin_periodo,
        estado_exclusiones,
        extra=extra,
    )
    try:
        r = calcular_libro(libro, cfg, codigos_resumen)
        segundos["calculo"] = time.perf_counter() - t0
        salida.resultado = r
        salida.kpi_mensual = _kpi_mensual(r)
        salida.anual = _anual(cfg, salida.kpi_mensual, repo) if salida.kpi_mensual else []
        salida.reconciliacion = reconciliar(r, referencia)

        if repo is not None:
            t1 = time.perf_counter()
            paquete = construir_paquete(r, meta, repo.cargar_catalogo(), salida.kpi_mensual, salida.anual)
            salida.resumen_calidad = generar_resumen(r, paquete.anomalias_persistencia)
            salida.filas_guardadas = repo.guardar_corrida(paquete).filas_por_tabla
            if referencia is not None:
                repo.guardar_referencia_excel(referencia)
            repo.guardar_reconciliacion(
                cfg.id_corrida, referencia.id_referencia if referencia else None, a_filas(salida.reconciliacion)
            )
            segundos["persistencia"] = time.perf_counter() - t1
        else:
            salida.resumen_calidad = generar_resumen(r)

        salida.estado = salida.reconciliacion.estado_corrida
        salida.resumen_calidad["reconciliacion"] = resumen(salida.reconciliacion)
        salida.resumen_calidad["segundos"] = {k: round(v, 1) for k, v in segundos.items()}
        if repo is not None:
            repo.finalizar(
                cfg.id_corrida,
                salida.estado,
                salida.resumen_calidad,
                minutos_muestreo_derivado=r.disponibilidad.minutos_muestreo_derivado,
                id_referencia_excel=referencia.id_referencia if referencia else None,
            )
        log.info(
            "corrida terminada: %s (C14=%r, C16=%r)",
            salida.estado,
            r.disponibilidad.bloques_racks_indisponibles,
            r.disponibilidad.disponibilidad_periodo,
            extra=extra,
        )
    except Exception as exc:
        salida.estado, salida.error = "failed", f"{type(exc).__name__}: {exc}"
        log.error("corrida fallida: %s", salida.error, extra=extra)
        if repo is not None:
            try:
                repo.finalizar(cfg.id_corrida, "failed", mensaje_error=traceback.format_exc()[-4000:])
            except Exception:  # la falla original es la que importa
                log.exception("no se pudo registrar el estado failed", extra=extra)
        raise
    finally:
        salida.segundos = {k: round(v, 1) for k, v in {**segundos, "total": time.perf_counter() - t0}.items()}
    return salida


def periodo_mes_en_curso(ultimo_dato: date) -> tuple[date, date]:
    """Período por defecto de la corrida semanal (D-07): día 1 del mes del último dato → ese día."""
    return ultimo_dato.replace(day=1), ultimo_dato

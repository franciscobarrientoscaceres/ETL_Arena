"""Corrida completa (fase [E] + [R] de design.md; tarea 3.1).

config → ingesta → normalización → enriquecimiento → disponibilidad → eventos → agregaciones →
calidad → reconciliación → publicación del mes (ADR-12). Con repositorio, el ``etl_run`` se registra antes
de calcular, así que toda falla queda en SQL con ``Estado = failed`` y su mensaje (R11.3, R11.7).
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
from etl_arena.persistence.repositorio import OpcionesPublicacion, ResumenPublicacion
from etl_arena.pipeline import ResultadoCalculo, calcular_libro
from etl_arena.reconciliation import ReporteReconciliacion, reconciliar
from etl_arena.reconciliation.servicio import a_filas, resumen
from etl_arena.reporting import generar_resumen

log = logging.getLogger("etl_arena.corrida")


@dataclass
class ResultadoEjecucion:
    id_corrida: str
    estado: str  # success | parity_failed | failed
    estado_exclusiones: str = "sin_exclusiones"  # D-17
    num_corrida: int | None = None  # etl_run.NumCorrida (1, 2, 3…); None sin base de datos
    resultado: ResultadoCalculo | None = None
    reconciliacion: ReporteReconciliacion | None = None
    resumen_calidad: dict = field(default_factory=dict)
    kpi_mensual: KpiMensual | None = None
    anual: list[FilaAnual] = field(default_factory=list)
    publicada: bool = False  # reemplazó el estado vigente del mes (ADR-12, D-22)
    publicacion: ResumenPublicacion | None = None
    motivo_no_publicada: str | None = None
    segundos: dict[str, float] = field(default_factory=dict)
    error: str | None = None


def _kpi_mensual(r: ResultadoCalculo) -> KpiMensual | None:
    """KPI del mes si el período es un mes desde el día 1 (R10.2, D-07)."""
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
    opciones: OpcionesPublicacion | None = None,
    publicar_aunque_no_cuadre: bool = False,
) -> ResultadoEjecucion:
    """Calcula, reconcilia y (con ``repo``) registra la corrida y **publica** el mes (ADR-12).

    Publica —reemplaza el estado vigente del mes— solo si termina ``success`` (incluye ``sin_referencia``) o
    con ``publicar_aunque_no_cuadre``; las corridas ``golden`` nunca publican (D-22). El período de una
    corrida que publica debe ser mensual (D-20). ``estado_exclusiones`` por defecto: según si la
    ``Exclusion_Matrix`` del mes ya se cargó (D-17, R19.5)."""
    from etl_arena.persistence import MetadatosCorrida, construir_paquete_mes, validar_periodo_mensual

    extra = {"id_corrida": cfg.id_corrida}
    t0 = time.perf_counter()
    segundos: dict[str, float] = {}
    publica = repo is not None and cfg.tipo_corrida != "golden"
    if publica:
        validar_periodo_mensual(cfg)  # antes de registrar nada
    if estado_exclusiones is None:
        cargada = repo is not None and repo.exclusiones_cargadas(
            cfg.id_proyecto, cfg.fin_periodo.year, cfg.fin_periodo.month
        )
        estado_exclusiones = "con_exclusiones" if cargada else "sin_exclusiones"
    meta = MetadatosCorrida(
        hash_archivo, estado_exclusiones=estado_exclusiones, filas_nuevas_export=filas_nuevas_export
    )
    salida = ResultadoEjecucion(cfg.id_corrida, "running", estado_exclusiones=estado_exclusiones)
    if repo is not None:
        salida.num_corrida = repo.iniciar(cfg, meta)
    log.info(
        "corrida %s iniciada: %s %s..%s (%s)",
        f"N° {salida.num_corrida}" if salida.num_corrida else "(sin BD)",
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
        salida.estado = salida.reconciliacion.estado_corrida

        if repo is not None:
            t1 = time.perf_counter()
            if referencia is not None:
                repo.guardar_referencia_excel(referencia)
            repo.guardar_reconciliacion(
                cfg.id_corrida, referencia.id_referencia if referencia else None, a_filas(salida.reconciliacion)
            )
            anomalias_extra = []
            if publica and (salida.estado == "success" or publicar_aunque_no_cuadre):
                paquete = construir_paquete_mes(r, meta, repo.cargar_catalogo(), salida.kpi_mensual)
                anomalias_extra = paquete.anomalias_persistencia
                salida.publicacion = repo.publicar_mes(
                    cfg.id_corrida, salida.num_corrida, paquete, opciones, cfg.archivo_origen, hash_archivo
                )
                salida.publicada = True
            elif publica:
                salida.motivo_no_publicada = (
                    f"{salida.estado}: Python y Excel no cuadran; el estado vigente no se tocó "
                    "(--publicar-aunque-no-cuadre para publicar igual)"
                )
            else:
                salida.motivo_no_publicada = "corrida golden: no publica"
            salida.resumen_calidad = generar_resumen(r, anomalias_extra)
            segundos["persistencia"] = time.perf_counter() - t1
        else:
            salida.resumen_calidad = generar_resumen(r)

        salida.resumen_calidad["reconciliacion"] = resumen(salida.reconciliacion)
        if salida.publicacion is not None:
            salida.resumen_calidad["publicacion"] = salida.publicacion.a_dict()
        elif salida.motivo_no_publicada:
            salida.resumen_calidad["publicacion"] = {"publicada": False, "motivo": salida.motivo_no_publicada}
        salida.resumen_calidad["segundos"] = {k: round(v, 1) for k, v in segundos.items()}
        if repo is not None:
            repo.finalizar(
                cfg.id_corrida,
                salida.estado,
                salida.resumen_calidad,
                minutos_muestreo_derivado=r.disponibilidad.minutos_muestreo_derivado,
                id_referencia_excel=referencia.id_referencia if referencia else None,
                publicada=salida.publicada,
                meses_publicados=[salida.publicacion.mes_texto] if salida.publicacion else [],
            )
        log.info(
            "corrida terminada: %s%s (C14=%r, C16=%r)",
            salida.estado,
            f", publicada {salida.publicacion.mes_texto}" if salida.publicacion else ", sin publicar",
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

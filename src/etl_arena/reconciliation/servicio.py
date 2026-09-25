"""Reconciliación de una corrida contra la referencia Excel (R13)."""

from __future__ import annotations

from etl_arena.model import FilaReconciliacion, ReferenciaExcel
from etl_arena.pipeline import ResultadoCalculo
from etl_arena.reconciliation import invariantes
from etl_arena.reconciliation.modelo import ReporteReconciliacion
from etl_arena.reconciliation.niveles import nivel_1, nivel_2, nivel_3, nivel_4, nivel_5
from etl_arena.reconciliation.tolerancias import TOLERANCIAS, Tolerancias


def reconciliar(
    r: ResultadoCalculo, ref: ReferenciaExcel | None, tol: Tolerancias = TOLERANCIAS
) -> ReporteReconciliacion:
    """Niveles 1–5 + invariantes. Sin referencia: ``sin_referencia`` (la corrida no falla, R13.10)."""
    if ref is None:
        return ReporteReconciliacion(r.cfg.id_corrida, None, [], invariantes.evaluar(r, tol), sin_referencia=True)
    niveles = [f(r, ref, tol) for f in (nivel_1, nivel_2, nivel_3, nivel_4, nivel_5)]
    return ReporteReconciliacion(r.cfg.id_corrida, ref.id_referencia, niveles, invariantes.evaluar(r, tol))


def _texto(v: object) -> str | None:
    return None if v is None else str(v)[:100]


def a_filas(reporte: ReporteReconciliacion) -> list[FilaReconciliacion]:
    """Filas de ``reconciliation_result``: un resumen por nivel + cada discrepancia guardada."""
    filas: list[FilaReconciliacion] = []
    if reporte.sin_referencia:
        filas.append(FilaReconciliacion(0, "sin_referencia", None, None, None, None, None, True))
    for n in [*reporte.niveles, *([reporte.invariantes] if reporte.invariantes else [])]:
        clave = "; ".join(n.no_aplica)[:200] if n.no_aplica else n.nombre
        filas.append(
            FilaReconciliacion(
                n.nivel,
                "no_aplica" if n.no_aplica else "resumen",
                clave,
                f"{n.comparaciones} comparaciones",
                f"{n.fallidas} fallidas",
                n.max_delta,
                None,
                n.aprobado,
            )
        )
        for d in n.discrepancias:
            filas.append(
                FilaReconciliacion(
                    d.nivel, d.metrica, d.clave[:200], _texto(d.python), _texto(d.excel), d.delta, d.tolerancia, False
                )
            )
    return filas


def resumen(reporte: ReporteReconciliacion) -> dict:
    """Resumen corto para ``etl_run.ResumenCalidad`` y la notificación (R18.1)."""
    if reporte.sin_referencia:
        return {"estado": "sin_referencia"}
    return {
        "estado": "pass" if reporte.aprobado_global else "fail",
        "niveles": {
            n.nombre: {
                "aprobado": n.aprobado,
                "comparaciones": n.comparaciones,
                "fallidas": n.fallidas,
                "max_delta": n.max_delta,
            }
            for n in reporte.niveles
        },
        "invariantes": (
            {"no_aplica": reporte.invariantes.no_aplica}
            if reporte.invariantes.no_aplica
            else {"aprobado": reporte.invariantes.aprobado, "max_delta": reporte.invariantes.max_delta}
        ),
    }

"""Reconciliación Python vs Excel (R13)."""

from etl_arena.reconciliation.modelo import Discrepancia, ReporteReconciliacion, ResultadoNivel
from etl_arena.reconciliation.tolerancias import TOLERANCIAS, Tolerancias


def reconciliar(*args, **kwargs):
    """Import diferido: ``servicio`` depende de ``pipeline``, que depende de los motores."""
    from etl_arena.reconciliation.servicio import reconciliar as _reconciliar

    return _reconciliar(*args, **kwargs)


__all__ = ["TOLERANCIAS", "Discrepancia", "ReporteReconciliacion", "ResultadoNivel", "Tolerancias", "reconciliar"]

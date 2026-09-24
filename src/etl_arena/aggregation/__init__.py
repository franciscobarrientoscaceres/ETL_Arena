"""Agregaciones: Daily (R9) y KPI mensual/anual (R10)."""

from etl_arena.aggregation.diaria import calcular_diaria
from etl_arena.aggregation.mensual_anual import (
    ErrorMensual,
    FilaAnual,
    bloques_calendario,
    calcular_anual,
    registrar_mes_oficial,
)

__all__ = [
    "ErrorMensual",
    "FilaAnual",
    "bloques_calendario",
    "calcular_anual",
    "calcular_diaria",
    "registrar_mes_oficial",
]

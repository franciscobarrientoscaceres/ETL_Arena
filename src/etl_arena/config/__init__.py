"""Configuración de corrida (R1)."""

from etl_arena.config.construir import construir_config
from etl_arena.config.defaults import CONFIG_POR_DEFECTO, VERSION_ALGORITMO
from etl_arena.config.desde_excel import CELDAS_PARAMETROS, ParametrosExcel, parametros_desde_celdas
from etl_arena.config.models import ConfiguracionCalculo, ErrorConfiguracion

__all__ = [
    "CELDAS_PARAMETROS",
    "CONFIG_POR_DEFECTO",
    "VERSION_ALGORITMO",
    "ConfiguracionCalculo",
    "ErrorConfiguracion",
    "ParametrosExcel",
    "construir_config",
    "parametros_desde_celdas",
]

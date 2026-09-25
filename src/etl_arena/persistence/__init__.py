"""Persistencia SQL Server append-only (R11, R12; ADR-08, ADR-09, ADR-10)."""

from etl_arena.persistence.conexion import (
    ErrorConexion,
    cargar_env,
    crear_base_si_no_existe,
    crear_engine,
    eliminar_base,
    url_configurada,
)
from etl_arena.persistence.esquema import ORDEN_SCRIPTS, aplicar_esquema, lotes, reiniciar_esquema
from etl_arena.persistence.paquete import MetadatosCorrida, PaqueteCorrida, Tabla, construir_paquete, filas_staging
from etl_arena.persistence.repositorio import RepositorioCorridas, ResultadoGuardado

__all__ = [
    "ORDEN_SCRIPTS",
    "ErrorConexion",
    "MetadatosCorrida",
    "PaqueteCorrida",
    "RepositorioCorridas",
    "ResultadoGuardado",
    "Tabla",
    "aplicar_esquema",
    "cargar_env",
    "construir_paquete",
    "crear_base_si_no_existe",
    "crear_engine",
    "eliminar_base",
    "filas_staging",
    "lotes",
    "reiniciar_esquema",
    "url_configurada",
]

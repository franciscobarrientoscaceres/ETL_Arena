"""Fixtures compartidos. El libro real solo se lee en tests con marker ``golden``."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
LIBRO_REAL = Path(
    os.environ.get(
        "ETL_ARENA_XLSM", RAIZ / "data" / "AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm"
    )
)
if not LIBRO_REAL.is_absolute():
    LIBRO_REAL = RAIZ / LIBRO_REAL


@pytest.fixture(scope="session")
def ruta_libro_real() -> Path:
    if not LIBRO_REAL.exists():
        pytest.skip(f"libro real no disponible: {LIBRO_REAL}")
    return LIBRO_REAL


@pytest.fixture(scope="session")
def libro_real_sep(ruta_libro_real):
    """(cfg, LibroCrudo) del libro de septiembre con sus parámetros efectivos; se lee una vez."""
    from etl_arena.config import construir_config, parametros_desde_celdas
    from etl_arena.ingestion import leer_celdas_parametros, leer_libro

    params = parametros_desde_celdas(leer_celdas_parametros(ruta_libro_real))
    cfg = construir_config(**params.valores, tipo_corrida="golden", archivo_origen=ruta_libro_real.name)
    return cfg, leer_libro(ruta_libro_real, cfg)


@pytest.fixture(scope="session")
def matriz_real_sep(libro_real_sep):
    """(cfg, MatrizPCS, DatosActividad, anomalías de enriquecimiento)."""
    from etl_arena.enrichment import asociar_actividad
    from etl_arena.normalization import a_matriz, validar_esquema

    cfg, libro = libro_real_sep
    mapa, anomalias = validar_esquema(libro.encabezado, cfg)
    assert not anomalias, anomalias
    m = a_matriz(libro, mapa, cfg)
    act, anomalias_pa = asociar_actividad(m, libro.actividad)
    return cfg, m, act, anomalias_pa

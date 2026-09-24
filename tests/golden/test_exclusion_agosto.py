"""Exclusion_Matrix contra el libro de agosto 2026 (F-37, marker golden).

El libro ``…_20260923_agosto_2026.xlsm`` trae la primera ``Exclusion_Matrix`` (solo agosto)
y su ``Calculation-Availability`` ya calculado con la regla de exclusión (macro de agosto,
C19 = "Yes"). Las macros y celdas oficiales siguen siendo las de septiembre: aquí solo se
comparan los valores cacheados del KPI y de la tabla de resultados (niveles 2–4). La
``ListOfFaults`` de ese libro no se recalculó (quedó con L2/L4/L10 de septiembre) y su
``Daily`` usa otra estructura: no son referencia.
"""

from __future__ import annotations

from collections import Counter
from datetime import date
from pathlib import Path

import numpy as np
import pytest

from etl_arena.config import construir_config
from etl_arena.ingestion import LibroXlsx
from etl_arena.pipeline import calcular_libro
from etl_arena.reconciliation import TOLERANCIAS as T

pytestmark = pytest.mark.golden
LIBRO_AGOSTO = (
    Path(__file__).resolve().parents[2] / "data" / "AvailabilityCalculation_PCS&Batteries_20260923_agosto_2026.xlsm"
)
HOJA = "Calculation-Availability"


@pytest.fixture(scope="module")
def agosto():
    if not LIBRO_AGOSTO.exists():
        pytest.skip(f"libro de agosto no disponible: {LIBRO_AGOSTO}")
    cfg = construir_config(inicio_periodo=date(2026, 8, 1), fin_periodo=date(2026, 8, 31), tipo_corrida="golden")
    resultado = calcular_libro(LIBRO_AGOSTO, cfg)
    with LibroXlsx(LIBRO_AGOSTO) as libro:
        tabla = dict(libro.filas(HOJA, min_fila=4, max_col=67))
        celdas = libro.celdas([f"{HOJA}!C{f}" for f in (6, 8, 12, 14, 16, 19, 21)])
    return resultado, tabla, celdas


def test_parametros_del_libro(agosto):
    _, _, c = agosto
    assert (c[f"{HOJA}!C6"], c[f"{HOJA}!C8"]) == (46235.0, 46265.0)  # 01-ago .. 31-ago
    assert c[f"{HOJA}!C19"] == "Yes" and c[f"{HOJA}!C21"] == "No"  # excusable / solo operacional


def test_kpi(agosto):
    r, _, c = agosto
    d = r.disponibilidad
    assert d.bloques_muestreo == c[f"{HOJA}!C12"] == 2976
    assert abs(d.bloques_racks_indisponibles - c[f"{HOJA}!C14"]) <= T.c14
    assert abs(d.disponibilidad_periodo - c[f"{HOJA}!C16"]) <= T.kpi_disponibilidad


def test_objetivo_paridad_exacta(agosto):
    r, tabla, c = agosto
    d = r.disponibilidad
    assert d.bloques_racks_indisponibles == c[f"{HOJA}!C14"] == 264731.7679999999
    assert d.disponibilidad_periodo == c[f"{HOJA}!C16"]
    distintas = []
    for k in range(d.bloques_muestreo):
        fila = tabla.get(4 + k, {})
        assert fila.get(5) == d.serial[k]
        assert fila.get(67) == d.bo_evento_excusado_pa[k]
        for j in range(61):
            v, py = fila.get(6 + j), d.ponderadas[k, j]
            if (v is None) != bool(np.isnan(py)) or (v is not None and v != py):
                distintas.append((k, j, py, v))
    assert distintas == []


def test_matriz_de_agosto(agosto):
    r, _, _ = agosto
    exc = r.exclusion
    valores = Counter(exc.valor.ravel().tolist())
    assert valores[1.0] == 4982 and valores[2.0] == 1498  # celdas PCS01..PCS61 con EE
    assert Counter(c for c in exc.comentario if c) == {
        "CPF": 362,
        "External": 11,
        "Actualizacion de firware de PCSs 1,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18": 1,
    }
    # los 6 PCS con EE valor 2 tenían 3 módulos antes del evento → 1 batería previa
    previas = {m for m in exc.baterias_previas[~np.isnan(exc.baterias_previas)].tolist()}
    assert previas == {1.0}
    # marca aislada: PCS61 = 1 en la fila 2 (08-abr 00:15), fuera de agosto
    assert exc.valor[0, 60] == 1.0 and not exc.valor[1:11000].any()


def test_sin_efecto_sin_flag(agosto):
    """Con C31 = "No" la matriz no pondera: el KPI vuelve al valor sin exclusiones."""
    r, _, _ = agosto
    from etl_arena.availability import calcular

    cfg = construir_config(
        inicio_periodo=date(2026, 8, 1), fin_periodo=date(2026, 8, 31), aplicar_evento_excusable=False
    )
    sin = calcular(r.matriz, r.actividad, cfg, r.exclusion)
    assert abs(sin.bloques_racks_indisponibles - 527396.044) <= T.c14  # agosto sin exclusiones
    assert sin.bloques_racks_indisponibles > r.disponibilidad.bloques_racks_indisponibles

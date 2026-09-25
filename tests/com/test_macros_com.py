"""Macros reales vía COM (tarea 4.3; R3.5, R14): la referencia extraída tras correrlas iguala al golden.

Abren Excel en el escritorio: solo corren con ``ETL_ARENA_EXCEL=1`` en la PC local y sin usar Excel
mientras tanto. Siempre sobre una copia en ``tmp``; el libro de ``data/`` no se toca.
"""

from __future__ import annotations

import dataclasses
import json
import math
import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parents[1] / "golden"))
from test_golden_runner import cfg_desde_golden  # noqa: E402

pytestmark = [
    pytest.mark.excel,
    pytest.mark.golden,
    pytest.mark.skipif(sys.platform != "win32", reason="COM solo en Windows"),
    pytest.mark.skipif(os.environ.get("ETL_ARENA_EXCEL") != "1", reason="ETL_ARENA_EXCEL=1 para abrir Excel"),
]
GOLDEN = Path(__file__).parents[1] / "golden" / "data" / "golden_2026_09_september.json"


def _mapa(resumen):
    return {tuple(f)[0]: tuple(f)[1:] for f in resumen}


@pytest.fixture(scope="module")
def corrida_com(ruta_libro_real, tmp_path_factory):
    from etl_arena.workbook import copiar_libro_trabajo, ejecutar_macros, extraer_referencia

    g = json.loads(GOLDEN.read_text(encoding="utf-8"))
    cfg = cfg_desde_golden(g)
    libro = copiar_libro_trabajo(ruta_libro_real, tmp_path_factory.mktemp("com"))
    tiempos = ejecutar_macros(libro, cfg, timeout_s=1800)
    return g, cfg, tiempos, extraer_referencia(libro, "com"), extraer_referencia(ruta_libro_real, "original")


def test_macros_corren_en_orden(corrida_com):
    from etl_arena.workbook.macros import MACROS

    tiempos = corrida_com[2]
    assert [k for k in tiempos if k in MACROS] == list(MACROS)  # en orden; además: abrir, recalcular, guardar
    assert {"abrir", "guardar"} <= set(tiempos)


def _cerca(a, b) -> bool:
    if isinstance(a, float) and isinstance(b, float):
        return math.isclose(a, b, rel_tol=1e-12, abs_tol=1e-9)
    if isinstance(a, (tuple, list)) and isinstance(b, (tuple, list)):
        return len(a) == len(b) and all(_cerca(x, y) for x, y in zip(a, b, strict=True))
    return a == b


def test_referencia_com_iguala_al_libro_original(corrida_com):
    """KPI, tabla y Daily idénticos. Eventos y N:Q al último bit: el libro original se corrió con
    L14 = "Yes" y ahora se escribe "No" para no excusar con PlantActivity!D (F-44); el VBA calcula H
    por otro camino aritmético (Δ ≈ 1e-16 en H)."""
    _, _, _, ref, original = corrida_com
    assert (ref.c12, ref.c14, ref.c16, ref.c19) == (original.c12, original.c14, original.c16, original.c19)
    assert ref.tabla == original.tabla and ref.bo == original.bo
    assert ref.diario == original.diario
    assert len(ref.eventos) == len(original.eventos)
    assert all(
        _cerca(dataclasses.astuple(a), dataclasses.astuple(b))
        for a, b in zip(ref.eventos, original.eventos, strict=True)
    )
    esperado, obtenido = _mapa(original.resumen_codigos), _mapa(ref.resumen_codigos)  # sin Add2 no se ordena
    assert esperado.keys() == obtenido.keys() and all(_cerca(esperado[k], obtenido[k]) for k in esperado)


def test_referencia_com_iguala_al_golden(corrida_com):
    g, _, _, ref, _ = corrida_com
    assert len(ref.eventos) == g["fault_events_summary"]["total_events"]
    assert len(ref.diario) == len(g["daily"])


def test_python_reconcilia_con_la_referencia_com(corrida_com, ruta_libro_real):
    from etl_arena.pipeline import calcular_libro
    from etl_arena.reconciliation import reconciliar

    _, cfg, _, ref, _ = corrida_com
    rep = reconciliar(calcular_libro(ruta_libro_real, cfg), ref)
    assert all(n.aprobado for n in rep.niveles), [n for n in rep.niveles if not n.aprobado]
    assert all(not n.max_delta for n in rep.niveles if n.nivel <= 4)  # KPI, tabla y Daily: bit a bit
    assert max(n.max_delta or 0.0 for n in rep.niveles) < 1e-9  # eventos: último bit (L14 = "No", F-44)

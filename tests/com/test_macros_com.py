"""Macros reales vía COM (tarea 4.3; R3.5, R14): la referencia extraída tras correrlas iguala al golden.

Abren Excel en el escritorio: solo corren con ``ETL_ARENA_EXCEL=1`` en la PC local y sin usar Excel
mientras tanto. Siempre sobre una copia en ``tmp``; el libro de ``data/`` no se toca.
"""

from __future__ import annotations

import json
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

    assert list(corrida_com[2]) == list(MACROS)


def test_referencia_com_iguala_al_libro_original(corrida_com):
    _, _, _, ref, original = corrida_com
    assert (ref.c12, ref.c14, ref.c16, ref.c19) == (original.c12, original.c14, original.c16, original.c19)
    assert ref.tabla == original.tabla and ref.bo == original.bo
    assert ref.diario == original.diario and ref.eventos == original.eventos
    assert _mapa(ref.resumen_codigos) == _mapa(original.resumen_codigos)  # sin Add2 (Excel 2016) no se ordena


def test_referencia_com_iguala_al_golden(corrida_com):
    g, _, _, ref, _ = corrida_com
    assert len(ref.eventos) == g["fault_events_summary"]["total_events"]
    assert len(ref.diario) == len(g["daily"])


def test_python_reconcilia_con_la_referencia_com(corrida_com, ruta_libro_real):
    from etl_arena.pipeline import calcular_libro
    from etl_arena.reconciliation import reconciliar

    _, cfg, _, ref, _ = corrida_com
    rep = reconciliar(calcular_libro(ruta_libro_real, cfg), ref)
    assert all(n.aprobado and not n.max_delta for n in rep.niveles), [n for n in rep.niveles if not n.aprobado]

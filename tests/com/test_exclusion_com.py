"""Carga de la ``Exclusion_Matrix`` vía COM (tarea 4.12; R3.8, R19.5, D-17): la matriz de agosto escrita en una
copia del libro de septiembre produce el diff esperado, es idempotente y da el C14 del libro de agosto.

Abre Excel: solo con ``ETL_ARENA_EXCEL=1`` en la PC local y sin usar Excel mientras tanto. Siempre sobre una
copia en ``tmp``; los libros de ``data/`` no se tocan.
"""

from __future__ import annotations

import hashlib
import os
import sys
from datetime import date
from pathlib import Path

import pytest

pytestmark = [
    pytest.mark.excel,
    pytest.mark.golden,
    pytest.mark.skipif(sys.platform != "win32", reason="COM solo en Windows"),
    pytest.mark.skipif(os.environ.get("ETL_ARENA_EXCEL") != "1", reason="ETL_ARENA_EXCEL=1 para abrir Excel"),
]
LIBRO_AGOSTO = Path(__file__).parents[2] / "data" / "AvailabilityCalculation_PCS&Batteries_20260923_agosto_2026.xlsm"
C14_AGOSTO = 264731.7679999999  # libro de agosto, "Con Exclusiones" (4.12, validado 2026-09-25)


def _sha(ruta: Path) -> str:
    return hashlib.sha256(ruta.read_bytes()).hexdigest()


@pytest.mark.skipif(not LIBRO_AGOSTO.exists(), reason="falta el libro de agosto en data/")
def test_matriz_de_agosto_en_copia_de_septiembre(ruta_libro_real, tmp_path):
    from etl_arena.config import construir_config
    from etl_arena.pipeline import calcular_libro
    from etl_arena.workbook import copiar_libro_trabajo
    from etl_arena.workbook import exclusion as em

    cfg = construir_config(
        inicio_periodo=date(2026, 8, 1), fin_periodo=date(2026, 8, 31), tipo_corrida="cierre_mensual"
    )
    p = cfg.total_pcs
    originales = {r: _sha(r) for r in (ruta_libro_real, LIBRO_AGOSTO)}
    copia = copiar_libro_trabajo(ruta_libro_real, tmp_path)

    entrega = em.leer_entrega(LIBRO_AGOSTO, 2026, 8, p)
    plan = em.planificar(copia, entrega, p)
    assert not plan.hoja_existe and len(plan.filas) == 2976  # 31 días × 96 bloques
    assert len(plan.cambios) == 7383  # mismo diff que la carga de TEST/QA (2026-09-26)
    assert all(c.fila in plan.filas for c in plan.cambios)

    em.aplicar_plan(copia, plan, p, timeout_s=900)
    repetido = em.planificar(copia, entrega, p)
    assert repetido.hoja_existe and repetido.cambios == []  # cargar dos veces la misma entrega no cambia nada

    r = calcular_libro(copia, cfg)
    assert r.disponibilidad.bloques_muestreo == 2976
    assert r.disponibilidad.bloques_racks_indisponibles == C14_AGOSTO
    assert {r: _sha(r) for r in originales} == originales  # los libros de data/ no se tocaron

"""Ingesta, normalización y enriquecimiento sobre el libro real de septiembre (marker golden).

Cifras del perfilado en docs/data-contract-libro.md.
"""

from datetime import date

import numpy as np
import pytest

pytestmark = pytest.mark.golden


def test_parametros_efectivos(libro_real_sep):
    cfg, _ = libro_real_sep
    assert (cfg.inicio_periodo, cfg.fin_periodo, cfg.fin_diario) == (
        date(2026, 9, 1),
        date(2026, 9, 21),
        date(2026, 9, 21),
    )
    assert cfg.total_racks == 2928 and not cfg.solo_tiempo_operacional
    assert cfg.aplicar_evento_excusable is False and cfg.aplicar_evento_excusable_eventos is True  # F-05


def test_filas_y_seriales(libro_real_sep):  # R4.1, R4.2
    _, libro = libro_real_sep
    assert len(libro.filas) == 15990 and libro.filas_descartadas == 0
    assert libro.filas[0].numero_fila == 2 and libro.filas[0].serial == 46120.010416666664
    assert libro.filas[-1].numero_fila == 15991 and libro.filas[-1].serial == 46286.59375


def test_unica_anomalia_de_timestamp_es_dst(libro_real_sep):  # R16.3
    _, libro = libro_real_sep
    a = [x for x in libro.anomalias if x.tipo not in ("celda_a_vacia", "filas_truncadas", "exclusion_matrix_ausente")]
    assert [(x.tipo, x.severidad) for x in a] == [("dst_salto", "info")]
    assert "2026-09-06 00:00 → 2026-09-06 01:00" in a[0].detalle


def test_matriz(matriz_real_sep):  # R5, F-02
    cfg, m, _, _, _ = matriz_real_sep
    assert m.modulos.shape == (15990, 61)
    assert int(m.modulos_nulo.sum()) == 566
    fracc = ~m.modulos_nulo & (m.modulos != np.floor(np.nan_to_num(m.modulos)))
    assert int(fracc.sum()) == 15655
    assert sum(isinstance(v, float) for v in m.falla.ravel()) == 1705  # F-14


def test_plant_activity(matriz_real_sep):  # R6, F-01, F-31 — GT-8
    _, _, act, anomalias, exc = matriz_real_sep
    tipos = [a.tipo for a in anomalias]
    assert tipos.count("pa_sin_timestamp") == 659 and "pa_desalineado" not in tipos and "pa_vacio" not in tipos
    primera = min(a.numero_fila for a in anomalias if a.tipo == "pa_sin_timestamp")
    assert primera == 15333
    assert not exc.valor.any()  # el libro de septiembre no trae Exclusion_Matrix (F-37)
    assert int((act.factor_operacional == 1).sum()) == 10504 and int((act.evento_excusado_pa == 0).sum()) == 553


def test_catalogos(ruta_libro_real):  # F-19, F-35
    """PCS-Fault: 167 filas, 163 códigos distintos (F228, F230, F231, F232 repetidos con Meaning
    Crítico/Parcial, D-14). ListOfFaults!N: 166 filas (F230-F232 dos veces), mismos 163 códigos."""
    from collections import Counter

    from etl_arena.ingestion.catalogos import leer_catalogo_fallas, leer_codigos_resumen

    catalogo = leer_catalogo_fallas(ruta_libro_real)
    assert len(catalogo) == 167 and catalogo[0].codigo == "F0" and catalogo[-1].codigo == "F257"
    repetidos = sorted(c for c, k in Counter(x.codigo for x in catalogo).items() if k > 1)
    assert repetidos == ["F228", "F230", "F231", "F232"]
    resumen = leer_codigos_resumen(ruta_libro_real)
    assert len(resumen) == 166 and {c for c, _ in resumen} == {c.codigo for c in catalogo}

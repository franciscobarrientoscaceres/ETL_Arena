"""Properties 1 y 2: forma de la matriz y anomalías de timestamp."""

import tempfile
from pathlib import Path

import numpy as np
from fabricas import BLOQUE, SERIAL_SEP_1, config_prueba, crear_libro, fila_raw
from hypothesis import given, settings
from hypothesis import strategies as st

from etl_arena.ingestion import detectar_anomalias_timestamp, leer_libro
from etl_arena.normalization import a_matriz, validar_esquema

_modulo = st.one_of(st.none(), st.floats(min_value=0, max_value=4, allow_nan=False))


# Feature: etl-arena-availability, Property 1
@settings(max_examples=200, deadline=None)
@given(
    st.integers(min_value=1, max_value=3).flatmap(
        lambda p: st.lists(st.lists(_modulo, min_size=p, max_size=p), min_size=1, max_size=12)
    )
)
def test_forma_y_valores_bit_a_bit(filas_modulos):
    p = len(filas_modulos[0])
    cfg = config_prueba(total_pcs=p)
    filas = [fila_raw(SERIAL_SEP_1 + (i + 1) * BLOQUE, mods) for i, mods in enumerate(filas_modulos)]
    with tempfile.TemporaryDirectory() as d:
        libro = leer_libro(crear_libro(Path(d) / "l.xlsx", filas, total_pcs=p), cfg)
    mapa, _ = validar_esquema(libro.encabezado, cfg)
    m = a_matriz(libro, mapa, cfg)
    assert m.modulos.shape == (len(filas_modulos), p)
    disp = m.modulos_disponibles(cfg.baterias_por_pcs)
    for i, mods in enumerate(filas_modulos):
        for j, v in enumerate(mods):
            if v is None:
                assert m.modulos_nulo[i, j] and disp[i, j] == cfg.baterias_por_pcs
            else:
                assert not m.modulos_nulo[i, j] and m.modulos[i, j] == v  # sin redondear ni truncar


_evento = st.sampled_from(["ok", "duplicado", "fuera_de_orden", "hueco", "frecuencia_distinta"])


# Feature: etl-arena-availability, Property 2
@settings(max_examples=300)
@given(st.lists(_evento, min_size=1, max_size=40))
def test_anomalias_inyectadas_se_reportan_sin_falsos_positivos(eventos):
    # Serie en mayo (sin DST); cada evento define el paso respecto a la fila anterior.
    base = 46150.5
    seriales, esperadas = [base], []
    for k, ev in enumerate(eventos, start=1):
        paso = {"ok": 15, "duplicado": 0, "fuera_de_orden": -15, "hueco": 45, "frecuencia_distinta": 10}[ev]
        seriales.append(seriales[-1] + paso / 1440)
        if ev != "ok":
            esperadas.append((ev, k + 2))
    obtenidas = [
        (a.tipo, a.numero_fila) for a in detectar_anomalias_timestamp(seriales, list(range(2, len(seriales) + 2)), 15)
    ]
    assert obtenidas == esperadas
    assert np.all(np.isfinite(seriales))

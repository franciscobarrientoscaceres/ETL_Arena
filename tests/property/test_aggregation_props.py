"""Property 9: Daily."""

from datetime import date

from fabricas import actividad, config_prueba, matriz, serial_min
from hypothesis import given, settings
from hypothesis import strategies as st

from etl_arena.aggregation import calcular_diaria
from etl_arena.availability import calcular

_modulo = st.one_of(st.none(), st.just(4.0), st.floats(min_value=0, max_value=4, allow_nan=False))


# Feature: etl-arena-availability, Property 9
@settings(max_examples=200, deadline=None)
@given(
    st.lists(st.tuples(st.integers(min_value=0, max_value=3 * 96 - 1), _modulo, _modulo), min_size=2, max_size=40),
    st.lists(st.floats(min_value=0, max_value=1, allow_nan=False), min_size=40, max_size=40),
    st.booleans(),
)
def test_daily(filas, operacional, solo_operacional):
    filas = sorted(filas, key=lambda f: f[0])
    seriales = [serial_min(15 * (b + 1)) for b, _, _ in filas]
    m = matriz([[a, b] for _, a, b in filas], seriales=seriales)
    act = actividad(m.n, operacional=operacional[: m.n])
    cfg = config_prueba(
        inicio_periodo=date(2026, 9, 1),
        fin_periodo=date(2026, 9, 3),
        solo_tiempo_operacional=solo_operacional,
        aplicar_evento_excusable=False,
    )
    res = calcular(m, act, cfg)
    dias = calcular_diaria(res, cfg)
    assert len(dias) == 3
    total = sum(12 * v for fila in res.ponderadas_o_cero().tolist() for v in fila)
    assert abs(sum(d.diario for d in dias) - total) <= 1e-9 * max(1.0, total)
    assert all(b.acumulado >= a.acumulado for a, b in zip(dias, dias[1:], strict=False))
    for d in dias:
        if res.minutos_muestreo_derivado:
            assert d.disponibilidad == 1 - d.acumulado / (
                cfg.total_racks * 24 * 60 * d.numero_dia / res.minutos_muestreo_derivado
            )
    # el diario no depende del factor operacional (F-08)
    cfg_sin = config_prueba(
        inicio_periodo=date(2026, 9, 1),
        fin_periodo=date(2026, 9, 3),
        solo_tiempo_operacional=not solo_operacional,
        aplicar_evento_excusable=False,
    )
    assert [d.diario for d in calcular_diaria(calcular(m, act, cfg_sin), cfg_sin)] == [d.diario for d in dias]

"""Properties 6, 7 y 11 del MotorEventosFalla."""

from datetime import date

from fabricas import BLOQUE, config_prueba, matriz
from hypothesis import given, settings
from hypothesis import strategies as st

from etl_arena.excel_semantics import datetime_a_serial
from etl_arena.fault_events import detectar_eventos
from etl_arena.fault_events.motor import _descripcion

_modulo = st.one_of(st.none(), st.just(4.0), st.just(4.0), st.floats(min_value=0, max_value=3.99, allow_nan=False))


def _rachas(columna):
    """Rachas maximales de filas con módulo < 4 (separadas por 4 o vacío)."""
    rachas, actual = [], 0
    for v in columna:
        if v is not None and v < 4:
            actual += 1
        elif actual:
            rachas.append(actual)
            actual = 0
    if actual:
        rachas.append(actual)
    return rachas


# Feature: etl-arena-availability, Property 6
@settings(max_examples=250)
@given(
    st.integers(min_value=1, max_value=3).flatmap(
        lambda p: st.lists(st.lists(_modulo, min_size=p, max_size=p), min_size=1, max_size=25)
    )
)
def test_cobertura_de_eventos(filas):
    p = len(filas[0])
    # todas las filas dentro del período y la fila siguiente a la última vacía: nunca hay arrastre
    r = detectar_eventos(
        matriz(filas, primera_fila=3),
        None,
        config_prueba(total_pcs=p, aplicar_evento_excusable_eventos=False),
        15.0,
    )
    esperadas = [n for j in range(p) for n in _rachas([f[j] for f in filas])]
    assert [e.numero_bloques for e in r.cerrados] == esperadas
    assert sum(e.numero_bloques for e in r.cerrados) == sum(1 for f in filas for v in f if v is not None and v < 4)
    assert r.incompleto is None and not any(e.arrastrado_excel for e in r.cerrados)


_desc = st.one_of(st.just("NO FAULTS"), st.just(""), st.none(), st.just("F55 EXTERNAL"), st.just(169.0))


# Feature: etl-arena-availability, Property 7
@settings(max_examples=200)
@given(_desc, _desc)
def test_arbol_de_descripcion(actual, anterior):
    g, fallback = _descripcion(actual, anterior, True)
    if actual == "NO FAULTS":
        if anterior == "NO FAULTS":
            assert (g, fallback) == ("F13 NO MODULES", False)
        elif anterior in (None, ""):
            assert (g, fallback) == ("F1 Watchdog", True)
        else:
            assert (g, fallback) == (anterior, True)
    elif actual in (None, ""):
        assert (g, fallback) == ("F1 Watchdog", False)
    else:
        assert (g, fallback) == (actual, False)


# Feature: etl-arena-availability, Property 11
@settings(max_examples=200)
@given(st.floats(min_value=0, max_value=3.99, allow_nan=False), st.floats(min_value=0, max_value=3.99, allow_nan=False))
def test_arrastre(m1, m2):
    fin = datetime_a_serial(date(2026, 9, 2))
    seriales = [fin - 2 * BLOQUE, fin - BLOQUE, fin, fin + BLOQUE]
    # PCS 1 en falla a las 23:45 y sigue a las 00:00 (= L4+1); PCS 2 falla a las 23:45
    filas = [[4.0, 4.0], [m1, m2], [m1, 4.0], [4.0, 4.0]]
    r = detectar_eventos(
        matriz(filas, seriales=seriales, primera_fila=3),
        None,
        config_prueba(total_pcs=2, aplicar_evento_excusable_eventos=False),
        15.0,
    )
    (e,) = r.cerrados
    assert e.arrastrado_excel and e.numero_pcs == 2 and e.numero_bloques == 2
    assert e.suma_bloques == ((0.0 + 4) - m1 + 4) - m2

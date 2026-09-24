"""Properties 3, 4 y 5 del MotorDisponibilidad."""

from fabricas import BLOQUE, SERIAL_SEP_1, actividad, config_prueba, matriz
from hypothesis import given, settings
from hypothesis import strategies as st

from etl_arena.availability import calcular

_modulo = st.one_of(
    st.none(), st.sampled_from([0.0, 1.0, 2.0, 3.0, 4.0]), st.floats(min_value=0, max_value=4, allow_nan=False)
)
_factor = st.floats(min_value=0, max_value=1, allow_nan=False)


# Feature: etl-arena-availability, Property 3
@settings(max_examples=300)
@given(_modulo, _factor, _factor, st.booleans(), st.booleans())
def test_impacto_por_celda(m, fe, fo, excusable, operacional):
    cfg = config_prueba(total_pcs=1, aplicar_evento_excusable=excusable, solo_tiempo_operacional=operacional)
    r = calcular(matriz([[m]]), actividad(1, operacional=fo, excusable=fe), cfg)
    if m is None or not m < 4:
        esperado = 0.0
    else:
        pond = (4 - m) * fe if excusable else 4 - m
        esperado = 12 * pond * fo if operacional else 12 * pond
    assert r.bloques_racks_indisponibles == esperado


# Feature: etl-arena-availability, Property 4
@settings(max_examples=200)
@given(st.lists(st.integers(min_value=0, max_value=3), min_size=1, max_size=30))
def test_c12_cuenta_filas_con_duplicados(repeticiones):
    seriales = []
    for k, rep in enumerate(repeticiones):
        seriales += [SERIAL_SEP_1 + (k + 1) * BLOQUE] * (rep + 1)
    r = calcular(matriz([[4, 4]] * len(seriales), seriales=seriales), actividad(len(seriales)), config_prueba())
    assert r.bloques_muestreo == len(seriales)


# Feature: etl-arena-availability, Property 5
@settings(max_examples=200)
@given(
    st.integers(min_value=1, max_value=20).flatmap(
        lambda n: st.tuples(
            st.lists(st.lists(_modulo, min_size=3, max_size=3), min_size=n, max_size=n),
            st.lists(_factor, min_size=n, max_size=n),
        )
    )
)
def test_c14_es_suma_secuencial_fila_pcs(datos):
    filas, fe = datos
    r = calcular(matriz(filas), actividad(len(filas), excusable=fe), config_prueba(total_pcs=3))
    suma = 0.0
    for k in range(len(filas)):
        for j in range(3):
            suma = suma + float(r.impacto[k, j])
    assert r.bloques_racks_indisponibles == suma  # igualdad exacta

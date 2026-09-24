"""Properties 3, 4 y 5 del MotorDisponibilidad."""

from fabricas import BLOQUE, SERIAL_SEP_1, actividad, config_prueba, exclusion, matriz
from hypothesis import given, settings
from hypothesis import strategies as st

from etl_arena.availability import calcular

_modulo = st.one_of(
    st.none(), st.sampled_from([0.0, 1.0, 2.0, 3.0, 4.0]), st.floats(min_value=0, max_value=4, allow_nan=False)
)
_factor = st.floats(min_value=0, max_value=1, allow_nan=False)
_ee = st.sampled_from([None, 0.0, 1.0, 2.0])


# Feature: etl-arena-availability, Property 3
@settings(max_examples=400)
@given(_modulo, _modulo, _ee, _factor, st.booleans(), st.booleans())
def test_impacto_por_celda(m_previo, m, ee, fo, excusable, operacional):
    """Fila 1 sin EE (módulos previos); fila 2 con el valor de Exclusion_Matrix ``ee`` (F-37)."""
    cfg = config_prueba(total_pcs=1, aplicar_evento_excusable=excusable, solo_tiempo_operacional=operacional)
    mat = matriz([[m_previo], [m]])
    r = calcular(mat, actividad(2, operacional=fo), cfg, exclusion(mat, [[0.0], [ee]]))
    if m is None or not m < 4:
        esperado = 0.0
    else:
        if not excusable or ee in (None, 0.0):
            pond = 4 - m if not excusable else (4 - m) * (1 - 0.0)
        elif ee == 1.0:
            pond = (4 - m) * (1 - 1.0)
        else:  # 2: baterías que faltaban antes del EE
            pond = 4 - m_previo if m_previo is not None and m_previo < 4 else 0.0
        esperado = 12 * pond * fo if operacional else 12 * pond
    assert r.impacto[1, 0] == esperado


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
            st.lists(st.lists(_ee, min_size=3, max_size=3), min_size=n, max_size=n),
        )
    )
)
def test_c14_es_suma_secuencial_fila_pcs(datos):
    filas, ee = datos
    mat = matriz(filas)
    r = calcular(mat, actividad(len(filas)), config_prueba(total_pcs=3), exclusion(mat, ee))
    suma = 0.0
    for k in range(len(filas)):
        for j in range(3):
            suma = suma + float(r.impacto[k, j])
    assert r.bloques_racks_indisponibles == suma  # igualdad exacta

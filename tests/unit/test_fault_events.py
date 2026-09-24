"""Tests de MotorEventosFalla (R8). Cada caso reproduce un camino del VBA de mcoCreateList."""

from datetime import date

import pytest
from fabricas import BLOQUE, SERIAL_SEP_1, actividad, config_prueba, matriz, serial_min

from etl_arena.excel_semantics import datetime_a_serial
from etl_arena.fault_events import detectar_eventos

C23 = 15.0


def _eventos(modulos, fallas=None, seriales=None, primera_fila=3, siguiente=None, fe=1.0, **cfg):
    m = matriz(modulos, fallas=fallas, seriales=seriales, primera_fila=primera_fila, siguiente=siguiente)
    config = config_prueba(total_pcs=m.p, **cfg)
    return detectar_eventos(m, actividad(m.n, excusable=fe, primera_fila=primera_fila), config, C23), m


def test_evento_simple_inicio_menos_c23_y_fin_ultima_fila():  # F-03
    r, m = _eventos([[4], [3], [2], [4]], fallas=[["NO FAULTS"], ["F55 X"], ["F55 X"], ["NO FAULTS"]])
    (e,) = r.cerrados
    assert e.numero_pcs == 1 and e.codigo_falla == "F55" and e.descripcion_falla == "F55 X"
    assert e.serial_inicio == serial_min(15)  # primera fila en falla (00:30) − C23 = 00:15
    assert e.serial_fin == m.serial[2]  # última fila en falla
    assert e.duracion_horas == 24 * (e.serial_fin - e.serial_inicio)
    assert e.promedio_baterias == (1 + 2) / 2 and e.horas_rack == 12 * e.duracion_horas * 1.5
    assert r.horas_rack_totales == e.horas_rack and r.incompleto is None


def test_celda_vacia_parte_el_evento():
    r, _ = _eventos([[4], [3], [None], [3], [4]])
    assert [(e.numero_bloques, e.serial_fin is not None) for e in r.cerrados] == [(1, True), (1, True)]


def test_fraccionario_continua_el_evento():  # solo "= 4" o vacío cortan (F-02)
    r, _ = _eventos([[4], [3.9999], [3.5], [4]])
    assert len(r.cerrados) == 1 and r.cerrados[0].numero_bloques == 2


@pytest.mark.parametrize(
    ("actual", "anterior", "esperado", "fallback"),
    [
        ("F55 X", "NO FAULTS", "F55 X", False),  # descripción propia
        ("NO FAULTS", "F4 DISCHARGE", "F4 DISCHARGE", True),  # toma la anterior
        ("NO FAULTS", "NO FAULTS", "F13 NO MODULES", False),
        ("NO FAULTS", None, "F1 Watchdog", True),  # anterior vacía → "" → F1 (F-04)
        (None, "NO FAULTS", "F1 Watchdog", False),
        ("NO FAULTS", 169.0, 169.0, True),  # falla numérica anterior → "F169"
    ],
)
def test_arbol_de_descripcion(actual, anterior, esperado, fallback):  # R8.8
    r, _ = _eventos([[4], [3], [4]], fallas=[[anterior], [actual], ["NO FAULTS"]])
    (e,) = r.cerrados
    assert e.descripcion_falla == esperado and e.fallback is fallback


def test_codigo_falla_numerica():
    r, _ = _eventos([[4], [3], [4]], fallas=[["NO FAULTS"], [228.0], ["NO FAULTS"]])
    assert r.cerrados[0].codigo_falla == "F228"


def test_orden_de_suma_sin_excusable():  # (sumablocks + 4) - m
    r, _ = _eventos([[4], [0.1], [0.2], [4]], aplicar_evento_excusable_eventos=False, fe=0.0)
    esperado = ((0.0 + 4) - 0.1 + 4) - 0.2
    assert r.cerrados[0].suma_bloques == esperado


def test_excusable_eventos_pondera_con_d():  # L14 = "Yes"
    r, _ = _eventos([[4], [2], [4]], aplicar_evento_excusable_eventos=True, fe=0.0)
    assert r.cerrados[0].promedio_baterias == 0.0 and r.cerrados[0].horas_rack == 0.0


def test_filas_fuera_de_periodo_se_ignoran_e_inicio_por_anterior_fuera():
    antes = SERIAL_SEP_1 - BLOQUE  # 31-ago 23:45, fuera de L2
    seriales = [antes, SERIAL_SEP_1, SERIAL_SEP_1 + BLOQUE]
    r, _ = _eventos([[3], [3], [4]], seriales=seriales)
    (e,) = r.cerrados
    assert e.numero_bloques == 1 and e.serial_inicio == serial_min(-15)  # 31-ago 23:45


def test_cierre_estricto_y_arrastre_entre_pcs():  # F-06, Property 11
    # período = 01-sep; la fila siguiente a la última en rango es exactamente L4+1 (02-sep 00:00)
    fin = datetime_a_serial(date(2026, 9, 2))
    seriales = [fin - 2 * BLOQUE, fin - BLOQUE, fin]
    r, _ = _eventos([[4, 4], [3, 4], [3, 2]], seriales=seriales)
    # PCS 1: falla en 23:45 y sigue en 00:00 (fuera, = L4+1) → no cierra; PCS 2 no tiene falla en rango
    assert r.cerrados == [] and r.incompleto is not None and r.incompleto.numero_pcs == 1
    assert [a.tipo for a in r.anomalias] == ["evento_incompleto_excel"]

    seriales = [fin - 2 * BLOQUE, fin - BLOQUE, fin, fin + BLOQUE]
    r, _ = _eventos([[4, 4], [3, 2], [3, 4], [4, 4]], seriales=seriales)
    # PCS 2 hereda el bloque abierto de PCS 1 y sobrescribe su fila
    (e,) = r.cerrados
    assert e.arrastrado_excel and e.numero_pcs == 2 and e.pcs_iniciados == {1, 2}
    assert e.numero_bloques == 2 and e.suma_bloques == (0.0 + 4 - 3) + 4 - 2


def test_cierre_con_siguiente_mayor_que_l4_mas_uno():
    fin = datetime_a_serial(date(2026, 9, 2))
    r, _ = _eventos([[4], [3], [3]], seriales=[fin - 2 * BLOQUE, fin - BLOQUE, fin + BLOQUE])
    assert len(r.cerrados) == 1 and not r.cerrados[0].arrastrado_excel


def test_ultima_fila_consulta_la_fila_siguiente():
    r, _ = _eventos([[4], [3]], siguiente=[3.5])  # la fila donde corta el VBA trae 3.5: no cierra
    assert r.cerrados == [] and r.incompleto is not None
    r, _ = _eventos([[4], [3]])  # fila siguiente vacía: cierra
    assert len(r.cerrados) == 1


def test_evento_en_fila_2_excel_habria_fallado():  # F-11, D-08
    r, _ = _eventos([[3], [4]], primera_fila=2)
    (e,) = r.cerrados
    assert e.excel_habria_fallado and e.descripcion_falla == "F1 Watchdog" and e.fallback
    assert "excel_habria_fallado" in [a.tipo for a in r.anomalias]


def test_l8_l12():
    r, _ = _eventos([[4], [3], [4]])
    assert r.horas_periodo == 24.0
    assert r.disponibilidad_periodo == 1 - r.horas_rack_totales / (config_prueba(total_pcs=1).total_racks * 24.0)


def test_c23_indefinido():
    m = matriz([[3]])
    with pytest.raises(ValueError):
        detectar_eventos(m, actividad(1), config_prueba(total_pcs=1), None)


def test_resumen_por_codigo_sumif_insensible_a_mayusculas():
    r, _ = _eventos(
        [[4], [3], [4], [3], [4]], fallas=[["NO FAULTS"], ["f55 x"], ["NO FAULTS"], ["F55 Y"], ["NO FAULTS"]]
    )
    m = matriz(
        [[4], [3], [4], [3], [4]],
        fallas=[["NO FAULTS"], ["f55 x"], ["NO FAULTS"], ["F55 Y"], ["NO FAULTS"]],
        primera_fila=3,
    )
    r = detectar_eventos(
        m, actividad(5, primera_fila=3), config_prueba(total_pcs=1), C23, [("F0", "NO FAULT"), ("F55", "EXTERNAL")]
    )
    assert [(f.codigo, f.horas_rack) for f in r.resumen] == [("F55", r.horas_rack_totales), ("F0", 0.0)]
    assert r.resumen[0].porcentaje == 1.0

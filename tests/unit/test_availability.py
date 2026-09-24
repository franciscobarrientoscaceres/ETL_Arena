"""Tests de MotorDisponibilidad (R7)."""

import numpy as np
import pytest
from fabricas import BLOQUE, SERIAL_SEP_1, actividad, config_prueba, matriz

from etl_arena.availability import calcular


def test_tres_filas_dos_pcs_a_mano():
    # fila 1: PCS1=3 (1 batería), PCS2=4; fila 2: PCS1=4, PCS2=2.5 (1,5); fila 3: vacío y 0 (4)
    m = matriz([[3, 4], [4, 2.5], [None, 0]])
    cfg = config_prueba(aplicar_evento_excusable=False)
    r = calcular(m, actividad(3), cfg)
    assert r.bloques_muestreo == 3
    assert r.bloques_racks_indisponibles == 12 * 1 + 12 * 1.5 + 12 * 4
    assert r.disponibilidad_periodo == 1 - 78 / (cfg.total_racks * 3)
    assert r.disponibilidad_anual_acumulada == 1 - 78 / (cfg.total_racks * 35040)
    assert r.minutos_muestreo_derivado == 15.0
    assert np.isnan(r.ponderadas[0, 1]) and np.isnan(r.ponderadas[2, 0]) and r.ponderadas[2, 1] == 4.0


def test_filas_fuera_de_rango_no_cuentan():  # filtro A >= C5 y A < C7+1
    seriales = [SERIAL_SEP_1 - BLOQUE, SERIAL_SEP_1, SERIAL_SEP_1 + 1 - BLOQUE, SERIAL_SEP_1 + 1]
    m = matriz([[0, 0]] * 4, seriales=seriales)
    r = calcular(m, actividad(4), config_prueba())
    assert list(r.filas_procesadas) == [1, 2] and r.bloques_muestreo == 2


def test_excusable_solo_con_flag():  # R7.5, F-12
    m = matriz([[2, 4], [2, 4]])
    act = actividad(2, excusable=[0.0, 1.0])
    con = calcular(m, act, config_prueba(aplicar_evento_excusable=True))
    sin = calcular(m, act, config_prueba(aplicar_evento_excusable=False))
    assert con.bloques_racks_indisponibles == 24.0 and sin.bloques_racks_indisponibles == 48.0
    assert list(con.factor_excusable) == [0.0, 1.0]


def test_solo_tiempo_operacional_con_factor_cero():  # R7.6
    m = matriz([[3, 3], [3, 3]])
    act = actividad(2, operacional=[0.0, 1.0])
    r = calcular(m, act, config_prueba(solo_tiempo_operacional=True))
    assert r.bloques_racks_indisponibles == 24.0
    assert list(r.impacto[0]) == [0.0, 0.0] and list(r.ponderadas[0]) == [1.0, 1.0]  # la tabla no lleva el factor


def test_c23_derivado_en_salto_dst():  # F-13
    seriales = [46271.0, 46271.0 + 4 * BLOQUE, 46271.0 + 5 * BLOQUE]
    r = calcular(
        matriz([[4, 4]] * 3, seriales=seriales),
        actividad(3),
        config_prueba(
            inicio_periodo=__import__("datetime").date(2026, 9, 6), fin_periodo=__import__("datetime").date(2026, 9, 6)
        ),
    )
    assert r.minutos_muestreo_derivado == 60.0


def test_sin_filas_en_rango():
    r = calcular(matriz([[3, 3]], seriales=[SERIAL_SEP_1 + 5]), actividad(1), config_prueba())
    assert r.bloques_muestreo == 0 and r.disponibilidad_periodo is None and r.minutos_muestreo_derivado is None


def test_c2_menor_que_matriz_usa_solo_c2():  # For intRecPCS = 1 To C2
    r = calcular(matriz([[3, 0]]), actividad(1), config_prueba(total_pcs=1))
    assert r.bloques_racks_indisponibles == 12.0 and r.pcs == [1]


def test_c2_mayor_que_matriz_falla():
    with pytest.raises(ValueError):
        calcular(matriz([[3]]), actividad(1), config_prueba(total_pcs=2))


def test_muestras():
    r = calcular(matriz([[3, 4]]), actividad(1, excusable=0.5), config_prueba())
    muestras = list(r.muestras())
    assert len(muestras) == 2
    assert (
        muestras[0].baterias_indisponibles,
        muestras[0].baterias_ponderadas,
        muestras[0].impacto_rack_ponderado,
    ) == (1.0, 0.5, 6.0)
    assert muestras[1].impacto_rack_ponderado == 0.0

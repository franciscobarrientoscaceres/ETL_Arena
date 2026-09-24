"""Tests de enrichment.asociar_exclusion y de la regla de Exclusion_Matrix en el pipeline (F-37)."""

import numpy as np
import pytest
from fabricas import config_prueba, crear_libro, encabezado_exclusion, exclusion, fila_raw, matriz, serial_min

from etl_arena.enrichment import asociar_exclusion
from etl_arena.model import ErrorParidad
from etl_arena.pipeline import calcular_libro


def _filas_em(m, valores, extra=None):
    filas = {1: dict(enumerate(encabezado_exclusion(m.p), start=1))}
    for i, fila in enumerate(valores):
        celdas = {1: float(m.serial[i])}
        # el lector entrega números como float (bool y texto se pasan tal cual para probar el rechazo)
        celdas.update({j + 2: float(v) if type(v) is int else v for j, v in enumerate(fila) if v is not None})
        filas[int(m.numero_fila[i])] = celdas
    for fila, celdas in (extra or {}).items():
        filas.setdefault(fila, {}).update(celdas)
    return filas


class TestAsociarExclusion:
    def test_sin_hoja_todo_cero(self):
        m = matriz([[3, 3]])
        datos, anomalias = asociar_exclusion(m, None, config_prueba())
        assert not datos.valor.any() and np.isnan(datos.baterias_previas).all() and anomalias == []

    def test_valores_y_vacio(self):
        m = matriz([[3, 4], [2, 4]])
        datos = exclusion(m, [[1, None], [0, 2]])
        assert datos.valor.tolist() == [[1.0, 0.0], [0.0, 2.0]]

    @pytest.mark.parametrize("valor", [3.0, -1.0, 0.5, "CPF", True])
    def test_valor_invalido_rechazado(self, valor):
        m = matriz([[3, 3]])
        with pytest.raises(ErrorParidad) as e:
            asociar_exclusion(m, _filas_em(m, [[valor, None]]), config_prueba())
        assert e.value.anomalias[0].tipo == "em_valor_invalido"

    def test_encabezado_desplazado_rechazado(self):
        m = matriz([[3, 3]])
        filas = _filas_em(m, [[None, None]])
        filas[1][2], filas[1][3] = "PCS02", "PCS01"
        with pytest.raises(ErrorParidad):
            asociar_exclusion(m, filas, config_prueba())

    def test_desalineado_y_sin_timestamp(self):
        m = matriz([[3, 3], [3, 3]])
        filas = _filas_em(m, [[1, None], [1, None]])
        filas[2][1] = serial_min(999)  # fila 2 con otro timestamp
        del filas[3][1]  # fila 3 con marcas y sin timestamp
        _, anomalias = asociar_exclusion(m, filas, config_prueba())
        assert [(a.tipo, a.numero_fila) for a in anomalias] == [("em_desalineado", 2), ("em_sin_timestamp", 3)]

    def test_columnas_resumen_y_comentario(self):
        m = matriz([[3, 3]])
        filas = _filas_em(m, [[1, None]], extra={2: {4: 1.0, 5: "CPF"}})
        datos, _ = asociar_exclusion(m, filas, config_prueba())
        assert datos.evento_excusado[0] == 1.0 and datos.comentario[0] == "CPF"

    def test_previas_por_tramo(self):
        # PCS1: 3 módulos antes del tramo 2 → 1 batería; PCS2: completo antes → 0
        m = matriz([[3, 4], [0, 0], [1, 2], [4, 4], [2, 1], [0, 4]])
        datos = exclusion(m, [[0, 0], [2, 2], [2, 2], [0, 0], [0, 0], [2, 0]])
        assert datos.baterias_previas[1:3, 0].tolist() == [1.0, 1.0]
        assert datos.baterias_previas[1:3, 1].tolist() == [0.0, 0.0]
        assert datos.baterias_previas[5, 0] == 2.0  # segundo tramo: la fila previa tenía 2 módulos
        assert np.isnan(datos.baterias_previas[0, 0]) and np.isnan(datos.baterias_previas[3, 0])

    def test_tramo_desde_la_primera_fila(self):
        m = matriz([[1], [0]], primera_fila=2)
        filas = _filas_em(m, [[2], [2]])
        datos, anomalias = asociar_exclusion(m, filas, config_prueba(total_pcs=1))
        assert datos.baterias_previas[:, 0].tolist() == [3.0, 3.0]
        assert [a.tipo for a in anomalias] == ["ee2_sin_fila_previa"]

    def test_fila_sin_falla_dentro_del_tramo_se_reporta(self):  # la macro de agosto dejaría 0 después
        m = matriz([[3], [0], [4], [0]])
        filas = _filas_em(m, [[0], [2], [2], [2]])
        datos, anomalias = asociar_exclusion(m, filas, config_prueba(total_pcs=1))
        assert datos.baterias_previas[1:, 0].tolist() == [1.0, 1.0, 1.0]
        assert [a.tipo for a in anomalias] == ["ee2_difiere_macro_agosto"]


def test_pipeline_con_hoja(tmp_path):
    s = [serial_min(15 * k) for k in range(1, 6)]
    filas = [
        fila_raw(s[0], [4, 4]),
        fila_raw(s[1], [3, 4]),
        fila_raw(s[2], [0, 2]),
        fila_raw(s[3], [0, 2]),
        fila_raw(s[4], [4, 4]),
    ]
    em = {
        2: [s[0], 0, 0, None, None],
        3: [s[1], 0, 0, None, None],
        4: [s[2], 2, 1, 2, "CPF"],
        5: [s[3], 2, 0, 2, "CPF"],
        6: [s[4], 0, 0, None, None],
    }
    actividad = {r: [None, s[r - 2], 1, 1] for r in range(2, 7)}
    ruta = crear_libro(tmp_path / "l.xlsx", filas, actividad=actividad, exclusion=em)
    r = calcular_libro(ruta, config_prueba(), codigos_resumen=[])
    # PCS1: 1 + 1 (previa) + 1 (previa); PCS2: 0 (valor 1) + 2
    assert r.disponibilidad.bloques_racks_indisponibles == 12 * (1 + 1 + 1 + 0 + 2)
    assert r.exclusion.comentario.tolist() == [None, None, "CPF", "CPF", None]
    assert r.anomalias == []
    assert abs(r.disponibilidad.bloques_racks_indisponibles - 4 * r.eventos.horas_rack_totales) <= 1e-6

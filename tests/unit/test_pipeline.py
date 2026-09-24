"""Tests de pipeline.calcular_libro sobre libros sintéticos (sin SQL)."""

import pytest
from fabricas import config_prueba, crear_libro, encabezado_raw, fila_raw, serial_min

from etl_arena.model import ErrorParidad
from etl_arena.pipeline import calcular_libro


def test_libro_completo(tmp_path):
    s = [serial_min(15 * k) for k in range(1, 5)]
    filas = [
        fila_raw(s[0], [4, 4]),
        fila_raw(s[1], [3, 4], ["F55 X", "NO FAULTS"]),
        fila_raw(s[2], [2, 4], ["F55 X", "NO FAULTS"]),
        fila_raw(s[3], [4, 4]),
    ]
    actividad = {r: [None, s[r - 2], 1, 1] for r in range(2, 6)}
    ruta = crear_libro(tmp_path / "l.xlsx", filas, actividad=actividad)
    r = calcular_libro(ruta, config_prueba(), codigos_resumen=[("F55", "EXTERNAL")])
    assert r.disponibilidad.bloques_muestreo == 4
    assert r.disponibilidad.bloques_racks_indisponibles == 12 * 1 + 12 * 2
    assert [(e.numero_pcs, e.codigo_falla, e.numero_bloques) for e in r.eventos.cerrados] == [(1, "F55", 2)]
    assert r.eventos.resumen[0].horas_rack == r.eventos.horas_rack_totales
    assert [d.diario for d in r.diario] == [36.0]
    assert [a.tipo for a in r.anomalias] == ["exclusion_matrix_ausente"]  # libro sin la hoja (F-37)
    assert not r.exclusion.valor.any()


def test_c23_distinto_se_reporta(tmp_path):
    s = [serial_min(15), serial_min(45)]
    ruta = crear_libro(tmp_path / "l.xlsx", [fila_raw(s[0], [4, 4]), fila_raw(s[1], [4, 4])])
    r = calcular_libro(ruta, config_prueba(), codigos_resumen=[])
    assert "c23_distinto" in [a.tipo for a in r.anomalias]


def test_encabezado_invalido_aborta(tmp_path):
    enc = encabezado_raw(2)
    enc[4] = "Arena - PCS 01 - POWERELECTRONICS HEM-k OTRA COSA"
    ruta = crear_libro(tmp_path / "l.xlsx", [fila_raw(serial_min(15), [4, 4])], encabezado=enc)
    with pytest.raises(ErrorParidad):
        calcular_libro(ruta, config_prueba(), codigos_resumen=[])

"""Tests de la reconciliación (R13, tarea 3.3)."""

import dataclasses

import pytest
from fabricas import BLOQUE, config_prueba, crear_libro, fila_raw, referencia_desde_resultado, serial_min

from etl_arena.pipeline import calcular_libro
from etl_arena.reconciliation import reconciliar
from etl_arena.reconciliation.servicio import a_filas, resumen


@pytest.fixture(scope="module")
def resultado(tmp_path_factory):
    s = [serial_min(15 * k) for k in range(1, 6)]
    filas = [
        fila_raw(s[0], [4, 4]),
        fila_raw(s[1], [3, 2], ["F55 X", 169.0]),
        fila_raw(s[2], [2.5, 2], ["F55 X", 169.0]),
        fila_raw(s[3], [4, 4]),
        fila_raw(s[4], [4, 4]),
    ]
    ruta = crear_libro(
        tmp_path_factory.mktemp("rec") / "l.xlsx", filas, actividad={r: [None, s[r - 2], 1, 1] for r in range(2, 7)}
    )
    cfg = config_prueba(aplicar_evento_excusable=True, aplicar_evento_excusable_eventos=True)
    return calcular_libro(ruta, cfg, codigos_resumen=[("F55", "EXT"), ("F169", "X")])


def test_datos_identicos_pasan_todo(resultado):
    rep = reconciliar(resultado, referencia_desde_resultado(resultado))
    assert rep.aprobado_global and rep.estado_corrida == "success"
    assert all(n.aprobado and n.comparaciones > 0 for n in rep.niveles)
    assert rep.invariantes.aprobado and not rep.invariantes.no_aplica  # C14 = 4·L10 aplica y se cumple
    assert resumen(rep)["estado"] == "pass"


def test_c14_alterado_falla_nivel_3(resultado):
    ref = referencia_desde_resultado(resultado)
    ref.c14 += 1e-3
    rep = reconciliar(resultado, ref)
    assert not rep.nivel(3).aprobado and rep.estado_corrida == "parity_failed"
    (d,) = rep.nivel(3).discrepancias
    assert d.metrica == "C14" and d.delta == pytest.approx(1e-3)


def test_c14_dentro_de_tolerancia_pasa(resultado):
    ref = referencia_desde_resultado(resultado)
    ref.c14 += 1e-9
    assert reconciliar(resultado, ref).nivel(3).aprobado


def test_evento_con_fin_15_min_despues_falla_nivel_5(resultado):
    ref = referencia_desde_resultado(resultado)
    ref.eventos[0] = dataclasses.replace(ref.eventos[0], d=ref.eventos[0].d + BLOQUE)
    rep = reconciliar(resultado, ref)
    assert not rep.nivel(5).aprobado and rep.estado_corrida == "parity_failed"
    assert [d.metrica for d in rep.nivel(5).discrepancias] == ["D"]


def test_celda_de_tabla_distinta_falla_solo_nivel_2(resultado):
    ref = referencia_desde_resultado(resultado)
    fila, serial, pcs, valor = ref.tabla[0]
    ref.tabla[0] = (fila, serial, pcs, valor + 0.5)
    rep = reconciliar(resultado, ref)
    assert not rep.nivel(2).aprobado and rep.estado_corrida == "success"  # R13.9: solo 3–5 deciden


def test_invariante_no_aplica_si_l14_distinto_de_c31(resultado):
    r = dataclasses.replace(resultado, cfg=dataclasses.replace(resultado.cfg, aplicar_evento_excusable_eventos=False))
    rep = reconciliar(r, None)
    assert rep.invariantes.no_aplica == ["C31 ≠ L14"] and rep.invariantes.comparaciones == 0


def test_sin_referencia(resultado):
    rep = reconciliar(resultado, None)
    assert rep.sin_referencia and rep.estado_corrida == "success" and rep.niveles == []
    assert resumen(rep) == {"estado": "sin_referencia"}
    assert a_filas(rep)[0].metrica == "sin_referencia"


def test_filas_para_sql(resultado):
    ref = referencia_desde_resultado(resultado)
    ref.c16 = 0.5
    filas = a_filas(reconciliar(resultado, ref))
    resumenes = [f for f in filas if f.metrica == "resumen"]
    assert [f.nivel for f in resumenes] == [1, 2, 3, 4, 5, 9]
    (falla,) = [f for f in filas if not f.aprobado and f.metrica != "resumen"]
    assert (falla.nivel, falla.metrica, falla.valor_excel) == (4, "C16", "0.5")

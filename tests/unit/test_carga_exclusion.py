"""Carga mensual de la Exclusion_Matrix sin Excel (4.12; R3.8, R19.6, D-13, D-17, F-37)."""

import csv
import importlib.util
import json
from datetime import date
from pathlib import Path

import pytest
from fabricas import crear_libro, fila_raw, serial_min

from etl_arena.workbook import exclusion as em

S = [serial_min(m) for m in (-30, -15, 0, 15)]  # 2026-08-31 23:30, 23:45, 2026-09-01 00:00, 00:15


@pytest.fixture
def libro(tmp_path):
    filas = [fila_raw(s, [3.0, 4.0]) for s in S]
    return crear_libro(tmp_path / "libro.xlsx", filas, actividad={r: [None, S[r - 2], 1, 1] for r in range(2, 6)})


def _entrega(tmp_path, filas, nombre="em.xlsx", total_pcs=2):
    return crear_libro(tmp_path / nombre, [fila_raw(S[0], [4.0] * total_pcs)], total_pcs=total_pcs, exclusion=filas)


def test_leer_entrega_toma_solo_el_mes(tmp_path):
    ruta = _entrega(tmp_path, {2: [S[0], 1.0, None, 1.0, "falla red"], 3: [S[1], 0.0, 2.0], 4: [S[2], 1.0, 1.0]})
    e = em.leer_entrega(ruta, 2026, 8, 2)
    assert [f.serial for f in e.filas] == S[:2] and e.filas_fuera_del_mes == 1
    assert e.filas[0].valores == (1.0, None) and e.filas[0].comentario == "falla red"
    assert len(e.sha256) == 64


@pytest.mark.parametrize(
    ("filas", "mensaje"),
    [
        ({2: [S[0], 3.0]}, "se admite 0, 1, 2"),
        ({2: ["2026-08-31 23:30", 1.0]}, "no es una fecha"),
        ({2: [None, 1.0]}, "sin Date/time"),
        ({2: [serial_min(0, date(2026, 7, 1)), 1.0]}, "no hay filas de 2026-08"),
    ],
)
def test_leer_entrega_rechaza(tmp_path, filas, mensaje):
    with pytest.raises(em.ErrorCargaExclusion, match=mensaje):
        em.leer_entrega(_entrega(tmp_path, filas), 2026, 8, 2)


def test_leer_entrega_valida_el_encabezado(tmp_path):
    ruta = crear_libro(tmp_path / "em.xlsx", [], total_pcs=3, exclusion={2: [S[0], 1.0]})
    with pytest.raises(em.ErrorCargaExclusion, match="encabezado"):
        em.leer_entrega(ruta, 2026, 8, 2)


def test_planificar_ubica_por_timestamp_y_calcula_el_diff(tmp_path, libro):
    # el libro ya trae una marca (fila 2, PCS02 = 1) que la entrega borra
    base = crear_libro(
        tmp_path / "base.xlsx",
        [fila_raw(s, [3.0, 4.0]) for s in S],
        exclusion={2: [S[0], None, 1.0], 3: [S[1], 0.0, None]},
    )
    e = em.leer_entrega(_entrega(tmp_path, {2: [S[1], 1.0, 0.0, 1.0, "x"], 3: [S[0], 0.0, None]}), 2026, 8, 2)
    plan = em.planificar(base, e, 2)
    assert plan.hoja_existe and sorted(plan.filas) == [2, 3]
    assert plan.filas[3][:3] == [S[1], 1.0, 0.0]  # 23:45 → fila 3 aunque venga primero en la entrega
    cambios = {(c.fila, c.campo): (c.anterior, c.nuevo) for c in plan.cambios}
    assert cambios == {
        (2, "PCS02"): (1.0, None),
        (3, "PCS01"): (0.0, 1.0),
        (3, "Excused Event"): (None, 1.0),
        (3, "Comments"): (None, "x"),
    }  # 0 ↔ vacío no es cambio
    out = em.escribir_cambios_csv(plan, tmp_path / "cambios.csv")
    lineas = list(csv.reader(out.read_text(encoding="utf-8-sig").splitlines(), delimiter=";"))
    assert lineas[0][0] == "Hoja" and len(lineas) == 5


def test_planificar_rechaza_timestamps_que_no_existen(tmp_path, libro):
    otro = serial_min(-60)  # 23:00: no está en RawData-PCS
    e = em.leer_entrega(_entrega(tmp_path, {2: [otro, 1.0]}), 2026, 8, 2)
    with pytest.raises(em.ErrorCargaExclusion, match="no existen en RawData-PCS"):
        em.planificar(libro, e, 2)


def test_planificar_sin_hoja_la_crea_y_todo_es_cambio(tmp_path, libro):
    e = em.leer_entrega(_entrega(tmp_path, {2: [S[0], 1.0, None]}), 2026, 8, 2)
    plan = em.planificar(libro, e, 2)
    assert not plan.hoja_existe and [(c.fila, c.campo) for c in plan.cambios] == [(2, "PCS01")]


def test_exclusiones_en_periodo(tmp_path):
    con = crear_libro(tmp_path / "c.xlsx", [fila_raw(s, [4.0, 4.0]) for s in S], exclusion={3: [S[1], 0.0, 2.0]})
    assert em.exclusiones_en_periodo(con, date(2026, 8, 1), date(2026, 8, 31), 2)
    assert not em.exclusiones_en_periodo(con, date(2026, 9, 1), date(2026, 9, 30), 2)
    sin = crear_libro(tmp_path / "s.xlsx", [fila_raw(s, [4.0, 4.0]) for s in S])
    assert not em.exclusiones_en_periodo(sin, date(2026, 8, 1), date(2026, 8, 31), 2)


def test_bloques_contiguos():
    assert em._bloques([2, 3, 4, 7, 9, 10]) == [(2, 4), (7, 7), (9, 10)]


# ------------------------------------------------------------------ run_lunes --stage load-exclusion-matrix
def _run_lunes():
    ruta = Path(__file__).parents[2] / "scripts" / "run_lunes.py"
    spec = importlib.util.spec_from_file_location("run_lunes", ruta)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def test_load_exclusion_matrix_cierra_el_mes_con_exclusiones(tmp_path, monkeypatch, capsys):
    """Sin Excel: el escritor COM se reemplaza por uno que reescribe el xlsx con la matriz del plan."""
    mods = [4.0] * 61
    filas = [fila_raw(S[0], [3.0, *mods[1:]]), fila_raw(S[1], mods), fila_raw(S[2], mods), fila_raw(S[3], mods)]
    actividad = {r: [None, S[r - 2], 1, 1] for r in range(2, 6)}
    maestro = crear_libro(tmp_path / "maestro.xlsx", filas, total_pcs=61, actividad=actividad)
    entrega = crear_libro(
        tmp_path / "em.xlsx", [], total_pcs=61, exclusion={2: [S[0], 1.0, *[None] * 60, 1.0, "corte de red"]}
    )

    def escribir_falso(libro, plan, total_pcs, timeout_s=0):
        crear_libro(libro, filas, total_pcs=61, actividad=actividad, exclusion=plan.filas)

    monkeypatch.setattr(em, "aplicar_plan", escribir_falso)
    rl = _run_lunes()
    work = tmp_path / "work"
    args = [
        "--stage",
        "load-exclusion-matrix",
        "--work",
        str(work),
        "--maestro",
        str(maestro),
        "--sin-bd",
        "--omitir-macros",
        "--mes",
        "2026-08",
        "--archivo-matriz",
        str(entrega),
    ]
    assert rl.main(args) == 0
    estado = json.loads((work / "matriz-2026-08" / "run_state.json").read_text(encoding="utf-8"))
    assert estado["write-exclusion-matrix"]["artefactos"]["cambios"] == 3  # PCS01, Excused Event, Comments
    assert estado["write-exclusion-matrix"]["artefactos"]["hoja_creada"]
    etl = estado["run-etl"]["artefactos"]
    assert etl["estado_exclusiones"] == "con_exclusiones" and etl["periodo"] == ["2026-08-01", "2026-08-31"]
    assert etl["kpi"]["C14"] == 0.0  # la única falla quedó excusada (valor 1)
    assert (work / "matriz-2026-08" / "cambios.csv").exists()
    assert json.loads((work / "cola_cierres.json").read_text(encoding="utf-8"))[0]["estado_exclusiones"] == (
        "con_exclusiones"
    )


def test_load_exclusion_matrix_valida_argumentos(tmp_path, capsys):
    rl = _run_lunes()
    assert rl.main(["--stage", "load-exclusion-matrix", "--work", str(tmp_path), "--mes", "2026-08"]) == 1
    assert "--archivo-matriz" in capsys.readouterr().err

"""Orquestador semanal (4.9) sin Excel: etapas, estado reanudable y período por defecto (D-07)."""

import importlib.util
import json
from datetime import date
from pathlib import Path

import pytest
from fabricas import crear_libro, fila_raw, serial_min


def _modulo():
    ruta = Path(__file__).parents[2] / "scripts" / "run_lunes.py"
    spec = importlib.util.spec_from_file_location("run_lunes", ruta)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


@pytest.fixture
def libro61(tmp_path):
    s = [serial_min(15 * k) for k in range(1, 5)]  # 01-09-2026 00:15 … 01:00
    mods = [4.0] * 61
    filas = [fila_raw(s[0], mods), fila_raw(s[1], [3.0, *mods[1:]]), fila_raw(s[2], mods), fila_raw(s[3], mods)]
    return crear_libro(
        tmp_path / "preparado.xlsx", filas, total_pcs=61, actividad={r: [None, s[r - 2], 1, 1] for r in range(2, 6)}
    )


def test_etapas_sin_excel(tmp_path, libro61, capsys):
    rl = _modulo()
    base = ["--corte", "2026-09-01", "--work", str(tmp_path / "work"), "--sin-bd"]
    assert rl.main(["--stage", "prepare-workbook", "--libro-preparado", str(libro61), *base]) == 0
    libro = tmp_path / "work" / "2026-09-01" / "libro.xlsm"
    assert libro.exists() and libro.with_suffix(".xlsm.bak").exists()
    assert rl.main(["--stage", "run-etl", *base]) == 0
    assert rl.main(["--stage", "reconcile", *base]) == 0
    assert rl.main(["--stage", "notify-bi", *base]) == 0
    estado = json.loads((tmp_path / "work" / "2026-09-01" / "run_state.json").read_text(encoding="utf-8"))
    etl = estado["run-etl"]["artefactos"]
    assert etl["estado"] == "success" and etl["periodo"] == ["2026-09-01", "2026-09-01"]  # D-07: mes en curso
    assert etl["kpi"]["C14"] == 12.0 and etl["estado_exclusiones"] == "sin_exclusiones"
    assert estado["reconcile"]["artefactos"]["reconciliacion"] == "sin_referencia"
    assert (tmp_path / "work" / "2026-09-01" / "notificacion.md").exists()
    assert all(estado[e]["estado"] == "ok" for e in ("prepare-workbook", "run-etl", "reconcile", "notify-bi"))


def test_prepare_sin_libro_preparado_explica_que_falta(tmp_path, capsys):
    rl = _modulo()
    codigo = rl.main(["--stage", "prepare-workbook", "--corte", "x", "--work", str(tmp_path)])
    assert codigo == 1
    estado = json.loads((tmp_path / "x" / "run_state.json").read_text(encoding="utf-8"))
    assert "0.6/0.7" in estado["prepare-workbook"]["error"]


def test_reconcile_detiene_la_cadena_si_parity_failed(tmp_path, capsys):
    rl = _modulo()
    dir_corte = tmp_path / "c"
    dir_corte.mkdir()
    (dir_corte / "run_state.json").write_text(
        json.dumps({"run-etl": {"estado": "ok", "artefactos": {"estado": "parity_failed", "id_corrida": "abc"}}}),
        encoding="utf-8",
    )
    assert rl.main(["--stage", "reconcile", "--corte", "c", "--work", str(tmp_path)]) == 2


def test_ultimo_dato(libro61):
    assert _modulo().ultimo_dato(libro61) == date(2026, 9, 1)

"""Registro del shadow mode (5.4): fila desde run_state.json y criterio de salida."""

import importlib.util
import json
import shutil
from pathlib import Path

import pytest

RAIZ = Path(__file__).parents[2]


def _modulo():
    spec = importlib.util.spec_from_file_location("shadow_log", RAIZ / "scripts" / "shadow_log.py")
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def _corte(work: Path, corte: str, *, tipo="semanal", rec="pass", estado="success", excl="sin_exclusiones"):
    d = work / corte
    d.mkdir(parents=True)
    etl = {
        "id_corrida": f"id-{corte}",
        "num_corrida": 5,
        "estado": estado,
        "tipo": tipo,
        "periodo": ["2026-09-01", "2026-09-21"],
        "estado_exclusiones": excl,
        "kpi": {"C12": 1975, "C14": 104134.292, "C16": 0.9819924099052362},
        "reconciliacion": {"estado": rec},
    }
    macros = {"C14": 104134.292, "referencia": True}
    (d / "run_state.json").write_text(
        json.dumps(
            {"run-etl": {"estado": "ok", "artefactos": etl}, "run-macros": {"estado": "ok", "artefactos": macros}}
        ),
        encoding="utf-8",
    )


@pytest.fixture
def log(tmp_path):
    destino = tmp_path / "shadow-log.md"
    shutil.copy(RAIZ / "docs" / "shadow-log.md", destino)
    return destino


def test_agrega_una_fila_con_delta_contra_el_kpi_de_alex(tmp_path, log, capsys):
    sl = _modulo()
    _corte(tmp_path, "2026-09-28")
    args = ["--corte", "2026-09-28", "--work", str(tmp_path), "--log", str(log), "--kpi-excel-oficial", "0.98199"]
    assert sl.main([*args, "--nota", "ok | sin novedades"]) == 0
    filas = sl.filas_registradas(log)
    assert len(filas) == 1
    f = filas[0]
    assert (f["Corte"], f["Tipo"], f["Reconciliación"], f["C12"], f["KPI Alex"]) == (
        "2026-09-28",
        "semanal",
        "pass",
        "1975",
        "0.981990",
    )
    assert f["Δ"] == "2.41e-06" and f["Nota"] == "ok / sin novedades"
    assert f["N° corrida"] == "5" and f["IdCorrida"] == "`id-2026-09-28`"


def test_corte_sin_run_etl_ok_falla(tmp_path, log, capsys):
    assert _modulo().main(["--corte", "x", "--work", str(tmp_path), "--log", str(log)]) == 1


def test_resumen_exige_4_semanas_seguidas_y_un_cierre(tmp_path, log, capsys):
    sl = _modulo()
    cortes = [("s1", {}), ("s2", {"rec": "fail"}), ("s3", {}), ("s4", {}), ("s5", {}), ("s6", {})]
    for corte, kw in cortes:
        _corte(tmp_path, corte, **kw)
        assert sl.main(["--corte", corte, "--work", str(tmp_path), "--log", str(log)]) == 0
    r = sl.evaluar(sl.filas_registradas(log))
    assert r["mejor_racha_semanal"] == 4 and not r["cumple"]  # s2 falló sin nota: la racha parte en s3
    _corte(tmp_path, "cierre-2026-09", tipo="cierre_mensual", rec="sin_referencia", excl="con_exclusiones")
    assert sl.main(["--corte", "cierre-2026-09", "--work", str(tmp_path), "--log", str(log)]) == 0
    assert sl.main(["--resumen", "--log", str(log)]) == 0
    r = sl.evaluar(sl.filas_registradas(log))
    assert r["cierres_con_exclusiones"] == 1 and r["cumple"]

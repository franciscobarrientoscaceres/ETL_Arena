"""Libro base y cola de cierres entre cortes (4.9; R3.1, R19.3, R19.6, D-17)."""

import importlib.util
import json
from datetime import date
from pathlib import Path

import pytest
from fabricas import crear_libro, fila_raw, serial_min

from etl_arena.orquestacion import (
    ColaCierres,
    encolar_cierres,
    libro_base,
    mes_cubierto,
    meses_candidatos,
    promover_libro_base,
)

ULTIMO_AGOSTO = serial_min(-15)  # 2026-08-31 23:45


def test_mes_cubierto_exige_el_ultimo_bloque():
    assert mes_cubierto(ULTIMO_AGOSTO, 2026, 8, 15)
    assert not mes_cubierto(serial_min(-30), 2026, 8, 15)  # 23:30
    assert not mes_cubierto(serial_min(60 * 24 * 20), 2026, 9, 15)


def test_meses_candidatos_cruzan_el_cambio_de_anio():
    assert meses_candidatos(date(2027, 1, 4)) == [(2026, 12), (2027, 1)]


def test_encolar_filtra_por_cobertura_desde_y_bd(tmp_path):
    cola = ColaCierres(tmp_path)
    kw = {"minutos_muestreo": 15, "corte": "c1", "desde": date(2026, 7, 1)}
    assert encolar_cierres(cola, serial_min(60), date(2026, 9, 1), **kw) == ["2026-08"]
    assert encolar_cierres(cola, serial_min(60), date(2026, 9, 1), **kw) == []  # ya estaba
    assert (
        encolar_cierres(
            ColaCierres(tmp_path / "x"), serial_min(60), date(2026, 9, 1), **kw, mes_cerrado=lambda a, m: True
        )
        == []
    )
    kw["desde"] = date(2026, 9, 1)
    assert encolar_cierres(ColaCierres(tmp_path / "y"), serial_min(60), date(2026, 9, 1), **kw) == []
    assert cola.pendientes() == ["2026-08"]
    cola.marcar_ejecutado("2026-08", "abc", "sin_exclusiones")
    assert cola.pendientes() == [] and cola.entradas()[0]["id_corrida"] == "abc"


def test_libro_base_es_el_maestro_hasta_la_primera_promocion(tmp_path):
    maestro = tmp_path / "maestro.xlsm"
    maestro.write_bytes(b"m")
    work = tmp_path / "work"
    assert libro_base(work, maestro) == maestro
    trabajo = work / "c1" / "libro.xlsm"
    trabajo.parent.mkdir(parents=True)
    trabajo.write_bytes(b"t")
    promover_libro_base(work, trabajo, "c1", "id1", "h")
    assert libro_base(work, maestro) == trabajo.resolve()
    trabajo.unlink()
    with pytest.raises(FileNotFoundError, match="c1"):
        libro_base(work, maestro)


# ------------------------------------------------------------------ run_lunes (sin Excel ni BD)
def _run_lunes():
    ruta = Path(__file__).parents[2] / "scripts" / "run_lunes.py"
    spec = importlib.util.spec_from_file_location("run_lunes", ruta)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


@pytest.fixture
def libro_cruza_mes(tmp_path):
    """Últimos bloques de agosto y primeros de septiembre: agosto queda completo."""
    s = [serial_min(m) for m in (-30, -15, 0, 15)]
    mods = [4.0] * 61
    filas = [fila_raw(s[0], [3.0, *mods[1:]]), fila_raw(s[1], mods), fila_raw(s[2], mods), fila_raw(s[3], mods)]
    return crear_libro(
        tmp_path / "base.xlsx", filas, total_pcs=61, actividad={r: [None, s[r - 2], 1, 1] for r in range(2, 6)}
    )


def test_semanal_encola_el_mes_completo_y_lo_notifica(tmp_path, libro_cruza_mes, capsys):
    rl = _run_lunes()
    work = tmp_path / "work"
    base = ["--corte", "2026-09-07", "--work", str(work), "--sin-bd"]
    assert (
        rl.main(
            ["--stage", "all", "--omitir-acquire", "--omitir-macros", "--libro-preparado", str(libro_cruza_mes), *base]
        )
        == 0
    )
    estado = json.loads((work / "2026-09-07" / "run_state.json").read_text(encoding="utf-8"))
    assert estado["run-etl"]["artefactos"]["cierres_encolados"] == ["2026-08"]
    assert "libro_base" not in estado["reconcile"]["artefactos"]  # sin BD no se promueve
    assert "2026-08" in (work / "2026-09-07" / "notificacion.md").read_text(encoding="utf-8")


def test_cierre_sin_matriz_exige_sin_exclusiones(tmp_path, libro_cruza_mes, capsys):
    rl = _run_lunes()
    work = tmp_path / "work"
    ColaCierres(work).encolar("2026-08", "c")
    args = [
        "--stage",
        "cierre-mensual",
        "--work",
        str(work),
        "--maestro",
        str(libro_cruza_mes),
        "--sin-bd",
        "--omitir-macros",
    ]
    assert rl.main(args) == 1
    assert "Exclusion_Matrix" in capsys.readouterr().err
    assert ColaCierres(work).pendientes() == ["2026-08"]

    assert rl.main([*args, "--sin-exclusiones"]) == 0
    estado = json.loads((work / "cierre-2026-08" / "run_state.json").read_text(encoding="utf-8"))
    etl = estado["run-etl"]["artefactos"]
    assert etl["periodo"] == ["2026-08-01", "2026-08-31"] and etl["estado_exclusiones"] == "sin_exclusiones"
    assert etl["kpi"]["C12"] == 2 and etl["kpi"]["C14"] == 12.0
    assert ColaCierres(work).pendientes() == []


def test_cierre_rechaza_un_mes_no_cubierto(tmp_path, libro_cruza_mes, capsys):
    rl = _run_lunes()
    args = [
        "--stage",
        "cierre-mensual",
        "--mes",
        "2026-09",
        "--work",
        str(tmp_path / "w"),
        "--maestro",
        str(libro_cruza_mes),
        "--sin-bd",
        "--sin-exclusiones",
        "--omitir-macros",
    ]
    assert rl.main(args) == 1
    assert "no cubre 2026-09" in capsys.readouterr().err

"""Entorno de prueba: otra base, otras carpetas, sin avisar a Misael (validar antes de cargar lo definitivo)."""

import importlib.util
import json
from pathlib import Path

import pytest
from fabricas import crear_libro, fila_raw, serial_min

from etl_arena.acquisition import acquire_wait
from etl_arena.persistence import conexion

PROD = "mssql+pyodbc://@srv.database.windows.net:1433/trina_etl?driver=ODBC+Driver+18+for+SQL+Server"
PRUEBA = PROD.replace("/trina_etl?", "/trina_etl_prueba?")


@pytest.fixture
def env(monkeypatch):
    for v in (conexion.VARIABLE_URL, conexion.VARIABLE_URL_PRUEBA, conexion.VARIABLE_ENTORNO):
        monkeypatch.delenv(v, raising=False)
    monkeypatch.setattr(conexion, "cargar_env", lambda *a, **k: {})
    monkeypatch.setenv(conexion.VARIABLE_URL, PROD)
    # run_lunes escribe ETL_ARENA_ENTORNO en os.environ: setenv la registra para restaurarla al terminar
    monkeypatch.setenv(conexion.VARIABLE_ENTORNO, "")
    return monkeypatch


def test_por_defecto_produccion(env):
    assert conexion.entorno() == "produccion" and conexion.url_configurada().database == "trina_etl"


def test_prueba_usa_su_propia_base(env):
    env.setenv(conexion.VARIABLE_ENTORNO, "prueba")
    env.setenv(conexion.VARIABLE_URL_PRUEBA, PRUEBA)
    assert conexion.url_configurada().database == "trina_etl_prueba"


@pytest.mark.parametrize(
    ("url_prueba", "mensaje"), [(None, "falta ETL_ARENA_DB_URL_PRUEBA"), (PROD, "apunta a la base de producción")]
)
def test_prueba_nunca_cae_en_produccion(env, url_prueba, mensaje):
    env.setenv(conexion.VARIABLE_ENTORNO, "prueba")
    if url_prueba:
        env.setenv(conexion.VARIABLE_URL_PRUEBA, url_prueba)
    with pytest.raises(conexion.ErrorConexion, match=mensaje):
        conexion.url_configurada()


def test_entorno_invalido(env):
    env.setenv(conexion.VARIABLE_ENTORNO, "qa")
    with pytest.raises(conexion.ErrorConexion, match="se admite"):
        conexion.entorno()


def test_acquire_en_prueba_copia_y_deja_el_export(tmp_path):
    inbox = tmp_path / "inbox"
    inbox.mkdir()
    (inbox / "raw_pcs_2026-09-28.csv").write_text("a;b\n", encoding="utf-8")
    a = acquire_wait(inbox, tmp_path / "processed", "c", timeout_s=0, poll_s=0, estabilidad_s=0, mover=False)
    assert a.ruta.exists() and (inbox / "raw_pcs_2026-09-28.csv").exists()


def _run_lunes():
    ruta = Path(__file__).parents[2] / "scripts" / "run_lunes.py"
    spec = importlib.util.spec_from_file_location("run_lunes", ruta)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def test_run_lunes_en_prueba_separa_carpetas_y_no_avisa(tmp_path, env, capsys):
    s = [serial_min(15 * k) for k in range(1, 5)]
    libro = crear_libro(
        tmp_path / "l.xlsx",
        [fila_raw(x, [4.0] * 61) for x in s],
        total_pcs=61,
        actividad={r: [None, s[r - 2], 1, 1] for r in range(2, 6)},
    )
    env.setenv("ETL_ARENA_NOTIFY_WEBHOOK", "http://127.0.0.1:9/no-debe-usarse")
    rl = _run_lunes()
    defecto = rl.construir_parser().parse_args(["--stage", "all"])
    args = rl.construir_parser().parse_args(["--stage", "all", "--entorno", "prueba"])
    rl.aplicar_entorno(args)
    assert args.work == defecto.work / "_prueba" and args.processed == defecto.processed / "_prueba"

    work = tmp_path / "w"  # --work explícito se respeta
    comando = [
        "--stage",
        "all",
        "--entorno",
        "prueba",
        "--omitir-acquire",
        "--omitir-macros",
        "--sin-bd",
        "--corte",
        "c",
        "--work",
        str(work),
        "--libro-preparado",
        str(libro),
    ]
    assert rl.main(comando) == 0
    estado = json.loads((work / "c" / "run_state.json").read_text(encoding="utf-8"))
    assert estado["run-etl"]["artefactos"]["entorno"] == "prueba"
    assert estado["notify-bi"]["artefactos"]["destino"].endswith("notificacion.md")
    assert (work / "c" / "notificacion.md").read_text(encoding="utf-8").startswith("# [PRUEBA]")

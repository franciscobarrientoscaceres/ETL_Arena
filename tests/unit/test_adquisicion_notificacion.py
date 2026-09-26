"""Tests de acquire-wait (4.5) y de la notificación a Power BI (4.8)."""

import http.server
import json
import threading

import pytest

from etl_arena.acquisition import ErrorAdquisicion, acquire_wait, sha256_de
from etl_arena.reporting.notificacion import construir_mensaje, enviar


def _adquirir(inbox, tmp_path, **kw):
    return acquire_wait(
        inbox, tmp_path / "processed", "2026-09-28", timeout_s=kw.pop("timeout_s", 1), poll_s=0.1, estabilidad_s=0, **kw
    )


class TestAcquireWait:
    def test_mueve_y_registra_sha256(self, tmp_path):
        inbox = tmp_path / "inbox"
        inbox.mkdir()
        archivo = inbox / "raw_pcs_2026-09-28.csv"
        archivo.write_bytes(b"a,b\n1,2\n")
        esperado = sha256_de(archivo)
        r = _adquirir(inbox, tmp_path)
        assert not archivo.exists() and r.ruta == tmp_path / "processed" / "2026-09-28" / archivo.name
        assert r.sha256 == esperado and r.bytes == 8
        assert (r.ruta.parent / f"{archivo.name}.sha256").read_text(encoding="ascii").startswith(esperado)

    def test_timeout_con_mensaje_accionable(self, tmp_path):
        (tmp_path / "inbox").mkdir()
        with pytest.raises(ErrorAdquisicion, match="TeamViewer"):
            _adquirir(tmp_path / "inbox", tmp_path, timeout_s=0.2)

    def test_vacio_rechazado(self, tmp_path):
        inbox = tmp_path / "inbox"
        inbox.mkdir()
        (inbox / "raw_pcs_x.csv").write_bytes(b"")
        with pytest.raises(ErrorAdquisicion, match="vacío"):
            _adquirir(inbox, tmp_path)

    def test_varios_candidatos_no_adivina(self, tmp_path):
        inbox = tmp_path / "inbox"
        inbox.mkdir()
        for n in ("raw_pcs_a.csv", "raw_pcs_b.csv"):
            (inbox / n).write_bytes(b"x")
        with pytest.raises(ErrorAdquisicion, match="dejar solo"):
            _adquirir(inbox, tmp_path)

    def test_ignora_parciales_y_no_pisa_procesados(self, tmp_path):
        inbox = tmp_path / "inbox"
        inbox.mkdir()
        (inbox / "raw_pcs_a.csv.part").write_bytes(b"x")
        (inbox / "raw_pcs_a.csv").write_bytes(b"x")
        _adquirir(inbox, tmp_path)
        (inbox / "raw_pcs_a.csv").write_bytes(b"y")
        with pytest.raises(ErrorAdquisicion, match="inmutable"):
            _adquirir(inbox, tmp_path)


MENSAJE = construir_mensaje(
    "abc",
    ("2026-09-01", "2026-09-28"),
    "success",
    "sin_exclusiones",
    {"C12": 2600, "C14": 1.5, "C16": 0.98},
    {"estado": "pass"},
    {"anomalias": {"total": 3, "por_severidad": {"info": 3}}},
    publicacion={"mes": "2026-09", "filas_borradas": {}, "filas_insertadas": {}, "detenciones": {}, "correcciones": 0},
)


def test_mensaje():
    assert MENSAJE["exclusiones"] == "Sin Exclusiones" and MENSAJE["reconciliacion"] == "pass"
    assert "Actualizar" in MENSAJE["accion"]
    falla = construir_mensaje("x", ("a", "b"), "parity_failed", "con_exclusiones", {}, None, None)
    assert falla["exclusiones"] == "Con Exclusiones" and "No actualizar" in falla["accion"]
    from etl_arena.reporting.notificacion import a_markdown

    assert "2026-09" in MENSAJE["accion"] and "mes 2026-09 reemplazado" in a_markdown(MENSAJE)
    motivo = {"motivo": "parity_failed: el estado vigente no se tocó"}
    sin_publicar = construir_mensaje(
        "x", ("a", "b"), "parity_failed", "sin_exclusiones", {}, None, None, publicacion=motivo
    )
    assert "No publicado" in a_markdown(sin_publicar)


def test_sin_webhook_escribe_markdown(tmp_path, capsys):
    ruta = enviar(MENSAJE, tmp_path, webhook="")
    texto = (tmp_path / "notificacion.md").read_text(encoding="utf-8")
    assert ruta.endswith("notificacion.md") and "abc" in texto and "Sin Exclusiones" in texto


def test_webhook(tmp_path):
    recibido = {}

    class Manejador(http.server.BaseHTTPRequestHandler):
        def do_POST(self):
            recibido.update(json.loads(self.rfile.read(int(self.headers["Content-Length"]))))
            self.send_response(200)
            self.end_headers()

        def log_message(self, *a):
            pass

    servidor = http.server.HTTPServer(("127.0.0.1", 0), Manejador)
    threading.Thread(target=servidor.handle_request, daemon=True).start()
    assert enviar(MENSAJE, tmp_path, webhook=f"http://127.0.0.1:{servidor.server_port}/") == "webhook"
    assert recibido["id_corrida"] == "abc" and "text" in recibido
    assert not (tmp_path / "notificacion.md").exists()


def test_webhook_caido_cae_a_markdown(tmp_path, capsys):
    assert enviar(MENSAJE, tmp_path, webhook="http://127.0.0.1:9/").endswith("notificacion.md")


def test_mensaje_con_numero_de_corrida():
    m = construir_mensaje(
        "abc", ("2026-09-01", "2026-09-21"), "success", "sin_exclusiones", {}, None, None, num_corrida=12
    )
    from etl_arena.reporting.notificacion import a_markdown

    assert "Corrida N° 12" in m["titulo"] and m["num_corrida"] == 12
    assert "N° 12 · IdCorrida `abc`" in a_markdown(m)
    sin_bd = construir_mensaje("abc", ("a", "b"), "success", "sin_exclusiones", {}, None, None)
    assert "N°" not in sin_bd["titulo"] and "sin base de datos" in a_markdown(sin_bd)

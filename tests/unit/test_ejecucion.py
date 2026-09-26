"""Tests de la corrida completa sin SQL (tareas 3.1 y 3.2): orquestación, calidad y CLI."""

from datetime import date

import pytest
from fabricas import config_prueba, crear_libro, fila_raw, serial_min

from etl_arena.ejecucion import ejecutar_corrida, periodo_mes_en_curso
from etl_arena.model import KpiMensual
from etl_arena.persistence import ResumenPublicacion

S = [serial_min(15 * k) for k in range(1, 6)]


@pytest.fixture
def libro(tmp_path):
    filas = [
        fila_raw(S[0], [4, 4]),
        fila_raw(S[1], [3, None], ["F55 X", None]),
        fila_raw(S[2], [4, 4]),
        fila_raw(S[3], [4, 4]),
        fila_raw(S[4], [4, 4]),
    ]
    return crear_libro(tmp_path / "l.xlsx", filas, actividad={r: [None, S[r - 2], 1, 1] for r in range(2, 7)})


class RepoFalso:
    """Registra las llamadas del ciclo de vida sin base de datos."""

    def __init__(self, fallar_en_guardar=False):
        self.llamadas, self.fallar = [], fallar_en_guardar

    def exclusiones_cargadas(self, *a):
        return False

    def iniciar(self, cfg, meta):
        self.llamadas.append(("iniciar", meta.estado_exclusiones))
        return 7  # NumCorrida que asigna la base

    def cargar_catalogo(self):
        return {"F55": 55}

    def kpi_mensuales_vigentes(self, id_proyecto, anio):
        return [KpiMensual(anio, 7, 31, 2976, 350972.3, "excel_manual"), KpiMensual(anio, 9, 1, 1, 999.0, "corrida")]

    def publicar_mes(self, id_corrida, num_corrida, paquete, opciones, archivo, sha):
        if self.fallar:
            raise RuntimeError("SQL caído")
        self.llamadas.append(("publicar", f"{paquete.anio}-{paquete.mes:02d}", num_corrida))
        self.paquete = paquete
        return ResumenPublicacion(paquete.anio, paquete.mes)

    def guardar_referencia_excel(self, ref):
        self.llamadas.append(("referencia", ref.id_referencia))

    def guardar_reconciliacion(self, id_corrida, id_ref, filas):
        self.llamadas.append(("reconciliacion", len(filas)))

    def finalizar(self, id_corrida, estado, resumen_calidad=None, mensaje_error=None, **kw):
        self.llamadas.append(("finalizar", estado, mensaje_error is not None))
        self.finalizar_kw = kw


def test_sin_repositorio(libro):
    res = ejecutar_corrida(libro, config_prueba(), hash_archivo="0" * 64)
    assert res.estado == "success" and res.resultado.disponibilidad.bloques_racks_indisponibles == 12.0
    cal = res.resumen_calidad
    assert cal["reconciliacion"] == {"estado": "sin_referencia"}
    assert cal["modulos_nulos"]["celdas_en_periodo"] == 1 and cal["eventos"]["total"] == 1
    assert cal["exclusion_matrix"]["presente"] is False and cal["c23"]["distinto"] is False
    assert res.kpi_mensual.mes == 9  # período desde el día 1: KPI del mes aunque no se publique
    assert res.publicada is False and "publicacion" not in cal


def test_ciclo_de_vida_con_repositorio(libro):
    repo = RepoFalso()
    cfg = config_prueba(tipo_corrida="semanal")
    res = ejecutar_corrida(libro, cfg, repo, hash_archivo="0" * 64)
    assert [c[0] for c in repo.llamadas] == ["iniciar", "reconciliacion", "publicar", "finalizar"]
    assert res.num_corrida == 7 and repo.llamadas[2] == ("publicar", "2026-09", 7)
    assert res.publicada and repo.finalizar_kw["publicada"] and repo.finalizar_kw["meses_publicados"] == ["2026-09"]
    assert res.resumen_calidad["publicacion"]["mes"] == "2026-09"
    assert repo.llamadas[0] == ("iniciar", "sin_exclusiones") and repo.llamadas[-1] == ("finalizar", "success", False)
    # KPI mensual oficial (desde el día 1) y anual: julio de SQL + septiembre de esta corrida (reemplaza el vigente)
    assert res.kpi_mensual.mes == 9 and res.kpi_mensual.bloques_muestreo == 1 + 4
    assert [(f.mes, f.bloques_racks_indisponibles) for f in res.anual] == [(7, 350972.3), (8, None), (9, 12.0)]
    assert res.resumen_calidad["eventos"]["sin_catalogo"] == 0


def test_falla_queda_registrada(libro):
    repo = RepoFalso(fallar_en_guardar=True)
    with pytest.raises(RuntimeError):
        ejecutar_corrida(libro, config_prueba(), repo, hash_archivo="0" * 64)
    assert repo.llamadas[-1] == ("finalizar", "failed", True)


def _no_cuadra(monkeypatch):
    import etl_arena.ejecucion as ej
    from etl_arena.reconciliation.modelo import ReporteReconciliacion

    class NoCuadra(ReporteReconciliacion):
        estado_corrida = property(lambda self: "parity_failed")

    real = ej.reconciliar

    def falso(*a, **kw):
        rep = real(*a, **kw)
        rep.__class__ = NoCuadra
        return rep

    monkeypatch.setattr(ej, "reconciliar", falso)


def test_parity_failed_no_publica(libro, monkeypatch):  # D-22
    _no_cuadra(monkeypatch)
    repo = RepoFalso()
    res = ejecutar_corrida(libro, config_prueba(), repo, hash_archivo="0" * 64)
    assert res.estado == "parity_failed" and not res.publicada and "publicar" not in [c[0] for c in repo.llamadas]
    assert "no se tocó" in res.motivo_no_publicada and repo.finalizar_kw["publicada"] is False


def test_publicar_aunque_no_cuadre(libro, monkeypatch):
    _no_cuadra(monkeypatch)
    repo = RepoFalso()
    res = ejecutar_corrida(libro, config_prueba(), repo, hash_archivo="0" * 64, publicar_aunque_no_cuadre=True)
    assert res.estado == "parity_failed" and res.publicada


def test_golden_nunca_publica(libro):
    repo = RepoFalso()
    res = ejecutar_corrida(libro, config_prueba(tipo_corrida="golden"), repo, hash_archivo="0" * 64)
    assert res.estado == "success" and not res.publicada and "publicar" not in [c[0] for c in repo.llamadas]


def test_periodo_no_mensual_se_rechaza_antes_de_registrar(libro):  # D-20
    from etl_arena.persistence import ErrorPeriodoMensual

    repo = RepoFalso()
    with pytest.raises(ErrorPeriodoMensual):
        ejecutar_corrida(libro, config_prueba(inicio_periodo=date(2026, 8, 31)), repo, hash_archivo="0" * 64)
    assert repo.llamadas == []


def test_periodo_mes_en_curso():
    assert periodo_mes_en_curso(date(2026, 9, 21)) == (date(2026, 9, 1), date(2026, 9, 21))


def test_cli_en_seco(tmp_path):
    # el CLI usa la configuración del activo (61 PCS): libro sintético con 61 PCS y una falla en el PCS 1
    mods = [4.0] * 61
    filas = [fila_raw(S[0], mods), fila_raw(S[1], [3.0, *mods[1:]]), fila_raw(S[2], mods)]
    libro = crear_libro(
        tmp_path / "l61.xlsx", filas, total_pcs=61, actividad={r: [None, S[r - 2], 1, 1] for r in range(2, 5)}
    )
    import importlib.util
    from pathlib import Path

    spec = importlib.util.spec_from_file_location(
        "ejecutar_etl", Path(__file__).parents[2] / "scripts" / "ejecutar_etl.py"
    )
    cli = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(cli)
    salida = tmp_path / "r.json"
    codigo = cli.main(
        [
            "--libro",
            str(libro),
            "--desde",
            "2026-09-01",
            "--hasta",
            "2026-09-01",
            "--sin-bd",
            "--salida-json",
            str(salida),
            "--log-texto",
        ]
    )
    assert codigo == 0
    import json

    datos = json.loads(salida.read_text(encoding="utf-8"))
    assert (
        datos["estado"] == "success" and datos["C14"] == 12.0 and datos["reconciliacion"]["estado"] == "sin_referencia"
    )


def test_log_json_con_id_corrida(capsys):
    import json
    import logging

    from etl_arena.registro import configurar_logging

    configurar_logging(json_=True)
    log = logging.getLogger("etl_arena.corrida")
    log.info("hola %s", "mundo", extra={"id_corrida": "abc"})
    try:
        raise ValueError("x")
    except ValueError:
        log.exception("fallo", extra={"id_corrida": "abc"})
    lineas = [json.loads(linea) for linea in capsys.readouterr().err.strip().splitlines()]
    assert lineas[0]["mensaje"] == "hola mundo" and lineas[0]["id_corrida"] == "abc" and lineas[0]["nivel"] == "INFO"
    assert "ValueError" in lineas[1]["excepcion"]
    configurar_logging(json_=False)


def _cli():
    import importlib.util
    from pathlib import Path

    ruta = Path(__file__).parents[2] / "scripts" / "ejecutar_etl.py"
    spec = importlib.util.spec_from_file_location("ejecutar_etl", ruta)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def test_cli_periodo_explicito_manda_sobre_el_libro(monkeypatch):  # Checkpoint C-1
    from etl_arena.config import ParametrosExcel

    cli = _cli()
    del_libro = {
        "inicio_periodo": date(2026, 9, 1),
        "fin_periodo": date(2026, 9, 21),
        "inicio_periodo_eventos": date(2026, 9, 1),
        "fin_periodo_eventos": date(2026, 9, 21),
        "fin_diario": date(2026, 9, 21),
        "aplicar_evento_excusable": False,
    }
    monkeypatch.setattr(cli, "leer_celdas_parametros", lambda ruta: {})
    monkeypatch.setattr(cli, "parametros_desde_celdas", lambda celdas: ParametrosExcel(dict(del_libro)))
    args = cli.construir_parser().parse_args(
        [
            "--libro",
            "x.xlsm",
            "--parametros-desde-libro",
            "--desde",
            "2026-09-01",
            "--hasta",
            "2026-09-28",
        ]
    )
    v = cli.valores_config(args)
    assert (v["fin_periodo"], v["fin_periodo_eventos"], v["fin_diario"]) == (date(2026, 9, 28),) * 3
    assert v["aplicar_evento_excusable"] is False  # lo demás sigue viniendo del libro
    args = cli.construir_parser().parse_args(
        ["--libro", "x.xlsm", "--parametros-desde-libro", "--hasta", "2026-09-28", "--eventos-fin", "2026-09-27"]
    )
    assert cli.valores_config(args)["fin_periodo_eventos"] == date(2026, 9, 27)  # explícito gana


def test_referencia_usa_el_mismo_texto_que_los_motores():  # Checkpoint C-3
    from etl_arena.excel_semantics import texto_excel
    from etl_arena.workbook.referencia import _txt

    for v in (169.0, 3.5, 0.1 + 0.2, "F55 X", None):
        assert _txt(v) == (None if v is None else texto_excel(v))

"""Persistencia sin base de datos: scripts, selección de staging y armado de filas (R11, D-09)."""

from datetime import date, datetime

import numpy as np
from fabricas import config_prueba, crear_libro, fila_raw, serial_min

from etl_arena.persistence import MetadatosCorrida, construir_paquete, filas_staging, lotes
from etl_arena.persistence.esquema import DIR_SQL, ORDEN_SCRIPTS
from etl_arena.persistence.paquete import _dt, _f, fila_etl_run
from etl_arena.pipeline import calcular_libro


def test_lotes_por_go():
    sql = ":setvar X 1\nSELECT 1\nGO\n\n  go  \nSELECT 2;\nGO;\n-- fin\n"
    assert lotes(sql) == ["SELECT 1", "SELECT 2;", "-- fin"]


def test_scripts_existen_y_sin_tipos_prohibidos():  # convención DATETIME (2026-09-24)
    for nombre in (*ORDEN_SCRIPTS, "00_database.sql", "07_audit_queries.sql"):
        lineas = (DIR_SQL / nombre).read_text(encoding="utf-8").splitlines()
        texto = " ".join(linea.split("--")[0] for linea in lineas).upper()  # sin comentarios
        assert "DATETIME2" not in texto and " DATE " not in texto and " DATE," not in texto, nombre


def test_conversiones():
    assert _f(np.nan) is None and _f(np.float64(1.5)) == 1.5 and _f(None) is None
    assert _dt(np.datetime64("NaT", "s")) is None
    assert _dt(np.datetime64("2026-09-01T00:15:00")) == datetime(2026, 9, 1, 0, 15)
    assert _dt(date(2026, 9, 1)) == datetime(2026, 9, 1)
    assert _dt(datetime(2026, 9, 1, 0, 15, 0, 999)) == datetime(2026, 9, 1, 0, 15)


def _resultado(tmp_path, **cfg):
    s = [serial_min(15 * k) for k in range(-1, 4)]  # 31-ago 23:45 … 01-sep 00:45
    filas = [
        fila_raw(s[0], [4, None]),
        fila_raw(s[1], [4, 4]),
        fila_raw(s[2], [3, 4], ["F55 X", "NO FAULTS"]),
        fila_raw(s[3], [4, 4]),
        fila_raw(s[4], [4, 4]),
    ]
    ruta = crear_libro(tmp_path / "l.xlsx", filas, actividad={r: [None, s[r - 2], 1, 1] for r in range(2, 7)})
    return calcular_libro(ruta, config_prueba(**cfg), codigos_resumen=[("F55", "EXT")])


def test_filas_staging_incluye_la_anterior_al_periodo(tmp_path):
    r = _resultado(tmp_path)
    assert filas_staging(r).tolist() == [0, 1, 2, 3, 4]  # fila 0 (31-ago) es la anterior al período


def test_paquete(tmp_path):
    r = _resultado(tmp_path)
    meta = MetadatosCorrida("f" * 64, filas_nuevas_export=frozenset({5, 6}))
    paq = construir_paquete(r, meta, {"F55": 55})
    raw = paq.tabla("raw_pcs_sample")
    assert len(raw.filas) == 5 * 2
    primera = raw.filas[1]  # fila 2, PCS 2: módulos vacíos
    assert primera[9] is None and primera[10] == 4.0 and primera[11] is True and primera[12] is False
    assert [f[12] for f in raw.filas[-4:]] == [True, True, True, True]  # filas 5 y 6 del export
    asr = paq.tabla("availability_sample_result")
    assert len(asr.filas) == 4 * 2 and sum(f[11] for f in asr.filas) == r.disponibilidad.bloques_racks_indisponibles
    det = paq.tabla("detencion").filas
    assert [(d[7], d[8]) for d in det] == [(55, "F55")] and paq.anomalias_persistencia == []
    assert paq.tabla("availability_run_result").filas[0][8] == 96.0  # 1 día * 24*60/15
    dqi = {f[1] for f in paq.tabla("data_quality_issue").filas}
    assert "exclusion_matrix_ausente" in dqi


def test_codigo_sin_catalogo_se_reporta(tmp_path):
    paq = construir_paquete(_resultado(tmp_path), MetadatosCorrida("f" * 64), {})
    assert paq.tabla("detencion").filas[0][7] is None
    assert [a.tipo for a in paq.anomalias_persistencia] == ["codigo_sin_catalogo"]


def test_fila_etl_run():
    cfg = config_prueba(es_oficial=True)
    columnas, valores = fila_etl_run(cfg, MetadatosCorrida("e" * 64, estado_exclusiones="con_exclusiones"))
    fila = dict(zip(columnas, valores, strict=True))
    assert fila["Estado"] == "running" and fila["EstadoExclusiones"] == "con_exclusiones"
    assert fila["InicioPeriodo"] == datetime(2026, 9, 1) and fila["VersionAlgoritmo"] == cfg.version_algoritmo

"""Persistencia sin base de datos: scripts y armado del paquete de un mes (R11 rev. 3, ADR-12)."""

from datetime import date, datetime

import numpy as np
import pytest
from fabricas import config_prueba, crear_libro, fila_raw, serial_min

from etl_arena.persistence import (
    ErrorPeriodoMensual,
    MetadatosCorrida,
    construir_paquete_mes,
    lotes,
    validar_periodo_mensual,
)
from etl_arena.persistence.esquema import DIR_SQL, ORDEN_SCRIPTS
from etl_arena.persistence.paquete import _dt, _f, _ocurrencias, fila_etl_run
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


def _dic(tabla):
    return [dict(zip(tabla.columnas, f, strict=True)) for f in tabla.filas]


def test_paquete_mes(tmp_path):
    r = _resultado(tmp_path)
    meta = MetadatosCorrida("f" * 64, filas_nuevas_export=frozenset({5, 6}))
    paq = construir_paquete_mes(r, meta, {"F55": 55})
    assert (paq.anio, paq.mes, paq.estado_exclusiones) == (2026, 9, "sin_exclusiones") and not paq.mes_completo
    mp = _dic(paq.muestras_pcs)
    assert len(mp) == 4 * 2  # solo las filas del período (la de 31-ago no es de septiembre)
    assert {(m["Anio"], m["Mes"]) for m in mp} == {(2026, 9)} and all(m["Ocurrencia"] == 1 for m in mp)
    assert sum(m["ImpactoRackPonderado"] for m in mp) == r.disponibilidad.bloques_racks_indisponibles
    assert [m["EsFilaNuevaDelExport"] for m in mp[-4:]] == [True] * 4  # filas 5 y 6 del export
    assert paq.crudos_pcs[(mp[2]["SerialFecha"], 1, 1)][2] == 3.0  # MODULES crudo para detectar correcciones
    assert len(paq.muestras_planta.filas) == 4 and len(paq.crudos_planta) == 4
    det = _dic(paq.detenciones)
    assert [(d["IdTipoDetencion"], d["CodigoFalla"], d["Ocurrencia"]) for d in det] == [(55, "F55", 1)]
    assert paq.anomalias_persistencia == []
    assert all(d["DuracionSegundos"] == round(d["DuracionHoras"] * 3600) for d in det)
    assert all((d["Anio"], d["Mes"]) == (2026, 9) for d in det)
    cols = list(paq.detenciones.columnas)
    assert cols.index("DuracionHoras") == cols.index("DuracionSegundos") + 1
    assert "NumCorrida" not in cols  # se agrega al publicar
    m = paq.mensual
    assert m["BloquesMuestreoCalendario"] == 96.0 and m["BloquesRacksIndisponibles"] == 12.0
    assert m["Origen"] == "corrida" and m["UltimoDato"] == paq.ultimo_dato and m["AcumulaEnAnual"] is True
    assert "exclusion_matrix_ausente" in {c["Tipo"] for c in _dic(paq.calidad)}
    assert [f[0] for f in paq.diario.filas] == [1]


def test_paquete_exige_periodo_mensual(tmp_path):  # D-20
    r = _resultado(tmp_path, inicio_periodo=date(2026, 8, 31))
    with pytest.raises(ErrorPeriodoMensual):
        construir_paquete_mes(r, MetadatosCorrida("f" * 64), {})
    with pytest.raises(ErrorPeriodoMensual):
        validar_periodo_mensual(config_prueba(fin_periodo=date(2026, 10, 1)))
    validar_periodo_mensual(config_prueba(fin_periodo=date(2026, 9, 30)))


def test_ocurrencias_por_timestamp_repetido():  # F-32: cambio de hora
    assert _ocurrencias([1.0, 2.0, 2.0, 3.0, 2.0]) == [1, 1, 2, 1, 3]


def test_codigo_sin_catalogo_se_reporta(tmp_path):
    paq = construir_paquete_mes(_resultado(tmp_path), MetadatosCorrida("f" * 64), {})
    assert _dic(paq.detenciones)[0]["IdTipoDetencion"] is None
    assert [a.tipo for a in paq.anomalias_persistencia] == ["codigo_sin_catalogo"]
    assert "codigo_sin_catalogo" in {c["Tipo"] for c in _dic(paq.calidad)}


def test_fila_etl_run():
    cfg = config_prueba()
    columnas, valores = fila_etl_run(cfg, MetadatosCorrida("e" * 64, estado_exclusiones="con_exclusiones"))
    fila = dict(zip(columnas, valores, strict=True))
    assert fila["Estado"] == "running" and fila["EstadoExclusiones"] == "con_exclusiones" and "EsOficial" not in fila
    assert fila["InicioPeriodo"] == datetime(2026, 9, 1) and fila["VersionAlgoritmo"] == cfg.version_algoritmo

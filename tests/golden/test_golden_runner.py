"""Golden runner (tarea 1.11, R14, R13.7): motores Python vs valores cacheados del libro.

Para cada golden cuyo libro fuente está disponible, ejecuta los motores con los
**parámetros efectivos** del golden y compara los niveles 2–5 con la tabla de tolerancias.
Hoy solo el libro de septiembre está en ``data/``.
"""

from __future__ import annotations

import gzip
import json
import math
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pytest

from etl_arena.aggregation import calcular_diaria
from etl_arena.availability import calcular
from etl_arena.config import construir_config
from etl_arena.fault_events import detectar_eventos
from etl_arena.ingestion.catalogos import leer_codigos_resumen
from etl_arena.reconciliation import TOLERANCIAS as T

pytestmark = pytest.mark.golden
DATA = Path(__file__).parent / "data"
GOLDEN_SEP = "golden_2026_09_september.json"


def _fecha(s: str) -> date:
    return date.fromisoformat(s)


def cfg_desde_golden(g: dict):
    ef = g["effective_parameters"]
    inicio = _fecha(ef["C5_period_start"])
    return construir_config(
        inicio_periodo=inicio,
        fin_periodo=_fecha(ef["C7_period_end"]),
        solo_tiempo_operacional=ef["C21_only_operational_time"] != "No",
        aplicar_evento_excusable=ef["C31_apply_excused_event"] == "Yes",
        inicio_periodo_eventos=_fecha(ef["L2_events_start"]),
        fin_periodo_eventos=_fecha(ef["L4_events_end"]),
        aplicar_evento_excusable_eventos=ef["L14_events_apply_excused_event"] == "Yes",
        fin_diario=inicio + timedelta(days=len(g["daily"]) - 1),  # Daily!C9.. (F-33)
        tipo_corrida="golden",
    )


@pytest.fixture(scope="module")
def corrida_sep(matriz_real_sep, ruta_libro_real):
    g = json.loads((DATA / GOLDEN_SEP).read_text(encoding="utf-8"))
    if g["_meta"]["source_file"] != ruta_libro_real.name:
        pytest.skip("el libro disponible no es la fuente del golden")
    with gzip.open(DATA / g["_meta"]["calc_table_file"]) as fh:
        tabla = json.loads(fh.read())
    _, m, act, _, exc = matriz_real_sep
    cfg = cfg_desde_golden(g)
    disp = calcular(m, act, cfg, exc)
    eventos = detectar_eventos(m, exc, cfg, disp.minutos_muestreo_derivado, leer_codigos_resumen(ruta_libro_real))
    return g, tabla, cfg, disp, eventos, calcular_diaria(disp, cfg)


def _cerca(a, b, tol):
    if a is None or b is None:
        return a is None and b is None
    return abs(a - b) <= tol


# ------------------------------------------------------------------ nivel 1–2
def test_nivel2_tabla_de_resultados(corrida_sep):
    g, tabla, _, disp, _, _ = corrida_sep
    filas = tabla["rows"]
    assert len(filas) == disp.bloques_muestreo == g["period"]["kpi"]["c12_sample_blocks"]
    malas = []
    for k, (serial, valores, bo) in enumerate(filas):
        assert serial == disp.serial[k]
        assert bo == disp.bo_evento_excusado_pa[k]
        for j, v in enumerate(valores):
            py = disp.ponderadas[k, j]
            if v is None:
                if not np.isnan(py):
                    malas.append((k, j, py, v))
            elif np.isnan(py) or abs(py - v) > T.ponderada_celda:
                malas.append((k, j, py, v))
    assert malas == []


# ------------------------------------------------------------------ nivel 3–4
def test_nivel3_acumulados(corrida_sep):
    g, _, _, disp, _, _ = corrida_sep
    kpi = g["period"]["kpi"]
    assert disp.bloques_muestreo == kpi["c12_sample_blocks"]
    assert _cerca(disp.bloques_racks_indisponibles, kpi["c14_unavailable_rack_blocks"], T.c14)
    assert disp.minutos_muestreo_derivado == g["effective_parameters"]["C23_sampling_minutes"]


def test_nivel4_kpi_y_daily(corrida_sep):
    g, _, _, disp, _, diario = corrida_sep
    kpi = g["period"]["kpi"]
    assert _cerca(disp.disponibilidad_periodo, kpi["c16_availability_period"], T.kpi_disponibilidad)
    assert _cerca(disp.disponibilidad_anual_acumulada, kpi["c19_accumulated_annual"], T.kpi_disponibilidad)
    assert len(diario) == len(g["daily"])
    for d, gd in zip(diario, g["daily"], strict=True):
        assert d.numero_dia == gd["day_number"] and d.dia.isoformat() == gd["date"]
        assert _cerca(d.diario, gd["daily_unavailable_rack_blocks"], T.daily_bloques)
        assert _cerca(d.acumulado, gd["accumulated_unavailable_rack_blocks"], T.daily_bloques)
        assert _cerca(d.disponibilidad, gd["availability"], T.kpi_disponibilidad)
        assert _cerca(d.variacion, gd["variation"], T.kpi_disponibilidad)


# ------------------------------------------------------------------ nivel 5
def test_nivel5_eventos(corrida_sep):
    g, _, _, _, eventos, _ = corrida_sep
    ge = g["fault_events_summary"]
    assert len(eventos.cerrados) == ge["total_events"] == len(ge["events"])
    assert eventos.incompleto is None and not any(e.get("incomplete") for e in ge["events"])
    for e, x in zip(eventos.cerrados, ge["events"], strict=True):
        clave = f"orden {x['order_excel']}"
        assert e.orden_excel == x["order_excel"], clave
        assert (e.numero_pcs, e.codigo_falla, e.descripcion_falla) == (
            x["pcs_number"],
            x["fault_code"],
            x["fault_description"],
        ), clave
        assert _cerca(e.serial_inicio, x["start_serial"], T.evento_serial_dias), clave
        assert _cerca(e.serial_fin, x["end_serial"], T.evento_serial_dias), clave
        assert _cerca(e.duracion_horas, x["duration_hours"], T.evento_duracion_promedio), clave
        assert _cerca(e.promedio_baterias, x["average_batteries_involved"], T.evento_duracion_promedio), clave
        assert _cerca(e.horas_rack, x["unavailable_rack_hours"], T.evento_horas_rack), clave
    assert _cerca(eventos.horas_rack_totales, ge["total_unavailable_rack_hours"], T.l10)


def test_nivel5_resumen_por_codigo(corrida_sep):
    """N:Q se compara como mapa código → P: el orden de empates depende del estado previo de la hoja."""
    g, _, _, _, eventos, _ = corrida_sep
    esperado = {x["fault_code"]: x["unavailable_rack_hours"] for x in g["fault_code_summary"]}
    obtenido = {f.codigo: f.horas_rack for f in eventos.resumen}
    assert set(obtenido) == set(esperado)
    assert all(_cerca(obtenido[c], p, T.l10) for c, p in esperado.items())
    ordenados = [x["fault_code"] for x in g["fault_code_summary"] if x["unavailable_rack_hours"]]
    assert [f.codigo for f in eventos.resumen if f.horas_rack] == ordenados


# ------------------------------------------------------------------ objetivo: paridad bit a bit
def test_objetivo_paridad_exacta(corrida_sep):
    """Columna "objetivo" de la tabla de tolerancias: hoy la paridad es exacta; cualquier
    cambio que la rompa (aunque quede dentro del gate) debe ser deliberado."""
    g, _, _, disp, eventos, diario = corrida_sep
    kpi = g["period"]["kpi"]
    assert disp.bloques_racks_indisponibles == kpi["c14_unavailable_rack_blocks"]
    assert disp.disponibilidad_periodo == kpi["c16_availability_period"]
    assert disp.disponibilidad_anual_acumulada == kpi["c19_accumulated_annual"]
    assert eventos.horas_rack_totales == g["fault_events_summary"]["total_unavailable_rack_hours"]
    assert [e.horas_rack for e in eventos.cerrados] == [
        x["unavailable_rack_hours"] for x in g["fault_events_summary"]["events"]
    ]
    assert [d.disponibilidad for d in diario] == [x["availability"] for x in g["daily"]]


def test_tolerancias_sincronizadas_con_golden_index():
    meta = json.loads((DATA / "golden_index.json").read_text(encoding="utf-8"))["_meta"]
    assert meta["tolerance_kpi"] == T.c14 == T.l10
    assert meta["tolerance_internal"] == T.kpi_disponibilidad
    assert meta["tolerance_daily"] == T.daily_bloques
    assert meta["tolerance_rack_hours"] == T.evento_horas_rack
    assert not math.isnan(T.ponderada_celda)

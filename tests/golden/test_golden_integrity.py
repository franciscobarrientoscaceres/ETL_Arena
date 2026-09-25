"""
test_golden_integrity.py
========================
Invariantes de coherencia interna sobre los golden JSON en tests/golden/data/.

NO requiere el motor de disponibilidad ni Excel: valida que los fixtures sean
coherentes consigo mismos. Soporta schema 1.1 (jul/ago: events_first_20) y
schema 1.2 (sep: todos los eventos, parámetros efectivos, tabla de resultados).

Correr:  pytest tests/golden/test_golden_integrity.py -v
"""

from __future__ import annotations

import gzip
import json
import math
from datetime import datetime
from pathlib import Path

import pytest
from extract_golden import codigo_falla_excel, validate_golden

DATA_DIR = Path(__file__).parent / "data"
INDEX_PATH = DATA_DIR / "golden_index.json"

TOL_EXACT = 1e-9
TOL_SUM = 1e-6
TOL_KPI = 1e-12
TOL_RH_V1 = 1e-3
LEGACY_EVENT_CAP = 295


def _load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _index() -> dict:
    return _load(INDEX_PATH)


def _golden_ids() -> list[str]:
    return [m["id"] for m in _index()["months"]]


def _index_entry(id_mes: str) -> dict:
    return next(m for m in _index()["months"] if m["id"] == id_mes)


def _golden_for(id_mes: str) -> dict:
    return _load(DATA_DIR / _index_entry(id_mes)["file"])


def _events(g: dict) -> list[dict]:
    ev = g["fault_events_summary"]
    return ev.get("events") or ev["events_first_20"]


def _is_v12(g: dict) -> bool:
    return g.get("_meta", {}).get("schema_version") == "1.2"


def _c21_no(g: dict) -> bool:
    return g["period"]["parameters"]["only_operational_time"] == "No"


@pytest.fixture(params=_golden_ids(), ids=_golden_ids())
def golden(request) -> dict:
    return _golden_for(request.param)


@pytest.fixture(params=_golden_ids(), ids=_golden_ids())
def index_entry(request) -> dict:
    return _index_entry(request.param)


def test_index_kpi_matches_fixture(index_entry: dict):
    g = _load(DATA_DIR / index_entry["file"])
    k = g["period"]["kpi"]
    assert index_entry["kpi_summary"]["c12"] == k["c12_sample_blocks"]
    assert index_entry["kpi_summary"]["c16_availability_period"] == pytest.approx(k["c16_availability_period"], abs=0)
    assert index_entry["period_start"] == g["period"]["period_start"]
    assert index_entry["period_end"] == g["period"]["period_end"]


def test_sum_daily_equals_c14(golden: dict):
    if not _c21_no(golden):
        pytest.skip("C21 != 'No': Daily no aplica factor operacional (F-08)")
    c14 = golden["period"]["kpi"]["c14_unavailable_rack_blocks"]
    s = sum(d["daily_unavailable_rack_blocks"] or 0 for d in golden["daily"])
    assert abs(s - c14) <= TOL_SUM, f"sum(daily)={s} != C14={c14}"


def test_last_accumulated_equals_c14(golden: dict):
    if not _c21_no(golden):
        pytest.skip("C21 != 'No' (F-08)")
    c14 = golden["period"]["kpi"]["c14_unavailable_rack_blocks"]
    last = golden["daily"][-1]["accumulated_unavailable_rack_blocks"]
    assert last is not None
    assert abs(last - c14) <= TOL_SUM, f"last_accum={last} != C14={c14}"


def test_c16_recompute(golden: dict):
    kpi = golden["period"]["kpi"]
    racks = golden["period"]["parameters"]["total_racks"]
    expected = 1.0 - kpi["c14_unavailable_rack_blocks"] / (racks * kpi["c12_sample_blocks"])
    assert math.isclose(kpi["c16_availability_period"], expected, rel_tol=0, abs_tol=TOL_KPI)


def test_july_august_daily_series_differ():
    j = [d["daily_unavailable_rack_blocks"] for d in _golden_for("2026_07")["daily"]]
    a = [d["daily_unavailable_rack_blocks"] for d in _golden_for("2026_08")["daily"]]
    assert j != a, "daily de agosto no puede ser copia de julio"


def test_fault_code_orientation(golden: dict):
    events = [e for e in _events(golden) if not e.get("incomplete")]
    assert events, "sin eventos"
    for e in events:
        fc, fd = e.get("fault_code"), e.get("fault_description")
        if _is_v12(golden):
            assert fc == codigo_falla_excel(fd), f"orden {e['order_excel']}: F={fc!r} vs G={fd!r}"
        else:
            assert str(fc).startswith("F") and str(fd).startswith(str(fc)), f"{fc!r} / {fd!r}"


def test_event_duration_multiple_of_sampling(golden: dict):
    block_h = golden["period"]["parameters"]["sampling_minutes"] / 60.0
    for e in _events(golden):
        if e.get("incomplete"):
            continue
        blocks = e["duration_hours"] / block_h
        assert abs(blocks - round(blocks)) <= 1e-6, f"duración {e['duration_hours']} no múltiplo de {block_h}"


def test_event_rack_hours_formula(golden: dict):
    tol = TOL_EXACT if _is_v12(golden) else TOL_RH_V1
    for e in _events(golden):
        if e.get("incomplete"):
            continue
        expected = 12.0 * e["duration_hours"] * e["average_batteries_involved"]
        assert abs(expected - e["unavailable_rack_hours"]) <= tol


def test_event_wall_clock_matches_duration(golden: dict):
    block_h = golden["period"]["parameters"]["sampling_minutes"] / 60.0
    for e in _events(golden):
        if e.get("incomplete"):
            continue
        if _is_v12(golden):
            assert abs(24 * (e["end_serial"] - e["start_serial"]) - e["duration_hours"]) <= TOL_EXACT
        else:
            t0 = datetime.strptime(e["start_timestamp"], "%Y-%m-%d %H:%M:%S")
            t1 = datetime.strptime(e["end_timestamp"], "%Y-%m-%d %H:%M:%S")
            wall = (t1 - t0).total_seconds() / 3600.0
            assert abs(wall - e["duration_hours"]) <= block_h / 2 + 1e-6


def test_legacy_event_cap_flagged_in_index(index_entry: dict):
    """Si el fixture aún tiene 295 eventos, el índice debe marcarlo como pending."""
    g = _load(DATA_DIR / index_entry["file"])
    if g["fault_events_summary"]["total_events"] == LEGACY_EVENT_CAP:
        assert index_entry["status"] != "verified"
        assert any(d.get("id") == "events_cap_295" for d in index_entry.get("discrepancies", []))


def test_september_c12_discrepancy_resolved_in_index():
    disc = next(d for d in _index_entry("2026_09")["discrepancies"] if d["id"] == "c12_vs_annual_sampling_blocks")
    assert disc["status"] == "resolved"
    assert "1975" in disc["resolution"]


def test_no_unresolved_verificar_lunes_notes():
    for path in [INDEX_PATH, *DATA_DIR.glob("golden_*.json")]:
        text = path.read_text(encoding="utf-8").lower()
        assert "verificar lunes" not in text, path.name
        assert "pendiente verificar" not in text, path.name


def test_annual_snapshot_divergence_is_documented_not_fatal():
    """Annual_AVA histórico puede divergir de C14; solo exigimos que esté documentado."""
    idx = _index()
    policy = idx["_meta"].get("annual_ava_policy", "")
    assert "GT-5" in policy or "historico" in policy.lower()
    for m in idx["months"]:
        g = _load(DATA_DIR / m["file"])
        ann = g.get("annual_ava_snapshot") or {}
        key = str(g["period"]["month"])
        if key not in ann:
            continue
        au = ann[key]["unavailable_rack_blocks"]
        c14 = g["period"]["kpi"]["c14_unavailable_rack_blocks"]
        if au is not None and abs(au - c14) > TOL_RH_V1 and g["period"]["month"] in (7, 8):
            assert "annual_ava_diverges_c14" in {d.get("id") for d in m.get("discrepancies", [])}


# ------------------------------------------------------------- solo schema 1.2


def _v12_ids() -> list[str]:
    return [m["id"] for m in _index()["months"] if m.get("schema_version") == "1.2"]


@pytest.fixture(params=_v12_ids(), ids=_v12_ids())
def golden_v12(request) -> dict:
    return _golden_for(request.param)


def _calc_table(g: dict) -> dict:
    with gzip.open(DATA_DIR / g["_meta"]["calc_table_file"], "rt", encoding="utf-8") as fh:
        return json.load(fh)


def test_v12_validate_golden_has_no_errors(golden_v12: dict):
    errors, _ = validate_golden(golden_v12, _calc_table(golden_v12))
    assert errors == []


def test_v12_effective_parameters_present(golden_v12: dict):
    eff = golden_v12["effective_parameters"]
    for k in (
        "C5_period_start",
        "C7_period_end",
        "C21_only_operational_time",
        "C31_apply_excused_event",
        "C23_sampling_minutes",
        "L2_events_start",
        "L4_events_end",
        "L14_events_apply_excused_event",
        "D5_daily_end",
    ):
        assert k in eff, k


def test_v12_all_events_present_in_write_order(golden_v12: dict):
    ev = golden_v12["fault_events_summary"]
    assert len(ev["events"]) == ev["total_events"]
    assert [e["order_excel"] for e in ev["events"]] == list(range(1, ev["total_events"] + 1))


def test_v12_l10_is_sequential_sum(golden_v12: dict):
    ev = golden_v12["fault_events_summary"]
    s = 0.0
    for e in ev["events"]:
        if not e.get("incomplete"):
            s = s + 12.0 * e["duration_hours"] * e["average_batteries_involved"]
    assert abs(s - ev["total_unavailable_rack_hours"]) <= TOL_SUM


def test_v12_calc_table_rows_equal_c12(golden_v12: dict):
    table = _calc_table(golden_v12)
    assert len(table["rows"]) == golden_v12["period"]["kpi"]["c12_sample_blocks"]
    assert all(len(vals) == table["total_pcs"] for _, vals, _ in table["rows"])


def test_v12_calc_table_sums_to_c14(golden_v12: dict):
    if not _c21_no(golden_v12):
        pytest.skip("C21 != 'No'")
    s = 0.0
    for _, vals, _ in _calc_table(golden_v12)["rows"]:
        for v in vals:
            s = s + 12 * (v or 0.0)
    assert abs(s - golden_v12["period"]["kpi"]["c14_unavailable_rack_blocks"]) <= TOL_SUM

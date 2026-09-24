"""
test_golden_integrity.py
========================
Invariantes de coherencia interna sobre los golden JSON en tests/golden/data/.

NO requiere el motor de disponibilidad ni Excel: valida que los fixtures
mismos son autosiguientes. Cualquier hand-edit que rompa sum(daily)=C14,
la orientacion de fault_code/description o la formula de C16 falla aqui.

Correr:  pytest tests/golden/test_golden_integrity.py -v
"""
from __future__ import annotations

import json
import math
import re
from datetime import datetime
from pathlib import Path

import pytest

DATA_DIR = Path(__file__).parent / "data"
INDEX_PATH = DATA_DIR / "golden_index.json"

TOL_DAILY = 1e-6
TOL_KPI = 1e-12
TOL_RH = 1e-3
CODE_RE = re.compile(r"^F\d+")
LEGACY_EVENT_CAP = 295


def _load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _golden_ids() -> list[str]:
    idx = _load(INDEX_PATH)
    return [m["id"] for m in idx["months"]]


def _golden_for(id_mes: str) -> dict:
    idx = _load(INDEX_PATH)
    entry = next(m for m in idx["months"] if m["id"] == id_mes)
    return _load(DATA_DIR / entry["file"])


def _index_entry(id_mes: str) -> dict:
    idx = _load(INDEX_PATH)
    return next(m for m in idx["months"] if m["id"] == id_mes)


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
    assert index_entry["kpi_summary"]["c16_availability_period"] == pytest.approx(
        k["c16_availability_period"], abs=0
    )
    assert index_entry["period_start"] == g["period"]["period_start"]
    assert index_entry["period_end"] == g["period"]["period_end"]


def test_sum_daily_equals_c14(golden: dict):
    c14 = golden["period"]["kpi"]["c14_unavailable_rack_blocks"]
    s = sum(d["daily_unavailable_rack_blocks"] or 0 for d in golden["daily"])
    assert abs(s - c14) <= TOL_DAILY, f"sum(daily)={s} != C14={c14}"


def test_last_accumulated_equals_c14(golden: dict):
    c14 = golden["period"]["kpi"]["c14_unavailable_rack_blocks"]
    last = golden["daily"][-1]["accumulated_unavailable_rack_blocks"]
    assert last is not None
    assert abs(last - c14) <= TOL_DAILY, f"last_accum={last} != C14={c14}"


def test_c16_recompute(golden: dict):
    kpi = golden["period"]["kpi"]
    racks = golden["period"]["parameters"]["total_racks"]
    expected = 1.0 - kpi["c14_unavailable_rack_blocks"] / (racks * kpi["c12_sample_blocks"])
    assert math.isclose(
        kpi["c16_availability_period"], expected, rel_tol=0, abs_tol=TOL_KPI
    ), f"C16={kpi['c16_availability_period']} != {expected}"


def test_july_august_daily_series_differ():
    jul = _golden_for("2026_07")
    ago = _golden_for("2026_08")
    j = [d["daily_unavailable_rack_blocks"] for d in jul["daily"]]
    a = [d["daily_unavailable_rack_blocks"] for d in ago["daily"]]
    assert j != a, "daily de agosto no puede ser copia de julio"


def test_fault_code_is_short_and_description_long(golden: dict):
    events = golden["fault_events_summary"]["events_first_20"]
    assert events, "events_first_20 vacio"
    for i, e in enumerate(events):
        fc = str(e.get("fault_code") or "")
        fd = str(e.get("fault_description") or "")
        assert CODE_RE.match(fc), f"event[{i}]: fault_code no es Fnn: {fc!r}"
        assert len(fd) > len(fc), f"event[{i}]: description no larga: {fd!r}"
        assert fd.startswith(fc), f"event[{i}]: description no prefija por code: {fd!r} / {fc!r}"


def test_event_duration_multiple_of_sampling(golden: dict):
    block_h = golden["period"]["parameters"]["sampling_minutes"] / 60.0
    for i, e in enumerate(golden["fault_events_summary"]["events_first_20"]):
        dur = e["duration_hours"]
        blocks = dur / block_h
        assert abs(blocks - round(blocks)) <= 1e-6, (
            f"event[{i}]: duration {dur} no multiple de {block_h}"
        )


def test_event_rack_hours_formula(golden: dict):
    for i, e in enumerate(golden["fault_events_summary"]["events_first_20"]):
        expected = 12.0 * e["duration_hours"] * e["average_batteries_involved"]
        actual = e["unavailable_rack_hours"]
        assert abs(expected - actual) <= TOL_RH, (
            f"event[{i}]: rh={actual} != 12*dur*avg={expected}"
        )


def test_event_wall_clock_matches_duration(golden: dict):
    block_h = golden["period"]["parameters"]["sampling_minutes"] / 60.0
    for i, e in enumerate(golden["fault_events_summary"]["events_first_20"]):
        t0 = datetime.strptime(e["start_timestamp"], "%Y-%m-%d %H:%M:%S")
        t1 = datetime.strptime(e["end_timestamp"], "%Y-%m-%d %H:%M:%S")
        wall = (t1 - t0).total_seconds() / 3600.0
        assert abs(wall - e["duration_hours"]) <= block_h / 2 + 1e-6, (
            f"event[{i}]: wall={wall} vs duration={e['duration_hours']}"
        )


def test_legacy_event_cap_flagged_in_index(index_entry: dict):
    """Si el fixture aun tiene 295 eventos, el indice debe marcarlo como pending."""
    g = _load(DATA_DIR / index_entry["file"])
    n = g["fault_events_summary"]["total_events"]
    if n == LEGACY_EVENT_CAP:
        assert index_entry["status"] != "verified", (
            "total_events=295 (cap extractor antiguo) no puede estar verified sin re-extract"
        )
        assert any(
            d.get("id") == "events_cap_295" for d in index_entry.get("discrepancies", [])
        ), "falta discrepancy events_cap_295 en el indice"


def test_september_c12_discrepancy_resolved_in_index():
    entry = _index_entry("2026_09")
    disc = next(
        d for d in entry["discrepancies"] if d["id"] == "c12_vs_annual_sampling_blocks"
    )
    assert disc["status"] == "resolved"
    assert "1975" in disc["resolution"]


def test_no_unresolved_verificar_lunes_notes():
    for name in ("golden_index.json",) + tuple(
        p.name for p in DATA_DIR.glob("golden_*.json")
    ):
        text = (DATA_DIR / name).read_text(encoding="utf-8")
        assert "verificar lunes" not in text.lower(), f"{name} aun contiene 'verificar lunes'"
        assert "pendiente verificar" not in text.lower(), f"{name} aun contiene 'PENDIENTE verificar'"


def test_annual_snapshot_divergence_is_documented_not_fatal():
    """Annual_AVA historico puede divergir de C14; solo exigimos que este documentado."""
    idx = _load(INDEX_PATH)
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
        if au is not None and abs(au - c14) > TOL_RH:
            # mes linkado debe coincidir en sep; jul/ago divergen y deben tener discrepancy
            disc_ids = {d.get("id") for d in m.get("discrepancies", [])}
            if g["period"]["month"] in (7, 8):
                assert "annual_ava_diverges_c14" in disc_ids, (
                    f"{m['id']}: divergence annual!=C14 sin documentar en discrepancies"
                )

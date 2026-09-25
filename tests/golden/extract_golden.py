#!/usr/bin/env python3
"""
extract_golden.py  (schema 1.2)
===============================
Extrae un golden reference desde el libro Excel ya calculado por las macros y lo
guarda en tests/golden/data/:

    golden_YYYY_MM_<label>.json           KPI, parámetros efectivos, Daily, Annual_AVA,
                                          todos los eventos de ListOfFaults, resumen N:Q
    golden_YYYY_MM_<label>_calc.json.gz   tabla Calculation-Availability!E4:BO (nivel 2)

Uso:
    python tests/golden/extract_golden.py \
        --excel "data/AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm" \
        --month 9 --year 2026 --label september

Prerequisito: el libro debe tener cmdCalcAvailability, mcoCreateList y
mcoDailyAvailability ya ejecutadas para el período (valores cacheados en el archivo).

Los valores se leen del XML crudo del libro (no con openpyxl) para conservar los
seriales de fecha exactos (design.md ADR-03).

El script FALLA (exit 1) si el libro no es internamente coherente; ver validate_golden().
Los invariantes que dependen de flags (C21, C31/L14) solo se exigen cuando aplican
(audit.md F-05, F-08).
"""
from __future__ import annotations

import argparse
import gzip
import json
import pathlib
import re
import sys
import zipfile
from datetime import datetime, timedelta
from xml.etree.ElementTree import iterparse

SCHEMA_VERSION = "1.2"
LEGACY_EVENT_CAP = 295  # max_row=300 del extractor antiguo (filas 6..300)
TOL_EXACT = 1e-9
TOL_SUM = 1e-6
TOL_KPI = 1e-12
EXCEL_EPOCH = datetime(1899, 12, 30)
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

SHEET_CALC = "Calculation-Availability"
SHEET_FAULTS = "ListOfFaults"
SHEET_DAILY = "Daily"
SHEET_ANNUAL = "Annual_AVA"


class GoldenInvariantError(RuntimeError):
    """El libro Excel no supera los invariantes de coherencia interna."""


# --------------------------------------------------------------------------- XML


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


class RawWorkbook:
    """Lector mínimo de valores cacheados del .xlsm (sin conversión de fechas)."""

    def __init__(self, path: str):
        self.zip = zipfile.ZipFile(path)
        self.shared = self._shared_strings()
        self.sheet_files = self._sheet_files()

    def _shared_strings(self) -> list[str]:
        if "xl/sharedStrings.xml" not in self.zip.namelist():
            return []
        out = []
        with self.zip.open("xl/sharedStrings.xml") as fh:
            for _, el in iterparse(fh):
                if el.tag == NS + "si":
                    out.append("".join(t.text or "" for t in el.iter(NS + "t")))
                    el.clear()
        return out

    def _sheet_files(self) -> dict[str, str]:
        rels = {}
        with self.zip.open("xl/_rels/workbook.xml.rels") as fh:
            for _, el in iterparse(fh):
                if el.tag.endswith("Relationship"):
                    rels[el.get("Id")] = el.get("Target")
        out = {}
        with self.zip.open("xl/workbook.xml") as fh:
            for _, el in iterparse(fh):
                if el.tag == NS + "sheet":
                    target = rels[el.get(REL_NS + "id")].lstrip("/")
                    out[el.get("name")] = target if target.startswith("xl/") else "xl/" + target
        return out

    def read(self, sheet: str, min_row: int = 1, max_row: int | None = None,
             min_col: int = 1, max_col: int | None = None) -> dict[tuple[int, int], object]:
        """Devuelve {(fila, col): valor} para las celdas con valor dentro del rango."""
        cells: dict[tuple[int, int], object] = {}
        ref_re = re.compile(r"([A-Z]+)(\d+)")
        with self.zip.open(self.sheet_files[sheet]) as fh:
            for _, el in iterparse(fh):
                if el.tag != NS + "c":
                    continue
                col_l, row_s = ref_re.fullmatch(el.get("r")).groups()
                row, col = int(row_s), _col_index(col_l)
                if (row < min_row or (max_row and row > max_row)
                        or col < min_col or (max_col and col > max_col)):
                    el.clear()
                    continue
                t = el.get("t")
                v = el.find(NS + "v")
                if t == "inlineStr":
                    val = "".join(x.text or "" for x in el.iter(NS + "t"))
                elif v is None or v.text is None:
                    val = None
                elif t == "s":
                    val = self.shared[int(v.text)]
                elif t in ("str", "e"):
                    val = v.text
                elif t == "b":
                    val = v.text == "1"
                else:
                    val = float(v.text)
                if val is not None:
                    cells[(row, col)] = val
                el.clear()
        return cells


def _serial_to_str(s: float | None) -> str | None:
    if not isinstance(s, float):
        return None
    dt = EXCEL_EPOCH + timedelta(days=s)
    dt = (dt + timedelta(microseconds=500_000)).replace(microsecond=0)  # al segundo más cercano
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def _serial_to_date(s: float | None) -> str | None:
    txt = _serial_to_str(s)
    return txt[:10] if txt else None


def _num(v):
    """'' (IFERROR con texto vacío) y 'N/A' -> None; floats se conservan."""
    return v if isinstance(v, float) else None


def _int(v):
    return int(v) if isinstance(v, float) else v


# Reglas de texto Excel: fuente única en etl_arena.excel_semantics (ADR-04; formato General de
# 15 dígitos significativos). Se re-exportan porque test_golden_integrity las importa desde aquí.
from etl_arena.excel_semantics import codigo_falla_excel, texto_excel  # noqa: E402,F401


def _fecha_guardado(excel_path: str) -> str | None:
    """``dcterms:modified`` de docProps/core.xml: fecha del libro, no del momento de extraer
    (el mismo libro produce siempre el mismo JSON)."""
    with zipfile.ZipFile(excel_path) as z:
        if "docProps/core.xml" not in z.namelist():
            return None
        xml = z.read("docProps/core.xml").decode("utf-8")
    m = re.search(r"<dcterms:modified[^>]*>([^<]+)</dcterms:modified>", xml)
    return m.group(1) if m else None


# --------------------------------------------------------------------- extracción


def extract(excel_path: str, year: int, month: int, label: str) -> tuple[dict, dict]:
    wb = RawWorkbook(excel_path)
    total_pcs = _int(wb.read(SHEET_CALC, min_row=2, max_row=2, min_col=3, max_col=3).get((2, 3)))
    if not isinstance(total_pcs, int) or total_pcs <= 0:
        raise GoldenInvariantError(f"Calculation-Availability!C2 inválido: {total_pcs!r}")
    col_bo = 5 + total_pcs + 1  # E = serial, F.. = PCS 1..C2, luego BO (factor excusable)
    calc = wb.read(SHEET_CALC, max_col=col_bo)
    lof = wb.read(SHEET_FAULTS, max_col=17)
    daily_c = wb.read(SHEET_DAILY, max_row=40, max_col=8)
    annual_c = wb.read(SHEET_ANNUAL, min_row=7, max_row=18, max_col=11)

    def c(cells, ref):
        col_l, row_s = re.fullmatch(r"([A-Z]+)(\d+)", ref).groups()
        return cells.get((int(row_s), _col_index(col_l)))

    if c(calc, "C12") is None:
        raise GoldenInvariantError("C12 vacío — ejecutar cmdCalcAvailability antes de extraer")

    effective = {
        "C5_period_start": _serial_to_date(c(calc, "C5")),
        "C7_period_end": _serial_to_date(c(calc, "C7")),
        "C21_only_operational_time": c(calc, "C21"),
        "C31_apply_excused_event": c(calc, "C31"),
        "C23_sampling_minutes": c(calc, "C23"),
        "L2_events_start": _serial_to_date(c(lof, "L2")),
        "L4_events_end": _serial_to_date(c(lof, "L4")),
        "L14_events_apply_excused_event": c(lof, "L14"),
        "D5_daily_end": _serial_to_date(c(daily_c, "D5")),
    }
    c5 = effective["C5_period_start"]
    if c5 is None or (int(c5[:4]), int(c5[5:7])) != (year, month):
        raise GoldenInvariantError(f"--year/--month = {year}-{month:02d} no coincide con C5 = {c5}")
    period = {
        "year": year,
        "month": month,
        "month_name": label,
        "period_start": effective["C5_period_start"],
        "period_end": effective["C7_period_end"],
        "parameters": {
            "total_pcs": total_pcs,
            "batteries_per_pcs": _int(c(calc, "C3")),
            "total_racks": _int(c(calc, "C11")),
            "sampling_minutes": _int(c(calc, "C23")),
            "only_operational_time": c(calc, "C21"),
            "apply_excused_event": c(calc, "C31"),
            "algorithm_version": "availability-v1-excel-parity",
            "source_file": pathlib.Path(excel_path).name,
        },
        "kpi": {
            "c12_sample_blocks": _int(c(calc, "C12")),
            "c14_unavailable_rack_blocks": c(calc, "C14"),
            "c16_availability_period": _num(c(calc, "C16")),
            "c19_accumulated_annual": _num(c(calc, "C19")),
            "l10_total_unavailable_rack_hours": c(lof, "L10"),
        },
    }

    daily = []
    for row in range(9, 40):
        d = c(daily_c, f"C{row}")
        if not isinstance(d, float):
            continue
        daily.append({
            "day_number": _int(c(daily_c, f"B{row}")),
            "date": _serial_to_date(d),
            "daily_unavailable_rack_blocks": _num(c(daily_c, f"D{row}")),
            "accumulated_unavailable_rack_blocks": _num(c(daily_c, f"E{row}")),
            "availability": _num(c(daily_c, f"F{row}")),
            "variation": _num(c(daily_c, f"G{row}")),
        })

    annual = {}
    for row in range(7, 19):
        mn = c(annual_c, f"B{row}")
        unav, avail = c(annual_c, f"F{row}"), c(annual_c, f"G{row}")
        if mn is None or (_num(unav) is None and _num(avail) is None):
            continue
        annual[str(_int(mn))] = {
            "month_name": c(annual_c, f"C{row}"),
            "days_in_month": c(annual_c, f"D{row}"),
            "sampling_blocks": c(annual_c, f"E{row}"),
            "unavailable_rack_blocks": _num(unav),
            "monthly_availability": _num(avail),
            "accumulated_sampling_blocks": _num(c(annual_c, f"H{row}")),
            "accumulated_unavailable_blocks": _num(c(annual_c, f"I{row}")),
            "accumulated_annual_availability": _num(c(annual_c, f"J{row}")),
            "contractual_availability": _num(c(annual_c, f"K{row}")),
        }

    # ListOfFaults B:I — todas las filas escritas, en orden de escritura (no se ordenan)
    event_rows = sorted({r for (r, col) in lof if r >= 6 and col == 2})
    events = []
    for row in event_rows:
        cs, ds = c(lof, f"C{row}"), c(lof, f"D{row}")
        events.append({
            "order_excel": row - 5,
            "pcs_number": _int(c(lof, f"B{row}")),
            "start_serial": cs,
            "end_serial": ds,
            "start_timestamp": _serial_to_str(cs),
            "end_timestamp": _serial_to_str(ds),
            "duration_hours": _num(c(lof, f"E{row}")),
            "fault_code": c(lof, f"F{row}"),
            "fault_description": c(lof, f"G{row}"),
            "average_batteries_involved": _num(c(lof, f"H{row}")),
            "unavailable_rack_hours": _num(c(lof, f"I{row}")),
            "incomplete": ds is None,
        })

    code_summary = []
    for row in range(6, 173):
        code = c(lof, f"N{row}")
        if code is None:
            continue
        code_summary.append({
            "rank": row - 5,
            "fault_code": code,
            "description_pe": c(lof, f"O{row}"),
            "unavailable_rack_hours": _num(c(lof, f"P{row}")),
            "percentage": _num(c(lof, f"Q{row}")),
        })

    # Calculation-Availability E4:BO — una fila por fila de RawData procesada
    calc_rows = []
    row = 4
    while isinstance(c(calc, f"E{row}"), float):
        vals = [_num(calc.get((row, 5 + j))) for j in range(1, total_pcs + 1)]
        calc_rows.append([calc[(row, 5)], vals, _num(calc.get((row, col_bo)))])
        row += 1

    base = f"golden_{year}_{month:02d}_{label}"
    data = {
        "_meta": {
            "schema_version": SCHEMA_VERSION,
            "description": "Golden reference extraída de los valores cacheados del libro (XML crudo)",
            "extraction_method": "extract_golden.py RawWorkbook (xl/worksheets/*.xml, sin conversión de fechas)",
            "source_saved_at": _fecha_guardado(excel_path),
            "source_file": pathlib.Path(excel_path).name,
            "calc_table_file": f"{base}_calc.json.gz",
        },
        "period": period,
        "effective_parameters": effective,
        "daily": daily,
        "annual_ava_snapshot": annual,
        "fault_events_summary": {
            "total_events": len(events),
            "total_unavailable_rack_hours": c(lof, "L10"),
            "events": events,
        },
        "fault_code_summary": code_summary,
    }
    calc_table = {
        "description": (
            "Calculation-Availability!E4:BO — "
            "[serial E, [F..BN baterías ponderadas por PCS, null=vacío], BO factor excusable]"
        ),
        "total_pcs": total_pcs,
        "rows": calc_rows,
    }
    return data, calc_table


# ------------------------------------------------------------------- invariantes


def validate_golden(data: dict, calc_table: dict | None = None) -> tuple[list[str], list[str]]:
    """Invariantes de coherencia interna. Devuelve (errores, avisos)."""
    errors: list[str] = []
    warnings: list[str] = []

    kpi = data["period"]["kpi"]
    params = data["period"]["parameters"]
    eff = data.get("effective_parameters", {})
    c12, c14, c16 = kpi["c12_sample_blocks"], kpi["c14_unavailable_rack_blocks"], kpi["c16_availability_period"]
    racks, sampling = params["total_racks"], params["sampling_minutes"]
    c21_no = params.get("only_operational_time") == "No"
    daily = data["daily"]

    # KPI
    if racks and c12:
        c16_calc = 1.0 - c14 / (racks * c12)
        if c16 is None or abs(c16 - c16_calc) > TOL_KPI:
            errors.append(f"C16={c16} != 1-C14/(Racks*C12)={c16_calc}")

    # Daily (F-08: sin factor operacional → Σdaily = C14 solo si C21 = "No")
    if not daily:
        errors.append("daily vacío")
    else:
        s = sum(d["daily_unavailable_rack_blocks"] or 0 for d in daily)
        last = daily[-1]["accumulated_unavailable_rack_blocks"]
        covers = eff.get("D5_daily_end") is None or eff.get("D5_daily_end") >= (data["period"]["period_end"] or "")
        if c21_no and covers:
            if abs(s - c14) > TOL_SUM:
                errors.append(f"Σdaily={s!r} != C14={c14!r}")
            if last is None or abs(last - c14) > TOL_SUM:
                errors.append(f"último acumulado={last!r} != C14={c14!r}")
        else:
            warnings.append("Σdaily = C14 no aplica (C21 != 'No' o Daily!D5 < C7)")
        n_days = len(daily)
        if c21_no and sampling and c12 == n_days * 24 * 60 / sampling:
            lda = daily[-1]["availability"]
            if lda is None or abs(lda - c16) > TOL_SUM:
                errors.append(f"disponibilidad último día={lda} != C16={c16}")
        else:
            warnings.append(
                "disponibilidad último día = C16 no aplica (C12 != días·bloques/día: período parcial o DST)"
            )

    # Eventos
    ev = data["fault_events_summary"]
    events = ev.get("events") or ev.get("events_first_20") or []
    if ev["total_events"] == LEGACY_EVENT_CAP:
        warnings.append(f"total_events={LEGACY_EVENT_CAP} = cap del extractor antiguo; posible truncamiento")
    block_h = sampling / 60.0
    rh_sum = 0.0
    for e in events:
        tag = f"evento orden={e.get('order_excel')} pcs={e.get('pcs_number')}"
        if e.get("incomplete"):
            warnings.append(f"{tag}: registro incompleto (sin fin) — posible arrastre VBA F-06")
            continue
        fc, fd = e.get("fault_code"), e.get("fault_description")
        if fc != codigo_falla_excel(fd):
            errors.append(f"{tag}: F={fc!r} != código emulado de G={fd!r}")
        dur, avg, rh = e["duration_hours"], e["average_batteries_involved"], e["unavailable_rack_hours"]
        if abs(dur / block_h - round(dur / block_h)) > 1e-6:
            errors.append(f"{tag}: duración {dur} no es múltiplo de {block_h} h")
        if abs(12.0 * dur * avg - rh) > TOL_EXACT:
            errors.append(f"{tag}: I={rh!r} != 12*E*H={12.0 * dur * avg!r}")
        if e.get("start_serial") is not None and abs(24 * (e["end_serial"] - e["start_serial"]) - dur) > TOL_EXACT:
            errors.append(f"{tag}: E != 24*(D-C)")
        rh_sum = rh_sum + 12.0 * dur * avg
    l10 = ev.get("total_unavailable_rack_hours")
    if "events" in ev and l10 is not None and abs(rh_sum - l10) > TOL_SUM:
        errors.append(f"Σ(12·E·H)={rh_sum!r} != L10={l10!r}")

    # Resumen por código (N:Q): Σ P = L10 salvo códigos fuera del catálogo
    summary = data.get("fault_code_summary")
    if summary and l10:
        sp = sum(x["unavailable_rack_hours"] or 0 for x in summary)
        if abs(sp - l10) > TOL_SUM:
            warnings.append(f"Σ resumen N:Q={sp} != L10={l10} (códigos de eventos fuera del catálogo PCS-Fault)")

    # Invariante cruzado C14 = 4·L10 (R13.8) — solo si ambos motores usan lo mismo
    if eff and l10 is not None:
        same = (eff["C5_period_start"] == eff["L2_events_start"] and eff["C7_period_end"] == eff["L4_events_end"]
                and (eff["C31_apply_excused_event"] == "Yes") == (eff["L14_events_apply_excused_event"] == "Yes")
                and c21_no)
        delta = c14 - 4 * l10
        if same and abs(delta) > TOL_SUM:
            errors.append(f"C14 != 4·L10 (Δ={delta!r}) con parámetros equivalentes — posible arrastre F-06")
        elif not same:
            warnings.append(f"C14 = 4·L10 no aplica (parámetros KPI ≠ eventos); Δ observado = {delta!r}")

    # Tabla de resultados (nivel 2)
    if calc_table is not None:
        rows = calc_table["rows"]
        if len(rows) != c12:
            errors.append(f"filas de la tabla de resultados={len(rows)} != C12={c12}")
        if c21_no:
            s12 = 0.0
            for _, vals, _ in rows:
                for v in vals:
                    s12 = s12 + 12 * (v or 0.0)
            if abs(s12 - c14) > TOL_SUM:
                errors.append(f"Σ 12·tabla={s12!r} != C14={c14!r}")

    return errors, warnings


# ---------------------------------------------------------------------------- CLI


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    ap.add_argument("--month", required=True, type=int)
    ap.add_argument("--year", required=True, type=int)
    ap.add_argument("--label", required=True, help="Nombre del mes en minúsculas, ej: september")
    ap.add_argument("--out-dir", default=str(pathlib.Path(__file__).parent / "data"))
    ap.add_argument("--force", action="store_true",
                    help="Escribe el JSON aunque fallen los invariantes (no recomendado)")
    args = ap.parse_args()

    try:
        data, calc_table = extract(args.excel, args.year, args.month, args.label)
    except GoldenInvariantError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)

    errors, warnings = validate_golden(data, calc_table)
    for w in warnings:
        print(f"AVISO: {w}", file=sys.stderr)
    if errors:
        for e in errors:
            print(f"INVARIANTE ROTO: {e}", file=sys.stderr)
        if not args.force:
            print("No se escribió el golden. Ejecutar las macros en el libro o usar --force.", file=sys.stderr)
            sys.exit(1)
    data["_meta"]["invariants_passed"] = not errors
    data["_meta"]["invariant_errors"] = errors
    data["_meta"]["invariant_warnings"] = warnings

    out_dir = pathlib.Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"golden_{args.year}_{args.month:02d}_{args.label}"
    out_file = out_dir / f"{base}.json"
    out_file.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    # mtime=0: archivo reproducible (sin diffs espurios en git)
    with gzip.GzipFile(out_dir / f"{base}_calc.json.gz", "wb", mtime=0) as fh:
        fh.write(json.dumps(calc_table).encode("utf-8"))

    kpi = data["period"]["kpi"]
    print(f"Guardado: {out_file} (+ {base}_calc.json.gz)")
    print(f"  C12={kpi['c12_sample_blocks']}  C14={kpi['c14_unavailable_rack_blocks']!r}"
          f"  C16={kpi['c16_availability_period']!r}")
    print(f"  Eventos={data['fault_events_summary']['total_events']}  L10={kpi['l10_total_unavailable_rack_hours']!r}"
          f"  Daily={len(data['daily'])} días  Tabla={len(calc_table['rows'])} filas")
    print(f"  Invariantes: {'OK' if not errors else 'FALLIDOS (--force)'}")


if __name__ == "__main__":
    main()

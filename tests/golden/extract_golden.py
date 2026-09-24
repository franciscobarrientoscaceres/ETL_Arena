#!/usr/bin/env python3
"""
extract_golden.py
=================
Script para extraer un golden reference desde el Excel y guardarlo en
tests/golden/data/golden_YYYY_MM_<name>.json

Uso:
    python tests/golden/extract_golden.py \
        --excel "data/AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm" \
        --month 8 --year 2026 --label "august"

Prerequisito: el libro debe tener la macro VBA ya ejecutada para el mes
indicado (C5=periodo_inicio, C7=periodo_fin, C12 y C14 ya calculados;
mcoCreateList y mcoDailyAvailability corridos en la misma sesion).

El script FALLA si el libro no es internamente coherente (invariantes):
  - sum(daily.unavailable) ~= C14
  - last_accumulated ~= C14
  - C16 ~= 1 - C14/(Total_Racks * C12)
  - last daily availability ~= C16 (solo cuando el periodo cubre el ultimo
    dia del mes con datos completos; en periodos parciales se omite)
  - orientacion fault_code (corto) vs fault_description (larga)
  - duration multiple de sampling_minutes/60
  - unavailable_rack_hours ~= 12 * duration * average_batteries
"""
import argparse
import json
import pathlib
import re
import sys
from datetime import datetime

import openpyxl

LEGACY_EVENT_CAP = 295  # max_row=300 del extractor antiguo (filas 6..300)
TOL_DAILY = 1e-6
TOL_KPI = 1e-12
TOL_RH = 1e-3
CODE_RE = re.compile(r"^F\d+")


class GoldenInvariantError(RuntimeError):
    """El libro Excel no supera los invariantes de coherencia interna."""


def _fmt(x):
    return f"{x:.12g}"


def validate_golden(data: dict) -> list[str]:
    """Valida invariantes sobre un golden ya extraido. Devuelve avisos (no fatales)."""
    errors: list[str] = []
    warnings: list[str] = []

    kpi = data["period"]["kpi"]
    params = data["period"]["parameters"]
    c12 = kpi["c12_sample_blocks"]
    c14 = kpi["c14_unavailable_rack_blocks"]
    c16 = kpi["c16_availability_period"]
    total_racks = params["total_racks"]
    sampling = params["sampling_minutes"]
    daily = data["daily"]
    period_end = data["period"].get("period_end")
    month = data["period"]["month"]
    year = data["period"]["year"]

    if not daily:
        errors.append("daily vacio")
    else:
        s = sum(d["daily_unavailable_rack_blocks"] or 0 for d in daily)
        if abs(s - c14) > TOL_DAILY:
            errors.append(
                f"sum(daily)={_fmt(s)} != C14={_fmt(c14)} (delta={_fmt(s - c14)})"
            )
        last = daily[-1]
        la = last["accumulated_unavailable_rack_blocks"]
        if la is None or abs(la - c14) > TOL_DAILY:
            errors.append(
                f"last_accumulated={_fmt(la) if la is not None else None} != C14={_fmt(c14)}"
            )

        # availability del ultimo dia == C16 solo si el periodo termina en el
        # ultimo dia del mes (o el daily no tiene dias posteriores al fin de periodo
        # con denominador creciente en vacio).
        if period_end:
            pe = datetime.strptime(period_end, "%Y-%m-%d")
            last_date = datetime.strptime(last["date"], "%Y-%m-%d")
            import calendar

            last_of_month = datetime(year, month, calendar.monthrange(year, month)[1])
            partial_tail = last_date.date() > pe.date() or pe.date() < last_of_month.date()
            if not partial_tail:
                lda = last["availability"]
                if lda is None or abs(lda - c16) > TOL_DAILY:
                    errors.append(
                        f"last daily availability={_fmt(lda) if lda is not None else None} != C16={_fmt(c16)}"
                    )
            else:
                warnings.append(
                    "periodo parcial: se omite invariante last_daily.availability == C16 "
                    f"(fin periodo={period_end}, ultimo daily={last['date']})"
                )

    c16_calc = 1.0 - c14 / (total_racks * c12) if total_racks and c12 else None
    if c16_calc is not None and abs(c16 - c16_calc) > TOL_KPI:
        errors.append(f"C16={_fmt(c16)} != 1-C14/(Racks*C12)={_fmt(c16_calc)}")

    ev = data["fault_events_summary"]
    n_events = ev["total_events"]
    events = ev.get("events_first_20") or []
    if n_events == LEGACY_EVENT_CAP:
        warnings.append(
            f"total_events={n_events} == cap historico del extractor (max_row=300); "
            "posible truncamiento — re-extraer con este script"
        )

    block_h = sampling / 60.0
    for i, e in enumerate(events):
        tag = f"event[{i}] pcs={e.get('pcs_number')}"
        fc = str(e.get("fault_code") or "")
        fd = str(e.get("fault_description") or "")
        if not CODE_RE.match(fc):
            errors.append(f"{tag}: fault_code no parece codigo corto: {fc!r}")
        if len(fd) <= len(fc):
            errors.append(
                f"{tag}: fault_description no parece descripcion larga: {fd!r} vs code {fc!r}"
            )
        if fd and fc and not fd.startswith(fc):
            warnings.append(f"{tag}: description no empieza con code ({fc!r} vs {fd!r})")

        dur = e.get("duration_hours")
        avg = e.get("average_batteries_involved")
        rh = e.get("unavailable_rack_hours")
        if dur is not None:
            blocks = dur / block_h
            if abs(blocks - round(blocks)) > 1e-6:
                errors.append(f"{tag}: duration {_fmt(dur)} no es multiple de {_fmt(block_h)}")
        if dur is not None and avg is not None and rh is not None:
            exp = 12.0 * dur * avg
            if abs(exp - rh) > TOL_RH:
                errors.append(
                    f"{tag}: rh={_fmt(rh)} != 12*dur*avg={_fmt(exp)} (delta={_fmt(exp - rh)})"
                )

        st, et = e.get("start_timestamp"), e.get("end_timestamp")
        if st and et and dur is not None:
            from datetime import datetime as _dt

            t0 = _dt.strptime(st, "%Y-%m-%d %H:%M:%S")
            t1 = _dt.strptime(et, "%Y-%m-%d %H:%M:%S")
            wall = (t1 - t0).total_seconds() / 3600.0
            if abs(wall - dur) > block_h / 2 + 1e-6:
                errors.append(
                    f"{tag}: wall(end-start)={_fmt(wall)}h vs duration={_fmt(dur)}h"
                )

    # Annual snapshot: el mes del periodo con fila linkada a C14 debe coincidir;
    # meses historicos pueden divergir (documentado, no fatal).
    ann = data.get("annual_ava_snapshot") or {}
    mkey = str(month)
    if mkey in ann:
        au = ann[mkey].get("unavailable_rack_blocks")
        if au is not None and abs(au - c14) > TOL_RH:
            warnings.append(
                f"annual[{mkey}].unavailable={_fmt(au)} != C14={_fmt(c14)} "
                "(snapshot historico / celda linkada — ver GT-5)"
            )

    return errors, warnings


def extract(excel_path: str, year: int, month: int, label: str) -> dict:
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    try:
        ws_calc = wb["Calculation-Availability"]
        ws_daily = wb["Daily"]
        ws_faults = wb["ListOfFaults"]
        ws_annual = wb["Annual_AVA"]

        def dt(v):
            return v.strftime("%Y-%m-%d %H:%M:%S") if isinstance(v, datetime) else str(v) if v else None

        period = {
            "year": year,
            "month": month,
            "month_name": label,
            "period_start": dt(ws_calc.cell(5, 3).value)[:10] if ws_calc.cell(5, 3).value else None,
            "period_end": dt(ws_calc.cell(7, 3).value)[:10] if ws_calc.cell(7, 3).value else None,
            "parameters": {
                "total_pcs": int(ws_calc.cell(2, 3).value),
                "batteries_per_pcs": int(ws_calc.cell(3, 3).value),
                "total_racks": int(ws_calc.cell(11, 3).value),
                "sampling_minutes": int(ws_calc.cell(23, 3).value),
                "only_operational_time": ws_calc.cell(21, 3).value,
                "apply_excused_event": ws_calc.cell(31, 3).value,
                "algorithm_version": "availability-v1-excel-parity",
                "source_file": pathlib.Path(excel_path).name,
                "extraction_date": datetime.now().strftime("%Y-%m-%d"),
            },
            "kpi": {
                "c12_sample_blocks": int(ws_calc.cell(12, 3).value),
                "c14_unavailable_rack_blocks": ws_calc.cell(14, 3).value,
                "c16_availability_period": ws_calc.cell(16, 3).value,
                "c19_accumulated_annual": ws_calc.cell(19, 3).value,
            },
        }

        if period["kpi"]["c12_sample_blocks"] is None:
            raise GoldenInvariantError(
                "C12 vacio — ejecutar cmdCalcAvailability en el libro antes de extraer"
            )

        daily = []
        skipped_empty_dates = 0
        for row in range(9, 40):
            c = ws_daily.cell(row, 3).value
            if c is None:
                skipped_empty_dates += 1
                if skipped_empty_dates >= 3:
                    break
                continue
            skipped_empty_dates = 0
            daily.append(
                {
                    "day_number": row - 8,
                    "date": dt(c)[:10] if c else None,
                    "daily_unavailable_rack_blocks": ws_daily.cell(row, 4).value,
                    "accumulated_unavailable_rack_blocks": ws_daily.cell(row, 5).value,
                    "availability": ws_daily.cell(row, 6).value,
                    "variation": ws_daily.cell(row, 7).value,
                }
            )

        annual = {}
        for row in range(7, 19):
            mn = ws_annual.cell(row, 2).value
            if not mn:
                continue
            unav = ws_annual.cell(row, 6).value
            avail = ws_annual.cell(row, 7).value
            if unav is None and avail is None:
                continue
            annual[str(int(mn))] = {
                "month_name": ws_annual.cell(row, 3).value,
                "days_in_month": ws_annual.cell(row, 4).value,
                "sampling_blocks": ws_annual.cell(row, 5).value,
                "unavailable_rack_blocks": unav,
                "monthly_availability": avail,
                "accumulated_sampling_blocks": ws_annual.cell(row, 8).value,
                "accumulated_unavailable_blocks": ws_annual.cell(row, 9).value,
            }

        total_rh = ws_faults.cell(10, 12).value
        events = []
        # Escanear hasta el final real de la hoja; cortar tras N filas vacias
        # consecutivas en la columna B (PCS) una vez iniciados eventos.
        max_scan = max(ws_faults.max_row or 0, 1000)
        empty_run = 0
        for r in ws_faults.iter_rows(min_row=6, max_row=max_scan, values_only=True):
            b = r[1] if len(r) > 1 else None
            if b is None:
                empty_run += 1
                if events and empty_run > 50:
                    break
                continue
            empty_run = 0
            c, d, e, f, g, h, i = r[2], r[3], r[4], r[5], r[6], r[7], r[8]
            events.append(
                {
                    "pcs_number": int(b),
                    "start_timestamp": dt(c),
                    "end_timestamp": dt(d),
                    "duration_hours": e,
                    # col F = codigo corto (ej "F55"), col G = descripcion larga
                    "fault_code": f,
                    "fault_description": g,
                    "average_batteries_involved": h,
                    "unavailable_rack_hours": i,
                }
            )

        data = {
            "_meta": {
                "description": "Golden reference extraida directamente del Excel con openpyxl data_only=True",
                "extraction_method": "openpyxl read_only=True, data_only=True",
                "extraction_date": datetime.now().strftime("%Y-%m-%d"),
                "tolerance_internal": 1e-9,
                "tolerance_kpi": 1e-6,
                "tolerance_rack_hours": 1e-3,
                "invariants": [
                    "sum(daily.unavailable) ~= C14",
                    "last_accumulated ~= C14",
                    "C16 ~= 1 - C14/(Total_Racks*C12)",
                    "fault_code corto / fault_description larga",
                    "duration multiple de sampling",
                    "rh ~= 12 * duration * average_batteries",
                ],
                "note": (
                    f"Corrida VBA con Only Operational Time={period['parameters']['only_operational_time']}, "
                    f"Excusable Event={period['parameters']['apply_excused_event']}"
                ),
            },
            "period": period,
            "daily": daily,
            "annual_ava_snapshot": annual,
            "fault_events_summary": {
                "total_events": len(events),
                "total_unavailable_rack_hours": total_rh,
                "events_first_20": events[:20],
            },
        }
        return data
    finally:
        wb.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    ap.add_argument("--month", required=True, type=int)
    ap.add_argument("--year", required=True, type=int)
    ap.add_argument("--label", required=True, help="Nombre del mes en minusculas, ej: july")
    ap.add_argument(
        "--force",
        action="store_true",
        help="Escribe el JSON aunque fallen los invariantes (no recomendado)",
    )
    args = ap.parse_args()

    try:
        data = extract(args.excel, args.year, args.month, args.label)
    except GoldenInvariantError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)

    errors, warnings = validate_golden(data)
    for w in warnings:
        print(f"AVISO: {w}", file=sys.stderr)
    if errors:
        for e in errors:
            print(f"INVARIANTE ROTO: {e}", file=sys.stderr)
        if not args.force:
            print(
                "No se escribio el golden. Correr la macro en el libro o usar --force.",
                file=sys.stderr,
            )
            sys.exit(1)
        data["_meta"]["invariants_passed"] = False
        data["_meta"]["invariant_errors"] = errors
    else:
        data["_meta"]["invariants_passed"] = True

    out_dir = pathlib.Path(__file__).parent / "data"
    out_dir.mkdir(exist_ok=True)
    out_file = out_dir / f"golden_{args.year}_{args.month:02d}_{args.label}.json"
    out_file.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Guardado: {out_file}")
    kpi = data["period"]["kpi"]
    print(f"  C12={kpi['c12_sample_blocks']}  C16={kpi['c16_availability_period']:.8f}")
    print(
        f"  Events={data['fault_events_summary']['total_events']}  "
        f"RackHours={data['fault_events_summary']['total_unavailable_rack_hours']:.4f}"
    )
    print(f"  Invariantes: {'OK' if data['_meta'].get('invariants_passed') else 'FALLIDOS (--force)'}")


if __name__ == "__main__":
    main()

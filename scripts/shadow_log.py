"""Registro del shadow mode (tarea 5.4): agrega una fila a ``docs/shadow-log.md`` desde ``run_state.json``.

    python scripts/shadow_log.py --corte 2026-09-28 --kpi-excel-oficial 0.9819 --nota "sin novedades"
    python scripts/shadow_log.py --resumen

``--resumen`` evalúa el criterio de salida: 4 semanales consecutivas en ``success`` con reconciliación
``pass`` (o nota que explique la diferencia) y al menos 1 cierre mensual en ``success``.
Código de salida: 0 = ok (con ``--resumen``: criterio cumplido) · 1 = error o criterio aún no cumplido.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
LOG = RAIZ / "docs" / "shadow-log.md"
SEMANAS_REQUERIDAS = 4


def _num(v, fmt: str) -> str:
    return "" if v is None else format(v, fmt)


def fila_desde_estado(dir_corte: Path, corte: str, kpi_excel: float | None, nota: str) -> str:
    ruta = dir_corte / "run_state.json"
    if not ruta.exists():
        raise FileNotFoundError(f"no existe {ruta}: ¿corrió run_lunes para el corte {corte}?")
    estado = json.loads(ruta.read_text(encoding="utf-8"))
    etl = estado.get("run-etl", {})
    if etl.get("estado") != "ok":
        raise RuntimeError(f"run-etl del corte {corte} no terminó ok ({etl.get('estado')}): {etl.get('error')}")
    a = etl["artefactos"]
    macros = estado.get("run-macros", {}).get("artefactos", {})
    c14_excel = macros.get("C14") if macros.get("referencia", True) is not False else None
    rec = (a.get("reconciliacion") or {}).get("estado", "sin_referencia")
    c16 = a["kpi"].get("C16")
    delta = None if kpi_excel is None or c16 is None else c16 - kpi_excel
    etiqueta = "Con Exclusiones" if a.get("estado_exclusiones") == "con_exclusiones" else "Sin Exclusiones"
    celdas = [
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        corte,
        a.get("tipo", "semanal"),
        f"{a['periodo'][0]} → {a['periodo'][1]}",
        etiqueta,
        a["estado"],
        rec,
        str(a["kpi"].get("C12", "")),
        _num(a["kpi"].get("C14"), ".3f"),
        _num(c14_excel, ".3f"),
        _num(c16, ".6f"),
        _num(kpi_excel, ".6f"),
        _num(delta, ".2e"),
        str(a.get("num_corrida") or ""),
        f"`{a['id_corrida']}`",
        nota.replace("|", "/"),
    ]
    return "| " + " | ".join(celdas) + " |"


def agregar(log: Path, fila: str) -> None:
    texto = log.read_text(encoding="utf-8")
    if not texto.endswith("\n"):
        texto += "\n"
    log.write_text(texto + fila + "\n", encoding="utf-8")


def filas_registradas(log: Path) -> list[dict]:
    """Filas de la tabla del registro (después del encabezado ``| Registrado | Corte | …``)."""
    lineas = log.read_text(encoding="utf-8").splitlines()
    try:
        inicio = next(i for i, x in enumerate(lineas) if x.startswith("| Registrado | Corte |"))
    except StopIteration:
        return []
    columnas = [c.strip() for c in lineas[inicio].strip("|").split("|")]
    salida = []
    for x in lineas[inicio + 2 :]:
        if x.startswith("|"):
            salida.append(dict(zip(columnas, (c.strip() for c in x.strip().strip("|").split("|")), strict=False)))
    return salida


def evaluar(filas: list[dict]) -> dict:
    semanales = [f for f in filas if f.get("Tipo") == "semanal"]
    racha = mejor = 0
    for f in semanales:  # en orden de registro
        explicada = f.get("Reconciliación") == "pass" or bool(f.get("Nota"))
        racha = racha + 1 if f.get("Estado") == "success" and explicada else 0
        mejor = max(mejor, racha)
    cierres = [f for f in filas if f.get("Tipo") == "cierre_mensual" and f.get("Estado") == "success"]
    return {
        "semanales_registradas": len(semanales),
        "mejor_racha_semanal": mejor,
        "cierres_success": len(cierres),
        "cierres_con_exclusiones": sum(1 for f in cierres if f.get("Etiqueta") == "Con Exclusiones"),
        "cumple": mejor >= SEMANAS_REQUERIDAS and len(cierres) >= 1,
    }


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--corte")
    ap.add_argument("--work", type=Path, default=RAIZ / "data" / "work")
    ap.add_argument("--kpi-excel-oficial", type=float, help="KPI (C16) que reportó Alex con su Excel")
    ap.add_argument("--nota", default="")
    ap.add_argument("--log", type=Path, default=LOG)
    ap.add_argument("--resumen", action="store_true")
    args = ap.parse_args(argv)
    if args.resumen:
        r = evaluar(filas_registradas(args.log))
        print(json.dumps(r, ensure_ascii=False, indent=2))
        return 0 if r["cumple"] else 1
    if not args.corte:
        print("falta --corte (o usar --resumen)", file=sys.stderr)
        return 1
    try:
        fila = fila_desde_estado(args.work / args.corte, args.corte, args.kpi_excel_oficial, args.nota)
    except (FileNotFoundError, RuntimeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    agregar(args.log, fila)
    print(fila)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

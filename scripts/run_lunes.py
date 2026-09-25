"""Orquestador semanal del lunes (tarea 4.9; R19, runbook-lunes.md).

Etapas (``--stage``): acquire-wait → prepare-workbook → run-macros → run-etl → reconcile → notify-bi,
o ``all``. El estado de cada etapa queda en ``data/work/<corte>/run_state.json``; ``--stage all``
salta las etapas ya ``ok`` (reanudable). Solo en la PC local (nunca en el server SCADA).

Mientras no exista el contrato SCADA (0.6/0.7), ``prepare-workbook`` usa ``--libro-preparado``:
el libro al que se le pegaron a mano las filas nuevas de RawData-PCS (proceso actual).

Ejemplo (lunes 2026-09-28, libro preparado a mano):
    python scripts/run_lunes.py --stage all --corte 2026-09-28 --libro-preparado C:/ruta/libro.xlsm --oficial

Código de salida: 0 = ok · 2 = parity_failed · 1 = error.
"""

from __future__ import annotations

import argparse
import dataclasses
import json
import sys
import traceback
from datetime import date, datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))

from etl_arena.config import construir_config  # noqa: E402
from etl_arena.excel_semantics import serial_a_datetime  # noqa: E402
from etl_arena.ingestion import LibroXlsx  # noqa: E402
from etl_arena.registro import configurar_logging  # noqa: E402

ETAPAS = ("acquire-wait", "prepare-workbook", "run-macros", "run-etl", "reconcile", "notify-bi")
NOMBRE_LIBRO = "libro.xlsm"


# ------------------------------------------------------------------ estado reanudable
class Estado:
    def __init__(self, dir_corte: Path):
        self.ruta = dir_corte / "run_state.json"
        self.datos: dict = json.loads(self.ruta.read_text(encoding="utf-8")) if self.ruta.exists() else {}

    def ok(self, etapa: str) -> bool:
        return self.datos.get(etapa, {}).get("estado") == "ok"

    def artefacto(self, etapa: str, clave: str):
        return self.datos.get(etapa, {}).get("artefactos", {}).get(clave)

    def registrar(
        self, etapa: str, estado: str, inicio: datetime, artefactos: dict | None = None, error: str | None = None
    ) -> None:
        self.datos[etapa] = {
            "estado": estado,
            "inicio": inicio.isoformat(timespec="seconds"),
            "fin": datetime.now().isoformat(timespec="seconds"),
            "artefactos": artefactos or {},
            "error": error,
        }
        self.ruta.parent.mkdir(parents=True, exist_ok=True)
        self.ruta.write_text(json.dumps(self.datos, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


# ------------------------------------------------------------------ período por defecto (D-07)
def ultimo_dato(libro: Path) -> date:
    """Fecha del último serial de RawData-PCS!A (el VBA corta en la primera A vacía)."""
    ultimo = None
    with LibroXlsx(libro) as lx:
        for _fila, c in lx.filas("RawData-PCS", min_fila=2, max_col=1):
            if not isinstance(c.get(1), float):
                break
            ultimo = c[1]
    if ultimo is None:
        raise RuntimeError(f"{libro.name}: RawData-PCS sin datos")
    return serial_a_datetime(ultimo).date()


def configuracion(args, libro: Path):
    fin = args.periodo_fin or ultimo_dato(libro)
    inicio = args.periodo_inicio or fin.replace(day=1)  # mes en curso hasta el último dato (D-07)
    # Parámetros oficiales (D-03, D-07): C21 = "No", C31 = L14 = "Yes", L2/L4/Daily = C5/C7
    return construir_config(
        inicio_periodo=inicio,
        fin_periodo=fin,
        tipo_corrida=args.tipo,
        es_oficial=args.oficial,
        archivo_origen=libro.name,
    )


# ------------------------------------------------------------------ etapas
def etapa_acquire(args, estado: Estado, dir_corte: Path) -> dict:
    from etl_arena.acquisition import acquire_wait

    a = acquire_wait(args.inbox, args.processed, args.corte, timeout_s=args.timeout_espera)
    return {"archivo": str(a.ruta), "sha256": a.sha256, "bytes": a.bytes}


def etapa_prepare(args, estado: Estado, dir_corte: Path) -> dict:
    from etl_arena.workbook import copiar_libro_trabajo

    if not args.libro_preparado:
        raise RuntimeError(
            "falta el contrato SCADA (tareas 0.6/0.7) para escribir el export en el libro: preparar el "
            "libro a mano (pegar las filas nuevas en RawData-PCS) y pasar --libro-preparado"
        )
    libro = copiar_libro_trabajo(args.libro_preparado, dir_corte, NOMBRE_LIBRO)
    return {"libro": str(libro), "origen": str(args.libro_preparado)}


def _libro(estado: Estado) -> Path:
    ruta = estado.artefacto("prepare-workbook", "libro")
    if not ruta:
        raise RuntimeError("no hay libro de trabajo: correr antes prepare-workbook")
    return Path(ruta)


def etapa_macros(args, estado: Estado, dir_corte: Path) -> dict:
    from etl_arena.workbook import ejecutar_macros, extraer_referencia

    libro = _libro(estado)
    cfg = configuracion(args, libro)
    tiempos = ejecutar_macros(libro, cfg, timeout_s=args.timeout_macros)
    ref = extraer_referencia(libro, args.corte)
    (dir_corte / "referencia_excel.json").write_text(
        json.dumps(dataclasses.asdict(ref), ensure_ascii=False, default=str), encoding="utf-8"
    )
    return {
        "periodo": [str(cfg.inicio_periodo), str(cfg.fin_periodo)],
        "segundos": {k: round(v, 1) for k, v in tiempos.items()},
        "C12": ref.c12,
        "C14": ref.c14,
        "C16": ref.c16,
    }


def etapa_etl(args, estado: Estado, dir_corte: Path) -> dict:
    from etl_arena.ejecucion import ejecutar_corrida
    from etl_arena.persistence import RepositorioCorridas, crear_engine
    from etl_arena.workbook import extraer_referencia, sha256_archivo

    libro = _libro(estado)
    cfg = configuracion(args, libro)
    referencia = extraer_referencia(libro, args.corte) if estado.ok("run-macros") else None
    repo = None if args.sin_bd else RepositorioCorridas(crear_engine())
    res = ejecutar_corrida(libro, cfg, repo, hash_archivo=sha256_archivo(libro), referencia=referencia)
    d = res.resultado.disponibilidad
    return {
        "id_corrida": res.id_corrida,
        "estado": res.estado,
        "periodo": [str(cfg.inicio_periodo), str(cfg.fin_periodo)],
        "estado_exclusiones": res.estado_exclusiones,
        "kpi": {"C12": d.bloques_muestreo, "C14": d.bloques_racks_indisponibles, "C16": d.disponibilidad_periodo},
        "reconciliacion": res.resumen_calidad.get("reconciliacion"),
        "anomalias": res.resumen_calidad.get("anomalias"),
    }


def etapa_reconcile(args, estado: Estado, dir_corte: Path) -> dict:
    """La reconciliación se calcula y persiste en run-etl; aquí se decide si la cadena sigue."""
    etl = estado.datos.get("run-etl", {}).get("artefactos", {})
    if not etl:
        raise RuntimeError("no hay resultado de run-etl")
    rec = (etl.get("reconciliacion") or {}).get("estado", "sin_referencia")
    if etl.get("estado") == "parity_failed":
        raise RuntimeError(
            "parity_failed: Python y Excel no cuadran; revisar reconciliation_result de "
            f"{etl['id_corrida']} (no se publica en Power BI)"
        )
    return {"reconciliacion": rec, "id_corrida": etl.get("id_corrida")}


def etapa_notify(args, estado: Estado, dir_corte: Path) -> dict:
    from etl_arena.reporting.notificacion import construir_mensaje, enviar

    etl = estado.datos.get("run-etl", {}).get("artefactos", {})
    if not etl:
        raise RuntimeError("no hay resultado de run-etl para notificar")
    mensaje = construir_mensaje(
        etl["id_corrida"],
        tuple(etl["periodo"]),
        etl["estado"],
        etl["estado_exclusiones"],
        etl["kpi"],
        etl.get("reconciliacion"),
        {"anomalias": etl.get("anomalias") or {}},
    )
    return {"destino": enviar(mensaje, dir_corte)}


FUNCIONES = {
    "acquire-wait": etapa_acquire,
    "prepare-workbook": etapa_prepare,
    "run-macros": etapa_macros,
    "run-etl": etapa_etl,
    "reconcile": etapa_reconcile,
    "notify-bi": etapa_notify,
}


# ------------------------------------------------------------------ CLI
def construir_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--stage", choices=(*ETAPAS, "all"), required=True)
    ap.add_argument("--corte", default=date.today().isoformat(), help="identificador del corte (por defecto, hoy)")
    ap.add_argument("--inbox", type=Path, default=RAIZ / "data" / "inbox")
    ap.add_argument("--processed", type=Path, default=RAIZ / "data" / "processed")
    ap.add_argument("--work", type=Path, default=RAIZ / "data" / "work")
    ap.add_argument("--libro-preparado", type=Path, help="libro con las filas nuevas ya pegadas (hasta tener 0.6/0.7)")
    ap.add_argument("--omitir-acquire", action="store_true", help="con --stage all, no esperar el export en inbox")
    ap.add_argument("--periodo-inicio", type=date.fromisoformat)
    ap.add_argument("--periodo-fin", type=date.fromisoformat)
    ap.add_argument("--tipo", choices=("semanal", "cierre_mensual", "reproceso"), default="semanal")
    ap.add_argument("--oficial", action="store_true")
    ap.add_argument("--sin-bd", action="store_true")
    ap.add_argument("--timeout-espera", type=float, default=3600)
    ap.add_argument("--timeout-macros", type=float, default=2400)
    ap.add_argument("--forzar", action="store_true", help="re-ejecutar etapas aunque estén ok")
    return ap


def main(argv: list[str] | None = None) -> int:
    args = construir_parser().parse_args(argv)
    configurar_logging(json_=True)
    dir_corte = args.work / args.corte
    estado = Estado(dir_corte)
    etapas = list(ETAPAS) if args.stage == "all" else [args.stage]
    if args.stage == "all" and args.omitir_acquire:
        etapas.remove("acquire-wait")
    for etapa in etapas:
        if args.stage == "all" and estado.ok(etapa) and not args.forzar:
            print(f"[{etapa}] ok (ya ejecutada)")
            continue
        inicio = datetime.now()
        print(f"[{etapa}] …")
        try:
            artefactos = FUNCIONES[etapa](args, estado, dir_corte)
        except Exception as exc:
            estado.registrar(etapa, "error", inicio, error=f"{type(exc).__name__}: {exc}")
            print(f"[{etapa}] ERROR: {exc}", file=sys.stderr)
            traceback.print_exc()
            return 2 if "parity_failed" in str(exc) else 1
        estado.registrar(etapa, "ok", inicio, artefactos)
        print(f"[{etapa}] ok {json.dumps(artefactos, ensure_ascii=False, default=str)[:300]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

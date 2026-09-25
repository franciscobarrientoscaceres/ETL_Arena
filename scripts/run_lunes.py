"""Orquestador semanal del lunes (tarea 4.9; R19, runbook-lunes.md).

Etapas (``--stage``): acquire-wait → prepare-workbook → run-macros → run-etl → reconcile → notify-bi,
o ``all``. El estado de cada etapa queda en ``data/work/<corte>/run_state.json``; ``--stage all``
salta las etapas ya ``ok`` (reanudable). Solo en la PC local (nunca en el server SCADA).

Entre cortes (``etl_arena.orquestacion``): una corrida ``--oficial`` exitosa y persistida promueve su
libro de trabajo a **libro base** de la siguiente (R3.1); ``run-etl`` encola el ``cierre_mensual`` de
los meses que los datos ya cubren completos (R19.3). ``--stage cierre-mensual [--mes AAAA-MM]`` corre
la cadena para el mes completo sobre una copia del libro base; exige la ``Exclusion_Matrix`` del mes
cargada o ``--sin-exclusiones`` explícito (R19.6, D-17).

Mientras no exista el contrato SCADA (0.6/0.7), ``prepare-workbook`` usa ``--libro-preparado``:
el libro al que se le pegaron a mano las filas nuevas de RawData-PCS (proceso actual).

Ejemplo (lunes 2026-09-28, libro preparado a mano):
    python scripts/run_lunes.py --stage all --corte 2026-09-28 --libro-preparado C:/ruta/libro.xlsm --oficial
Cierre de septiembre sin esperar la matriz (queda oficial "Sin Exclusiones"):
    python scripts/run_lunes.py --stage cierre-mensual --mes 2026-09 --sin-exclusiones

Código de salida: 0 = ok · 2 = parity_failed · 1 = error.
"""

from __future__ import annotations

import argparse
import copy
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
from etl_arena.orquestacion import (  # noqa: E402
    ColaCierres,
    encolar_cierres,
    libro_base,
    mes_cubierto,
    promover_libro_base,
    ultimo_dia,
)
from etl_arena.registro import configurar_logging  # noqa: E402

ETAPAS = ("acquire-wait", "prepare-workbook", "run-macros", "run-etl", "reconcile", "notify-bi")
ETAPAS_CIERRE = ETAPAS[1:]
# No se repiten con --forzar: data/processed es inmutable y la copia de trabajo no se pisa.
ETAPAS_UNA_VEZ = ("acquire-wait", "prepare-workbook")
NOMBRE_LIBRO = "libro.xlsm"
MAESTRO = RAIZ / "data" / "AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm"


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
        tmp = self.ruta.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(self.datos, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        tmp.replace(self.ruta)  # un corte a mitad de escritura no deja el estado ilegible


# ------------------------------------------------------------------ período por defecto (D-07)
def ultimo_serial(libro: Path) -> float:
    """Último serial de RawData-PCS!A (el VBA corta en la primera A vacía)."""
    ultimo = None
    with LibroXlsx(libro) as lx:
        for _fila, c in lx.filas("RawData-PCS", min_fila=2, max_col=1):
            if not isinstance(c.get(1), float):
                break
            ultimo = c[1]
    if ultimo is None:
        raise RuntimeError(f"{libro.name}: RawData-PCS sin datos")
    return ultimo


def ultimo_dato(libro: Path) -> date:
    return serial_a_datetime(ultimo_serial(libro)).date()


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
    res = ejecutar_corrida(
        libro,
        cfg,
        repo,
        hash_archivo=sha256_archivo(libro),
        referencia=referencia,
        estado_exclusiones="sin_exclusiones" if args.sin_exclusiones else None,
    )
    d = res.resultado.disponibilidad
    encolados = []
    if args.tipo == "semanal" and res.estado == "success":
        serial = ultimo_serial(libro)
        encolados = encolar_cierres(
            ColaCierres(args.work),
            serial,
            serial_a_datetime(serial).date(),
            cfg.minutos_muestreo,
            args.corte,
            desde=cfg.inicio_acumulado_anual,
            mes_cerrado=(lambda a, m: repo.mes_cerrado(cfg.id_proyecto, a, m)) if repo else None,
        )
    return {
        "id_corrida": res.id_corrida,
        "estado": res.estado,
        "periodo": [str(cfg.inicio_periodo), str(cfg.fin_periodo)],
        "estado_exclusiones": res.estado_exclusiones,
        "kpi": {"C12": d.bloques_muestreo, "C14": d.bloques_racks_indisponibles, "C16": d.disponibilidad_periodo},
        "reconciliacion": res.resumen_calidad.get("reconciliacion"),
        "anomalias": res.resumen_calidad.get("anomalias"),
        "cierres_encolados": encolados,
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
    salida = {"reconciliacion": rec, "id_corrida": etl.get("id_corrida")}
    if args.oficial and not args.sin_bd and etl.get("estado") == "success":
        from etl_arena.workbook import sha256_archivo

        libro = _libro(estado)
        salida["libro_base"] = promover_libro_base(
            args.work, libro, args.corte, etl["id_corrida"], sha256_archivo(libro)
        )
    return salida


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
        ColaCierres(args.work).pendientes(),
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
    ap.add_argument("--stage", choices=(*ETAPAS, "all", "cierre-mensual"), required=True)
    ap.add_argument("--corte", default=date.today().isoformat(), help="identificador del corte (por defecto, hoy)")
    ap.add_argument("--inbox", type=Path, default=RAIZ / "data" / "inbox")
    ap.add_argument("--processed", type=Path, default=RAIZ / "data" / "processed")
    ap.add_argument("--work", type=Path, default=RAIZ / "data" / "work")
    ap.add_argument("--libro-preparado", type=Path, help="libro con las filas nuevas ya pegadas (hasta tener 0.6/0.7)")
    ap.add_argument("--maestro", type=Path, default=MAESTRO, help="libro base de la primera corrida (R3.1)")
    ap.add_argument("--mes", help="cierre-mensual: mes AAAA-MM (por defecto, el primero pendiente en la cola)")
    ap.add_argument(
        "--sin-exclusiones",
        action="store_true",
        help="cierre-mensual oficial sin la Exclusion_Matrix del mes (R19.6); la corrida queda 'Sin Exclusiones'",
    )
    ap.add_argument("--omitir-acquire", action="store_true", help="con --stage all, no esperar el export en inbox")
    ap.add_argument(
        "--omitir-macros",
        action="store_true",
        help="con --stage all / cierre-mensual, no correr Excel: la corrida queda sin referencia (R3.7)",
    )
    ap.add_argument("--periodo-inicio", type=date.fromisoformat)
    ap.add_argument("--periodo-fin", type=date.fromisoformat)
    ap.add_argument("--tipo", choices=("semanal", "cierre_mensual", "reproceso"), default="semanal")
    ap.add_argument("--oficial", action="store_true")
    ap.add_argument("--sin-bd", action="store_true")
    ap.add_argument("--timeout-espera", type=float, default=3600)
    ap.add_argument("--timeout-macros", type=float, default=2400)
    ap.add_argument("--forzar", action="store_true", help="re-ejecutar etapas aunque estén ok")
    return ap


def cierre_mensual(args) -> int:
    """Cadena completa del mes sobre una copia del libro base; marca el mes como ejecutado en la cola."""
    cola = ColaCierres(args.work)
    mes = args.mes or next(iter(cola.pendientes()), None)
    if mes is None:
        print("[cierre-mensual] no hay cierres pendientes en la cola")
        return 0
    try:
        anio, m = (int(x) for x in mes.split("-"))
        date(anio, m, 1)
    except ValueError:
        print(f"[cierre-mensual] ERROR: --mes {mes!r} no es AAAA-MM", file=sys.stderr)
        return 1
    inicio, fin = date(anio, m, 1), ultimo_dia(anio, m)
    base = libro_base(args.work, args.maestro)
    cfg = construir_config(inicio_periodo=inicio, fin_periodo=fin)
    if not mes_cubierto(ultimo_serial(base), anio, m, cfg.minutos_muestreo):
        print(f"[cierre-mensual] ERROR: el libro base {base.name} no cubre {mes} completo", file=sys.stderr)
        return 1
    if not args.sin_exclusiones:
        cargada = False
        if not args.sin_bd:
            from etl_arena.persistence import RepositorioCorridas, crear_engine

            cargada = RepositorioCorridas(crear_engine()).exclusiones_cargadas(cfg.id_proyecto, anio, m)
        if not cargada:
            print(
                f"[cierre-mensual] ERROR: la Exclusion_Matrix de {mes} no está cargada (exclusion_matrix_carga). "
                "Esperar la entrega de Alex (load-exclusion-matrix) o usar --sin-exclusiones (R19.6, D-17)",
                file=sys.stderr,
            )
            return 1
    sub = copy.copy(args)
    sub.stage, sub.corte, sub.libro_preparado = "all", f"cierre-{mes}", base
    sub.periodo_inicio, sub.periodo_fin, sub.tipo, sub.oficial = inicio, fin, "cierre_mensual", True
    dir_corte = args.work / sub.corte
    etapas = [e for e in ETAPAS_CIERRE if not (args.omitir_macros and e == "run-macros")]
    codigo = ejecutar_etapas(sub, etapas, dir_corte)
    etl = Estado(dir_corte).datos.get("run-etl", {}).get("artefactos", {})
    if codigo == 0 and etl.get("estado") == "success":
        cola.marcar_ejecutado(mes, etl["id_corrida"], etl["estado_exclusiones"])
    return codigo


def main(argv: list[str] | None = None) -> int:
    args = construir_parser().parse_args(argv)
    configurar_logging(json_=True)
    if args.stage == "cierre-mensual":
        return cierre_mensual(args)
    etapas = list(ETAPAS) if args.stage == "all" else [args.stage]
    if args.stage == "all" and args.omitir_acquire:
        etapas.remove("acquire-wait")
    if args.stage == "all" and args.omitir_macros:
        etapas.remove("run-macros")
    return ejecutar_etapas(args, etapas, args.work / args.corte)


def ejecutar_etapas(args, etapas: list[str], dir_corte: Path) -> int:
    estado = Estado(dir_corte)
    for etapa in etapas:
        if args.stage == "all" and estado.ok(etapa) and (not args.forzar or etapa in ETAPAS_UNA_VEZ):
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

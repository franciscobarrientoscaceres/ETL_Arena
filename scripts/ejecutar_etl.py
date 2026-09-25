"""Ejecuta una corrida del ETL sobre un libro de trabajo (tarea 3.1; design.md §Orquestación).

Ejemplos:
    # parámetros tal como están en el libro, sin base de datos, reconciliando contra el mismo libro
    python scripts/ejecutar_etl.py --libro data/work/libro.xlsm --parametros-desde-libro --sin-bd --referencia libro

    # corrida semanal oficial de septiembre (mes en curso), persistida en ETL_ARENA_DB_URL
    python scripts/ejecutar_etl.py --libro data/work/libro.xlsm --oficial
        --periodo-inicio 2026-09-01 --periodo-fin 2026-09-21        (en una sola línea)

Código de salida: 0 = success · 2 = parity_failed · 1 = failed / error.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))

from etl_arena.ambientes import AYUDA_ENTORNO, OPCIONES_ENTORNO  # noqa: E402
from etl_arena.config import construir_config, parametros_desde_celdas  # noqa: E402
from etl_arena.ejecucion import ejecutar_corrida  # noqa: E402
from etl_arena.ingestion import leer_celdas_parametros  # noqa: E402
from etl_arena.registro import configurar_logging  # noqa: E402
from etl_arena.workbook import extraer_referencia, sha256_archivo  # noqa: E402

SALIDA = {"success": 0, "parity_failed": 2, "failed": 1}


def _si_no(v: str | None) -> bool | None:
    if v is None:
        return None
    if v not in ("Yes", "No"):
        raise argparse.ArgumentTypeError("usar Yes o No")
    return v == "Yes"


def _fecha(v: str | None) -> date | None:
    return date.fromisoformat(v) if v else None


def construir_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--libro", type=Path, required=True, help="libro de trabajo .xlsm ya preparado")
    ap.add_argument(
        "--parametros-desde-libro",
        action="store_true",
        help="tomar C5/C7/C21/C31/L2/L4/L14/Daily!C del libro (los argumentos explícitos mandan)",
    )
    ap.add_argument("--periodo-inicio", type=_fecha)
    ap.add_argument("--periodo-fin", type=_fecha)
    ap.add_argument("--eventos-inicio", type=_fecha)
    ap.add_argument("--eventos-fin", type=_fecha)
    ap.add_argument("--diario-fin", type=_fecha)
    ap.add_argument("--solo-operacional", type=_si_no, metavar="Yes|No")
    ap.add_argument("--excusable", type=_si_no, metavar="Yes|No", help="C31: aplicar Exclusion_Matrix al KPI")
    ap.add_argument(
        "--eventos-excusable", type=_si_no, metavar="Yes|No", help="L14: aplicar Exclusion_Matrix a eventos"
    )
    ap.add_argument("--modo-huecos", choices=("excel", "continuar"))
    ap.add_argument("--tipo", choices=("semanal", "cierre_mensual", "reproceso", "golden"), default="semanal")
    ap.add_argument("--oficial", action="store_true", help="corrida oficial (vigente en Power BI)")
    ap.add_argument(
        "--estado-exclusiones",
        choices=("sin_exclusiones", "con_exclusiones"),
        help="por defecto se deduce de exclusion_matrix_carga (D-17)",
    )
    ap.add_argument(
        "--referencia",
        default="none",
        help="'libro' (valores cacheados del mismo libro), ruta a un .xlsm con macros ejecutadas, o 'none'",
    )
    ap.add_argument("--sin-bd", action="store_true", help="calcular y reconciliar sin persistir")
    ap.add_argument("--entorno", choices=OPCIONES_ENTORNO, help=AYUDA_ENTORNO)
    ap.add_argument("--salida-json", type=Path, help="escribir aquí el resumen de la corrida")
    ap.add_argument("--log-texto", action="store_true", help="log legible en vez de JSON")
    return ap


def valores_config(args: argparse.Namespace) -> dict:
    """Parámetros de la corrida: los del libro (``--parametros-desde-libro``) y encima los explícitos.

    Si se da un período explícito, los parámetros de eventos (L2/L4) y el fin de Daily lo siguen
    salvo que también se den explícitamente (D-07: L2/L4 = C5/C7; Checkpoint C-1). Así nunca se
    mezcla el período pedido con los L2/L4/Daily!C que traía el libro.
    """
    valores: dict = {}
    if args.parametros_desde_libro:
        valores.update(parametros_desde_celdas(leer_celdas_parametros(args.libro)).valores)
    if args.periodo_inicio:
        valores["inicio_periodo"] = valores["inicio_periodo_eventos"] = args.periodo_inicio
    if args.periodo_fin:
        valores["fin_periodo"] = valores["fin_periodo_eventos"] = valores["fin_diario"] = args.periodo_fin
    explicitos = {
        "inicio_periodo_eventos": args.eventos_inicio,
        "fin_periodo_eventos": args.eventos_fin,
        "fin_diario": args.diario_fin,
        "solo_tiempo_operacional": args.solo_operacional,
        "aplicar_evento_excusable": args.excusable,
        "aplicar_evento_excusable_eventos": args.eventos_excusable,
        "modo_huecos": args.modo_huecos,
    }
    valores.update({k: v for k, v in explicitos.items() if v is not None})
    return valores


def main(argv: list[str] | None = None) -> int:
    args = construir_parser().parse_args(argv)
    configurar_logging(json_=not args.log_texto)
    if args.entorno:
        os.environ["ETL_ARENA_ENTORNO"] = args.entorno
    cfg = construir_config(
        **valores_config(args), tipo_corrida=args.tipo, es_oficial=args.oficial, archivo_origen=args.libro.name
    )

    referencia = None
    if args.referencia == "libro":
        referencia = extraer_referencia(args.libro)
    elif args.referencia != "none":
        referencia = extraer_referencia(Path(args.referencia))

    repo = None
    if not args.sin_bd:
        from etl_arena.persistence import RepositorioCorridas, crear_engine

        repo = RepositorioCorridas(crear_engine())
    try:
        resultado = ejecutar_corrida(
            args.libro,
            cfg,
            repo,
            hash_archivo=sha256_archivo(args.libro),
            estado_exclusiones=args.estado_exclusiones,
            referencia=referencia,
        )
    except Exception as exc:  # ya quedó registrado (log + etl_run failed)
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return SALIDA["failed"]

    d = resultado.resultado.disponibilidad
    resumen = {
        "num_corrida": resultado.num_corrida,
        "id_corrida": resultado.id_corrida,
        "estado": resultado.estado,
        "periodo": [str(cfg.inicio_periodo), str(cfg.fin_periodo)],
        "C12": d.bloques_muestreo,
        "C14": d.bloques_racks_indisponibles,
        "C16": d.disponibilidad_periodo,
        "C19": d.disponibilidad_anual_acumulada,
        "L10": resultado.resultado.eventos.horas_rack_totales,
        "reconciliacion": resultado.resumen_calidad.get("reconciliacion"),
        "segundos": resultado.segundos,
        "filas_guardadas": sum(resultado.filas_guardadas.values()),
    }
    texto = json.dumps(resumen, ensure_ascii=False, indent=2, default=str)
    print(texto)
    if args.salida_json:
        args.salida_json.write_text(texto, encoding="utf-8")
    return SALIDA.get(resultado.estado, 1)


if __name__ == "__main__":
    raise SystemExit(main())

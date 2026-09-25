"""Crea o actualiza las tablas de una base (esquema idempotente de ``sql/``), o vacía TEST/QA.

En Azure SQL (``*.database.windows.net``) la base debe existir (se crea en el portal con la oferta
gratuita): el script solo aplica el esquema. Ambientes: PROD = ``trina_etl`` (por defecto) y
TEST/QA = ``trina_etl_prueba`` (``--entorno prueba``).

Uso:
    python scripts/crear_base.py                         # PROD: crea/actualiza tablas (no borra nada)
    python scripts/crear_base.py --entorno prueba        # TEST/QA: crea/actualiza tablas
    python scripts/crear_base.py --entorno prueba --vaciar --confirmar trina_etl_prueba
                                                         # TEST/QA: borra TODO y la deja como nueva
    python scripts/crear_base.py --base X                # otra base en la misma instancia (desarrollo)

``--vaciar`` borra vistas y tablas y las vuelve a crear: los identificadores (IDENTITY) parten de 1, se
vuelven a cargar los datos maestros y se conservan roles y usuarios. Solo existe para TEST/QA y exige
escribir su nombre en ``--confirmar``; con PROD se niega siempre.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))

from etl_arena.ambientes import AYUDA_ENTORNO, OPCIONES_ENTORNO  # noqa: E402
from etl_arena.persistence import (  # noqa: E402
    aplicar_esquema,
    crear_base_si_no_existe,
    crear_engine,
    url_configurada,
    vaciar_ambiente_prueba,
)
from etl_arena.persistence.conexion import descripcion_ambiente, es_azure  # noqa: E402


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--base", help="nombre de la base (por defecto, la del ambiente)")
    ap.add_argument("--entorno", choices=OPCIONES_ENTORNO, help=AYUDA_ENTORNO)
    ap.add_argument("--vaciar", action="store_true", help="solo TEST/QA: borrar todo y dejar la base como nueva")
    ap.add_argument("--confirmar", default="", help="con --vaciar: escribir exactamente trina_etl_prueba")
    args = ap.parse_args()
    if args.entorno:
        os.environ["ETL_ARENA_ENTORNO"] = args.entorno
    url = url_configurada(args.base)
    print(f"Ambiente: {descripcion_ambiente()}")
    if not url.database:
        raise SystemExit("la URL no indica base de datos; usar --base")
    if not es_azure(url) and not args.vaciar:
        crear_base_si_no_existe(url.database, url)
    engine = crear_engine(url)
    try:
        if args.vaciar:
            print(f"Vaciando {url.host}/{url.database}: se borran todas las corridas y se recrean las tablas…")
            try:
                pasos = vaciar_ambiente_prueba(engine, args.confirmar)
            except RuntimeError as exc:
                raise SystemExit(f"ERROR: {exc}") from exc
        else:
            pasos = aplicar_esquema(engine)
        for script, n in pasos.items():
            print(f"  {script}: {n} lotes")
    finally:
        engine.dispose()
    print(f"{'Base vaciada y recreada' if args.vaciar else 'Esquema aplicado'} en {url.host}/{url.database}")


if __name__ == "__main__":
    main()

"""Crea la base de ``ETL_ARENA_DB_URL`` (si no existe) y aplica el esquema idempotente de ``sql/``.

En Azure SQL (``*.database.windows.net``) la base debe existir (se crea en el portal con la oferta
gratuita): el script solo aplica el esquema.

Uso:
    python scripts/crear_base.py            # base de la URL de .env (p. ej. ETL_Arena)
    python scripts/crear_base.py --base X   # otra base en la misma instancia
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))

from etl_arena.persistence import aplicar_esquema, crear_base_si_no_existe, crear_engine, url_configurada  # noqa: E402
from etl_arena.persistence.conexion import es_azure  # noqa: E402


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--base", help="nombre de la base (por defecto, el de ETL_ARENA_DB_URL)")
    args = ap.parse_args()
    url = url_configurada(args.base)
    if not url.database:
        raise SystemExit("la URL no indica base de datos; usar --base")
    if not es_azure(url):
        crear_base_si_no_existe(url.database, url)
    engine = crear_engine(url)
    try:
        for script, n in aplicar_esquema(engine).items():
            print(f"  {script}: {n} lotes")
    finally:
        engine.dispose()
    print(f"Esquema aplicado en {url.host}/{url.database}")


if __name__ == "__main__":
    main()

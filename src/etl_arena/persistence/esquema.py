"""Aplicación idempotente de los scripts de ``sql/`` (tarea 2.1)."""

from __future__ import annotations

import re
from pathlib import Path

from sqlalchemy.engine import Engine

DIR_SQL = Path(__file__).resolve().parents[3] / "sql"

# Orden de aplicación dentro de la base (00_database.sql se usa aparte con sqlcmd; 07 son consultas).
ORDEN_SCRIPTS: tuple[str, ...] = (
    "01_maestros.sql",
    "02_corrida.sql",
    "03_indices.sql",
    "04_vistas.sql",
    "05_seed_proyecto.sql",
    "05_seed_tipo_detencion.sql",
    "06_seed_annual_manual.sql",
    "06_roles.sql",
)

_GO = re.compile(r"^\s*GO\s*;?\s*$", re.IGNORECASE | re.MULTILINE)


def lotes(sql: str) -> list[str]:
    """Divide un script en lotes por ``GO`` (como sqlcmd); omite directivas ``:setvar`` y lotes vacíos."""
    sin_directivas = "\n".join(linea for linea in sql.splitlines() if not linea.lstrip().startswith(":"))
    return [lote.strip() for lote in _GO.split(sin_directivas) if lote.strip()]


def aplicar_script(engine: Engine, ruta: Path) -> int:
    """Ejecuta un script lote a lote en autocommit. Devuelve la cantidad de lotes."""
    partes = lotes(ruta.read_text(encoding="utf-8"))
    conn = engine.raw_connection()
    try:
        # raw_connection() es un proxy del pool: el autocommit se fija en la conexión pyodbc real
        conn.driver_connection.autocommit = True
        cur = conn.cursor()
        for lote in partes:
            cur.execute(lote)
            while cur.nextset():  # consumir resultados (p. ej. MERGE / SELECT)
                pass
        cur.execute("SET NOCOUNT OFF")  # no devolver al pool opciones de sesión de los scripts
        cur.close()
    finally:
        conn.driver_connection.autocommit = False  # no devolver al pool una conexión en autocommit
        conn.close()
    return len(partes)


def aplicar_esquema(engine: Engine, directorio: Path = DIR_SQL) -> dict[str, int]:
    """Aplica ``ORDEN_SCRIPTS``. Todos son idempotentes: se puede ejecutar cualquier número de veces."""
    return {nombre: aplicar_script(engine, directorio / nombre) for nombre in ORDEN_SCRIPTS}

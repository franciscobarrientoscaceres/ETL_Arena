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


_REINICIO = """
DECLARE @sql NVARCHAR(MAX) = N'';
SELECT @sql += N'ALTER TABLE ' + QUOTENAME(SCHEMA_NAME(t.schema_id)) + N'.' + QUOTENAME(t.name)
             + N' DROP CONSTRAINT ' + QUOTENAME(f.name) + N';' + NCHAR(10)
FROM sys.foreign_keys AS f JOIN sys.tables AS t ON t.object_id = f.parent_object_id;
SELECT @sql += N'DROP VIEW ' + QUOTENAME(SCHEMA_NAME(schema_id)) + N'.' + QUOTENAME(name) + N';' + NCHAR(10)
FROM sys.views WHERE is_ms_shipped = 0;
SELECT @sql += N'DROP TABLE ' + QUOTENAME(SCHEMA_NAME(schema_id)) + N'.' + QUOTENAME(name) + N';' + NCHAR(10)
FROM sys.tables WHERE is_ms_shipped = 0;
EXEC sys.sp_executesql @sql;
"""


def reiniciar_esquema(engine: Engine, directorio: Path = DIR_SQL) -> dict[str, int]:
    """Solo pruebas: elimina vistas y tablas y vuelve a aplicar el esquema. Se niega a operar sobre
    una base cuyo nombre no contenga "test" (p. ej. la base fija de pruebas en Azure)."""
    base = (engine.url.database or "").lower()
    if "test" not in base:
        raise RuntimeError(f"reiniciar_esquema solo opera sobre bases de prueba, no sobre {engine.url.database!r}")
    conn = engine.raw_connection()
    try:
        conn.driver_connection.autocommit = True
        cur = conn.cursor()
        cur.execute(_REINICIO)
        cur.close()
    finally:
        conn.driver_connection.autocommit = False
        conn.close()
    return aplicar_esquema(engine, directorio)

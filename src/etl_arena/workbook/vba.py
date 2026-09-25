"""¿Las macros del libro aplican la ``Exclusion_Matrix``? (F-37, F-44, D-19; lectura, nunca edición).

Las macros oficiales de septiembre, con ``C31``/``L14 = "Yes"``, excusan con ``PlantActivity!D``,
regla que **no** se usa (las exclusiones salen solo de ``Exclusion_Matrix`` hasta que Alex confirme
otra cosa). Solo el maestro v1.1 (tarea 4.0, ``design.md §Libro maestro v1.1``) lee la matriz en
``cmdCalcAvailability`` **y** en ``mcoCreateList``. Con otro libro, ``workbook.macros`` escribe
``"No"`` en C31/L14 y el orquestador omite la referencia Excel si el período tiene exclusiones.
"""

from __future__ import annotations

import re
from pathlib import Path

HOJA = "Exclusion_Matrix"
MACROS_CON_MATRIZ = ("cmdCalcAvailability", "mcoCreateList")


def _procedimiento(codigo: str, nombre: str) -> str:
    m = re.search(rf"^\s*(?:Public\s+|Private\s+)?Sub\s+{nombre}\b.*?^\s*End\s+Sub", codigo, re.S | re.M | re.I)
    return m.group(0) if m else ""


def codigo_vba(ruta: str | Path) -> str:
    from oletools.olevba import VBA_Parser

    parser = VBA_Parser(str(ruta))
    try:
        return "\n".join(codigo for _, _, _, codigo in parser.extract_macros())
    finally:
        parser.close()


def macros_aplican_matriz(ruta: str | Path) -> bool:
    """``True`` solo si ``cmdCalcAvailability`` y ``mcoCreateList`` leen la hoja ``Exclusion_Matrix``."""
    codigo = codigo_vba(ruta)
    return all(HOJA in _procedimiento(codigo, nombre) for nombre in MACROS_CON_MATRIZ)

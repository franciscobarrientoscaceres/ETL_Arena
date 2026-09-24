"""Catálogos de fallas del libro.

* ``PCS-Fault`` (fila 2 encabezado, B:G): fuente del seed ``tipo_detencion`` (167 códigos, F-19).
* ``ListOfFaults!N6:O172``: códigos que usa el resumen ``N:Q`` (``SUMIF``). En el libro de
  septiembre son 163, no los 167 del catálogo; para paridad del resumen manda esta lista.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from etl_arena.ingestion.xlsx_stream import LibroXlsx


@dataclass(frozen=True)
class CodigoFalla:
    numero: float | None  # B  Code Number
    codigo: str  # C  Code Fault ("F55")
    descripcion_pe: str  # D  Description PE
    codigo_descripcion: str  # E  Code + Description
    significado: str | None  # F  Meaning
    operativo: str | None  # G  Operative


def leer_catalogo_fallas(ruta: str | Path) -> list[CodigoFalla]:
    salida = []
    with LibroXlsx(ruta) as libro:
        for _, c in libro.filas("PCS-Fault", min_fila=3, max_col=7):
            if c.get(3) is None:
                continue
            salida.append(CodigoFalla(c.get(2), str(c[3]), str(c.get(4, "")), str(c.get(5, "")), c.get(6), c.get(7)))
    return salida


def leer_codigos_resumen(ruta: str | Path) -> list[tuple[str, str]]:
    """``[(N, O)]`` de ``ListOfFaults!N6:O172`` en el orden actual de la hoja."""
    with LibroXlsx(ruta) as libro:
        return [
            (str(c[14]), str(c.get(15, "")))
            for _, c in libro.filas("ListOfFaults", min_fila=6, max_fila=172, max_col=15)
            if c.get(14) is not None
        ]

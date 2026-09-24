"""Tabla resumen ``ListOfFaults!N:Q`` por código de falla (F-18).

``P = SUMIF($F$6:$F$50000, N, $I$6:$I$50000)``, ``Q = P / $L$10``; ``mcoOrder`` ordena
``N5:Q172`` por ``P`` descendente. ``SUMIF`` compara texto sin distinguir mayúsculas. El sort
de Excel es estable, así que los empates quedan en el orden previo de la hoja (desconocido
para Python): aquí se usa el orden del catálogo y la reconciliación compara como mapa.
"""

from __future__ import annotations

from dataclasses import dataclass

from etl_arena.model import RegistroLista

FILAS_SUMIF = 50000 - 6 + 1  # F6:F50000


@dataclass(frozen=True)
class FilaResumenCodigo:
    codigo: str  # N
    descripcion_pe: str  # O
    horas_rack: float  # P
    porcentaje: float | None  # Q (None = #DIV/0!)


def resumen_por_codigo(
    cerrados: list[RegistroLista], catalogo: list[tuple[str, str]], l10: float
) -> list[FilaResumenCodigo]:
    sumas: dict[str, float] = {}
    for r in cerrados[:FILAS_SUMIF]:
        clave = (r.codigo_falla or "").casefold()
        sumas[clave] = sumas.get(clave, 0.0) + r.horas_rack  # SUMIF recorre F6..F50000 en orden
    filas = []
    for codigo, descripcion in catalogo:
        p = sumas.get(codigo.casefold(), 0.0)
        filas.append(FilaResumenCodigo(codigo, descripcion, p, p / l10 if l10 else None))
    return sorted(filas, key=lambda f: -f.horas_rack)  # sorted es estable

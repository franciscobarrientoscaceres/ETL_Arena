"""Genera sql/05_seed_tipo_detencion.sql desde la hoja PCS-Fault del libro (R12.3, F-19, F-35, D-14).

Uso:
    python scripts/generar_seed_tipo_detencion.py [--libro <xlsm>] [--salida sql/05_seed_tipo_detencion.sql]

PCS-Fault trae 167 filas y 163 códigos distintos: F228, F230, F231 y F232 aparecen dos veces,
idénticos salvo "Meaning" (Crítico / Parcial). Mientras D-14 siga abierta se conserva la fila
"Crítico" (default conservador) y el script lo deja anotado en el SQL generado.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))

from etl_arena.ingestion.catalogos import CodigoFalla, leer_catalogo_fallas  # noqa: E402

LIBRO_POR_DEFECTO = RAIZ / "data" / "AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm"
SALIDA_POR_DEFECTO = RAIZ / "sql" / "05_seed_tipo_detencion.sql"
SIGNIFICADO_PREFERIDO = "Crítico"  # D-14 (default mientras negocio no defina)


def deduplicar(catalogo: list[CodigoFalla]) -> tuple[list[CodigoFalla], list[str]]:
    """Una fila por código; ante duplicados gana ``SIGNIFICADO_PREFERIDO``. Devuelve (filas, notas)."""
    por_codigo: dict[str, list[CodigoFalla]] = {}
    for c in catalogo:
        por_codigo.setdefault(c.codigo, []).append(c)
    filas, notas = [], []
    for codigo, variantes in por_codigo.items():
        elegida = next((v for v in variantes if v.significado == SIGNIFICADO_PREFERIDO), variantes[0])
        if len(variantes) > 1:
            otros = ", ".join(repr(v.significado) for v in variantes if v is not elegida)
            notas.append(
                f"{codigo}: {len(variantes)} filas; se usa Significado={elegida.significado!r} "
                f"(descartado {otros}) — D-14"
            )
        filas.append(elegida)
    return filas, notas


def _texto(v: object) -> str:
    return "NULL" if v is None else "N'" + str(v).replace("'", "''") + "'"


def generar_sql(filas: list[CodigoFalla], notas: list[str], origen: str) -> str:
    valores = ",\n".join(
        f"    ({int(c.numero)}, {_texto(c.codigo)}, {_texto(c.descripcion_pe)}, {_texto(c.codigo_descripcion)}, "
        f"{_texto(c.significado)}, {_texto(c.operativo)})"
        for c in filas
    )
    cabecera = "\n".join(f"--   {n}" for n in notas) or "--   (sin duplicados)"
    return f"""-- 05_seed_tipo_detencion.sql — GENERADO por scripts/generar_seed_tipo_detencion.py; no editar a mano.
-- Origen: hoja PCS-Fault de {origen}
-- {len(filas)} códigos distintos. Duplicados resueltos (F-35, D-14):
{cabecera}
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

MERGE dbo.tipo_detencion AS t
USING (VALUES
{valores}
) AS s (IdTipoDetencion, CodigoFalla, DescripcionFallaPE, CodigoDescripcion, Significado, Operativo)
ON t.IdTipoDetencion = s.IdTipoDetencion
WHEN MATCHED THEN UPDATE SET
    CodigoFalla = s.CodigoFalla, DescripcionFallaPE = s.DescripcionFallaPE, CodigoDescripcion = s.CodigoDescripcion,
    Significado = s.Significado, Operativo = s.Operativo
WHEN NOT MATCHED THEN
    INSERT (IdTipoDetencion, CodigoFalla, DescripcionFallaPE, CodigoDescripcion, Significado, Operativo)
    VALUES (s.IdTipoDetencion, s.CodigoFalla, s.DescripcionFallaPE, s.CodigoDescripcion, s.Significado, s.Operativo);
GO
"""


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--libro", type=Path, default=LIBRO_POR_DEFECTO)
    ap.add_argument("--salida", type=Path, default=SALIDA_POR_DEFECTO)
    args = ap.parse_args()
    filas, notas = deduplicar(leer_catalogo_fallas(args.libro))
    args.salida.write_text(generar_sql(filas, notas, args.libro.name), encoding="utf-8")
    print(f"{args.salida}: {len(filas)} códigos; {len(notas)} duplicados resueltos")
    for n in notas:
        print("  ", n)


if __name__ == "__main__":
    main()

# ADR-02 — Los motores consumen una matriz por fila origen (`MatrizPCS`)

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R4, R5, R17

## Contexto
El VBA recorre `RawData-PCS` por fila y consulta la fila **anterior** y **siguiente** de la hoja completa (fallback de descripción, cierre de evento). Un DataFrame largo (`timestamp × pcs`) filtrado al período pierde esa vecindad y puede reordenar o deduplicar timestamps (F-04, F-07).

## Opciones
1. DataFrame largo y `groupby` por PCS.
2. Matriz `n_filas × n_pcs` (arrays numpy) indexada por fila origen, más vectores de seriales y flags por fila.

## Decisión
Opción 2: `MatrizPCS` en `etl_arena.model`. El formato largo solo se genera para staging/persistencia (`a_formato_largo`).

## Consecuencias
- Emulación directa de `Cells(dblRec±1, col)`, incluida la fila de encabezado (D-08).
- ≈1M celdas por libro: cabe en memoria y los motores son O(filas × PCS).
- Los motores no importan pandas.

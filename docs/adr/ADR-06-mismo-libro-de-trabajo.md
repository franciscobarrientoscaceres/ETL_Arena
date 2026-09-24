# ADR-06 — Python lee el mismo libro de trabajo que ejecutan las macros

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R2, R14

## Contexto
La reconciliación Python vs Excel solo es válida si ambos procesan exactamente el mismo input (RawData, PlantActivity y parámetros).

## Opciones
1. Python lee el export SCADA directamente; Excel lee su copia.
2. Python lee `data/work/<corte>/libro.xlsm` después de `prepare-workbook` y de las macros.

## Decisión
Opción 2. El maestro en `data/` nunca se modifica; `data/processed/` es inmutable con sha256.

## Consecuencias
- Si el adapter introduce un error, ambos lados lo comparten; la validación del contrato SCADA (tarea 4.6) y la alineación con PlantActivity lo cubren.
- El sha256 del libro de trabajo se registra en `etl_run`.

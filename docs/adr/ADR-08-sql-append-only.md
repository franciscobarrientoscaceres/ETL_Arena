# ADR-08 — SQL append-only por `IdCorrida`; revisión de detenciones en tabla aparte

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R11

## Contexto
`mcoCleanTable` destruye los resultados anteriores. Cada lunes se recalcula el período y, si `detencion` se reescribiera, se duplicaría y se perdería `Observacion`/`EstadoRevision` (F-27).

## Opciones
1. Upsert/replace por período.
2. Append-only por `IdCorrida` + vistas `v_*_vigente` para la última corrida oficial; workflow humano en `detencion_revision` con clave de negocio `IdProyecto + NumeroPCS + FechaInicio`.

## Decisión
Opción 2.

## Consecuencias
- Sin `DELETE`/`UPDATE` salvo el estado de `etl_run` e inserciones en `detencion_revision`.
- Power BI consume solo las vistas `v_*_vigente`.
- Crecimiento de volumen: ver ADR-09.

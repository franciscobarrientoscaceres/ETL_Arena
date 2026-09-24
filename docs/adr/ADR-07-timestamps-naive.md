# ADR-07 — Timestamps naive en hora local de Chile

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R4, R11

## Contexto
Los datos llegan en hora local naive con saltos DST: 2026-09-06 00:00 → 01:00 (faltan filas; C12 = 1975) y en abril habrá timestamps locales repetidos. `pytz.localize(is_dst=None)` falla en el retroceso (F-32). El VBA cuenta filas, no instantes (F-07).

## Opciones
1. Localizar a `America/Santiago` y convertir a UTC.
2. Guardar naive local + serial, sin localizar en los motores.

## Decisión
Opción 2. Cualquier conversión a UTC es de presentación y debe documentarse.

## Consecuencias
- Los timestamps repetidos del DST se cuentan dos veces, igual que Excel; se reportan como anomalía.
- `proyecto.ZonaHoraria` es informativa.

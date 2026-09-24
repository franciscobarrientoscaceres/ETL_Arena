# ADR-03 — Tiempo como serial Excel (`float64`) leído del XML crudo

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R4, R17

## Contexto
El VBA filtra y resta fechas como `double`. Ejemplo: `ListOfFaults!E6 = 5.5000000001164153`; con `timedelta` Python obtiene 5,5 exacto y la diferencia en rack-hours (4,6e-9) supera la tolerancia de 1e-9 (F-15).

## Opciones
1. Convertir a `datetime` al leer y operar con `timedelta`.
2. Conservar `SerialFechaExcelOrigen` (`float64`) desde el `<v>` crudo de la celda y hacer filtros y restas sobre él; `datetime` solo para presentación.

## Decisión
Opción 2.

## Consecuencias
- La lectura del libro es por XML streaming (no `openpyxl` con conversión de fechas).
- SQL guarda el serial y el timestamp local (`SerialFechaExcelOrigen`, `MarcaTiempoLocalOrigen`).
- Las conversiones serial↔datetime viven en `excel_semantics.fechas`.

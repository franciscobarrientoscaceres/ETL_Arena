# ADR-05 — Escritura del libro por COM; lectura por XML streaming

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R14, R19

## Contexto
`openpyxl` con `keep_vba=True` conserva el VBA pero descarta dibujos, controles y gráficos (hoja `Graph`) (F-22). Leer ~1M celdas por COM es lento.

## Opciones
1. `openpyxl` para leer y escribir.
2. COM (`pywin32`) para leer y escribir.
3. COM para escribir y ejecutar macros; lectura por XML streaming del `.xlsm`.

## Decisión
Opción 3.

## Consecuencias
- `workbook/` solo corre en Windows con Excel instalado (marker `excel`), con Excel visible y la carpeta de trabajo como *Trusted Location* (las macros usan `.Select`).
- Tras escribir por COM se guarda el libro antes de que el pipeline lo lea por XML.

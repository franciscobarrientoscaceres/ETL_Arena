# ADR-04 — Semántica Excel/VBA centralizada en `etl_arena.excel_semantics`

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R17

## Contexto
El VBA depende de reglas implícitas: celda vacía = 0, texto > número en comparaciones, `"F" & n` para fallas numéricas, `FIND`/`MID` con espacio inicial, flags asimétricos (`C21 <> "No"` vs `C31 = "Yes"`) (F-12, F-14, F-16, F-21).

## Opciones
1. Que cada motor implemente las reglas que necesite.
2. Un módulo único con funciones puras y tests propios.

## Decisión
Opción 2 (R17): `es_vacio`, `igual_numero`, `texto_excel`, `codigo_falla_excel`, `flag_si`/`flag_no`, seriales, `redondear_excel`.

## Consecuencias
- Revisión reforzada de code-reviewer (tarea 1.2).
- Checklist de paridad: "toda comparación de celda pasa por `excel_semantics`".

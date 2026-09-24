# ADR-01 — Paquete `etl_arena` con layout *src*

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R1

## Contexto
El plan original (AGENTS §18) usaba paquetes top-level en `src/` (`config`, `staging`, `ingestion`…). Nombres como `config` colisionan con otros paquetes del entorno y hacen que los tests importen el código del working tree en vez del instalado.

## Opciones
1. Paquetes top-level en `src/<módulo>/`.
2. Un único paquete `etl_arena` en `src/etl_arena/<módulo>/`, instalado en modo editable.

## Decisión
Opción 2. `pyproject.toml` declara `where = ["src"]`; se instala con `pip install -e .[dev]`.

## Consecuencias
- Imports explícitos: `from etl_arena.availability import motor`.
- Los tests corren contra el paquete instalado (detecta archivos faltantes en el empaquetado).
- README, AGENTS y CLAUDE describen `src/etl_arena/` (sincronizado en la tarea 5.2).

# ADR-09 — *Clustered columnstore* en tablas de muestras

**Estado:** Ajustada por [ADR-12](./ADR-12-estado-vigente.md) (aceptada 2026-09-26: tablas de estado en rowstore con clave por tiempo) · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R11

## Contexto
El plan original suponía un export acumulado (01-01-2026 → domingo): ~16k filas × 61 PCS ≈ 1M filas por tabla por corrida, creciendo cada semana (F-28). Con D-07 (2026-09-24) el export es **incremental** (~672 filas × 61 ≈ 41k por semana) y cada corrida persiste solo las filas de su período (mes en curso: ≤ ~3.000 × 61 ≈ 180k filas por tabla). Con ~5 corridas por mes el volumen queda en el orden de 10M filas/año.

## Opciones
1. Rowstore con índice clustered por `(IdCorrida, MarcaTiempoMuestra, NumeroPCS)`.
2. *Clustered columnstore* en `raw_pcs_sample`, `plant_activity_sample` y `availability_sample_result`.
3. Persistir solo el período calculado.

## Decisión
Opción 2 (sigue siendo conveniente para agregaciones de Power BI). Las tablas de muestras guardan las filas del período calculado, con `EsFilaNuevaDelExport` para distinguir lo que aportó el export de la semana (D-09 ajustada).

## Consecuencias
- Carga con `fast_executemany` o bulk insert (requiere ODBC Driver 18).
- La compresión columnstore mantiene el tamaño manejable; la tarea 2.6 lo mide.

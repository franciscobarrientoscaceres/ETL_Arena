# ADR-11 — `Exclusion_Matrix` como fuente de los eventos de exclusión

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** definición de negocio 2026-09-24, `audit.md` F-37…F-39 · **Requisitos:** R6.5–R6.8, R7.5, R8.3

## Contexto
Hasta septiembre 2026 la exclusión ("Excusable event?", C31/L14) ponderaba con `PlantActivity!D` por fila. El negocio definió una hoja nueva, `Exclusion_Matrix`, con un valor por fila **y por PCS**: `0`/vacío sin evento, `1` = se consideran todos los módulos, `2` = se consideran los módulos previos al evento. PlantActivity queda solo para "Only Operational Time?" (C21). El libro de agosto trae la primera matriz y una macro que la aplica, pero con otras celdas y otra `Daily`; las definiciones oficiales siguen siendo las de septiembre.

## Opciones
1. Seguir con `PlantActivity!D` y tratar la matriz como un factor adicional.
2. Reemplazar `PlantActivity!D` por la matriz, con el valor 2 emulando la macro de agosto (`F(r) = F(r−1)`).
3. Reemplazar `PlantActivity!D` por la matriz, con el valor 2 según la definición de negocio (baterías previas al tramo) y reportar dónde la macro de agosto diferiría.

## Decisión
Opción 3. La matriz se une por fila, como el resto del libro. Con el flag en "Yes" y `M < 4`: `0` → `C3 − M`; `1` → `0`; `2` → baterías previas. Aplica al KPI (C14, tabla, Daily) y a los eventos (D-16). `VersionAlgoritmo = availability-v1.1-exclusion-matrix`.

## Consecuencias
- Contra el libro de agosto, C14 y la tabla E4:BO coinciden bit a bit; las opciones 2 y 3 dan lo mismo en todo el libro.
- Sin la hoja (libros hasta septiembre) los resultados no cambian.
- SQL: `availability_sample_result.ValorExclusion` reemplaza a `FactorExcusable`; nueva `exclusion_matrix_sample` con valor, baterías previas y comentario.
- Operación: la carga mensual pasa a ser la de la matriz (`load-exclusion-matrix`); mientras falte, el KPI semanal es preliminar (D-17).

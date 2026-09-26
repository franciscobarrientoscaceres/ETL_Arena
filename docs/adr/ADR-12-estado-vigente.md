# ADR-12 — Estado vigente por período (recarga por mes) en lugar de copias por `IdCorrida`

**Estado:** Aceptada (F. Barrientos, 2026-09-26) · **Fecha:** 2026-09-26 · **Reemplaza a:** ADR-08 (y ajusta ADR-09) · **Requisitos:** R10, R11, R12, R15, R18, R19 (rev. 3) · **Plan:** `docs/plan-estado-vigente.md` · **Decisiones:** D-20 a D-27

## Contexto
Con ADR-08 cada ejecución inserta una copia completa de su período (muestras, resultados, detenciones, KPI) bajo un
`IdCorrida` nuevo, y las vistas `v_*_vigente` eligen la última corrida oficial por mes. El negocio (F. Barrientos,
2026-09-26) pide lo contrario: que recargar un período **reemplace** los datos de ese período, con una sola tabla
simple por tema, para que Power BI (Misael) lea las tablas sin conocer las corridas. Con recargas semanales del mes en
curso, ADR-08 guarda 4–5 copias de cada semana, cada detención y cada KPI mensual.

Restricciones que se mantienen: paridad bit a bit con el Excel (el cálculo no cambia), trazabilidad de cualquier valor
hasta su archivo de origen (R15), registro de cada dato de SCADA corregido (D-13), revisiones humanas de detenciones
que sobreviven a los recálculos (F-27), y KPI oficial "Sin/Con Exclusiones" (D-17).

## Opciones
1. **Mantener ADR-08** (copias por corrida + vistas vigentes). Cumple todo, pero las tablas acumulan copias.
2. **Estado vigente con reemplazo por rango arbitrario** (desde–hasta cualquiera). Simple, pero rompe la paridad de
   detenciones: el VBA corta los eventos en el borde del período, así que recargar desde un día intermedio parte los
   eventos que cruzan ese día.
3. **Estado vigente con reemplazo por mes calendario** + registro mínimo (ejecuciones y datos corregidos).

## Decisión
Opción 3.

- **Unidad de reemplazo: el mes** (D-20). El inicio se ajusta al día 1; un rango de varios meses se procesa mes a mes.
  Cada carga reemplaza todo lo del mes: muestras, detenciones, día a día, KPI mensual y calidad.
- **Una transacción por mes**, con bloqueo por proyecto (`sp_getapplock`): comparar con lo vigente y registrar en
  `correccion_dato` cada dato crudo que cambió; borrar el rango; insertar lo nuevo; detenciones por clave de negocio
  `(IdProyecto, NumeroPCS, FechaInicio, Ocurrencia)` con `MERGE` (mismo `IdDetencion` si la detención sigue
  existiendo); `disponibilidad_mensual` por `(IdProyecto, Anio, Mes)`.
- **Solo publica una carga `success`** (o `sin_referencia`). `parity_failed` queda en `etl_run` con sus diferencias y
  no toca el estado, salvo `--publicar-aunque-no-cuadre` (D-22). Las corridas `golden` y `--sin-bd` nunca publican.
- **Protecciones** (D-21, D-23): recortar un mes ya cargado, reemplazar un mes `excel_manual` o bajar un mes de
  "Con Exclusiones" a "Sin Exclusiones" requieren un flag explícito.
- **Registro mínimo** (D-24): `etl_run` (una fila por ejecución), `correccion_dato`, `exclusion_matrix_carga`,
  `detencion_revision`, `excel_reference_run` (solo KPI) y `reconciliation_result` (resumen por nivel y diferencias).
  Cada fila vigente guarda `NumCorrida` (la carga que la escribió).
- **Esquema v2** (D-26): 6 tablas de estado (`muestra_pcs`, `muestra_planta`, `detencion`, `disponibilidad_diaria`,
  `disponibilidad_mensual`, `calidad_dato`) y vistas `v_disponibilidad_anual`, `v_resumen_codigo_mensual`,
  `v_detencion`. Claves de tiempo con `Ocurrencia` para la hora repetida del cambio de hora de abril (F-32).

## Consecuencias
- `etl_writer` recibe `DELETE`/`UPDATE` **solo** en las tablas de estado; las de registro siguen solo-inserción.
  `reiniciar_esquema` y `vaciar_ambiente_prueba` no cambian.
- Se pierde la historia de valores reemplazados, salvo lo registrado en `correccion_dato` y `etl_run` (aceptado por el negocio).
- ADR-09: las tablas de estado pasan a *rowstore* con clave agrupada por tiempo (los borrados por rango y el volumen,
  ~2,2 M filas/año en `muestra_pcs`, no justifican columnstore).
- La etapa de reproceso (4.13) se absorbe: corregir = recargar el mes con el export corregido (4.7 reemplaza las filas
  del libro de trabajo y registra las diferencias).
- `--oficial` deja de tener efecto: en PROD toda carga `success` publica; práctica en TEST/QA o `--sin-bd`.
- Migración: PROD está vacía (se crea v2 directo); TEST/QA se vacía y se recarga.

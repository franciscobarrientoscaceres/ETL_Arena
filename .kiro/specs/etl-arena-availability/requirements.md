# Requirements Document — ETL Arena Availability

## Introduction

Este proyecto reemplaza progresivamente el cálculo de disponibilidad del activo **Arena BESS** (almacenamiento de energía), actualmente implementado en un libro Excel con macros VBA, por un proceso reproducible y auditable en **Python + SQL Server**.

El objetivo de la primera fase es obtener **paridad exacta** con el Excel: reproducir los mismos valores de `C12`, `C14`, `C16` y `C19`, la misma lista de eventos de `ListOfFaults`, y los mismos KPI diarios y acumulados anuales, usando los mismos datos de entrada.

**Parámetros del activo (Arena BESS):**

| Parámetro | Valor |
|---|---|
| Fecha de inicio de operación | 08/Abril/2026 |
| Total PCS | 61 |
| Módulos BEC por PCS | 4 |
| Racks por BAC | 12 |
| Total racks | 2.928 |
| Frecuencia de muestreo | 15 minutos |
| Versión del algoritmo | `availability-v1-excel-parity` |

---

## Glossary

- **PCS**: Power Conversion System. Unidad de conversión de potencia. El activo tiene 61 unidades.
- **BEC**: Battery Energy Controller. Módulo de batería. Cada PCS tiene 4 módulos BEC.
- **Rack**: Unidad mínima de batería. Cada BAC contiene 12 racks. Total: 2.928 racks.
- **ETL_Engine**: El sistema Python que reproduce la lógica VBA del Excel.
- **Ingestion_Module**: Componente responsable de leer los archivos origen (Excel o exportaciones CSV).
- **Staging_Module**: Componente que almacena los datos crudos sin transformar en SQL Server.
- **Normalization_Module**: Componente que convierte el formato ancho (244 columnas) a formato largo (1 fila por PCS×timestamp).
- **Enrichment_Module**: Componente que une los datos PCS con los factores de PlantActivity.
- **Availability_Engine**: Componente que reproduce `cmdCalcAvailability` — calcula C12, C14 y el KPI del período.
- **Fault_Events_Engine**: Componente que reproduce `mcoCreateList` — consolida intervalos consecutivos en eventos de falla discretos.
- **Aggregation_Module**: Componente que calcula KPI diario, mensual y acumulado anual.
- **Persistence_Module**: Componente que escribe resultados en SQL Server de forma append-only.
- **Reconciliation_Module**: Componente que compara una corrida Python contra una corrida Excel en 5 niveles.
- **run_id**: Identificador único UUID de una ejecución del ETL_Engine. Clave de auditoría en todas las tablas.
- **algorithm_version**: Cadena que identifica la versión del algoritmo usada en una corrida (ej. `availability-v1-excel-parity`).
- **C12**: Contador de bloques de 15 minutos procesados en el período. Equivale a `sample_blocks` en SQL.
- **C14**: Acumulado de `(racks indisponibles) × bloques`, opcionalmente ponderado por factores excusable y operacional. Equivale a `unavailable_rack_blocks` en SQL.
- **C16**: Disponibilidad del período = `1 - C14 / (Total_Racks × C12)`.
- **C19**: Disponibilidad acumulada anual = `1 - C14 / (total_racks × 365 × 24 × 4)`.
- **NUMBER_OF_MODULES**: Campo `Arena - PCS XX - POWERELECTRONICS HEM-k NUMBER OF MODULES` en `RawData-PCS`. Valor numérico 0–4. Valor < 4 indica indisponibilidad. Vacío se trata como 4 (disponible) con flag de auditoría.
- **modules_available_is_null**: Flag booleano. `True` cuando `NUMBER_OF_MODULES` estaba vacío en el origen.
- **fault_description_fallback**: Flag booleano. `True` cuando la descripción del evento fue tomada del intervalo temporal anterior.
- **excused_factor**: Factor del campo `Excused Event` de PlantActivity (columna D). Valor 0 o 1.
- **operational_factor**: Factor de actividad operacional de PlantActivity (columna C). Valor 0 (inactivo) o 1 (activo).
- **DST**: Daylight Saving Time. Cambio de horario de verano/invierno de Chile. Aplica en septiembre 2026.
- **RawData-PCS**: Hoja del Excel con los registros crudos de estado y falla por intervalo de 15 min y por PCS (244 columnas de datos PCS).
- **PlantActivity**: Hoja del Excel con el factor operacional (columna C) y el indicador de evento excusable (columna D) por intervalo de 15 min.

---

## Requirements

### Requirement 1 — Ingesta y staging

**User Story:** Como ingeniero de datos, quiero leer los archivos origen y conservar una copia inmutable en staging, para que cualquier reprocesamiento pueda partir del mismo dato original y sea completamente trazable.

#### Acceptance Criteria

1. WHEN el Ingestion_Module recibe la ruta de un archivo Excel válido, THE Ingestion_Module SHALL leer la hoja `RawData-PCS` comenzando en la fila 2 y conservando todas las columnas presentes.
2. WHEN el Ingestion_Module recibe la ruta de un archivo Excel válido, THE Ingestion_Module SHALL leer la hoja `PlantActivity` comenzando en la fila 2 y conservando todas las columnas presentes.
3. WHEN el Ingestion_Module termina la lectura, THE Staging_Module SHALL persistir los registros exactamente como llegaron, sin ninguna transformación, asociados al `run_id` de la corrida actual.
4. IF el archivo fuente no existe o no es legible, THEN THE Ingestion_Module SHALL registrar el error en `etl_run.error_message` y marcar la corrida con `status = "failed"`.
5. WHEN el Ingestion_Module lee `RawData-PCS`, THE Ingestion_Module SHALL detectar y reportar como advertencias: huecos en la secuencia de timestamps, timestamps duplicados, timestamps fuera de orden ascendente, e intervalos con frecuencia distinta a la configurada.
6. WHEN el Ingestion_Module encuentra una celda vacía en la columna A (timestamp) de `RawData-PCS`, THE Ingestion_Module SHALL registrar la posición del hueco como advertencia y continuar el procesamiento, en lugar de detenerlo silenciosamente como hace el VBA.
7. THE Staging_Module SHALL conservar para cada registro de `RawData-PCS` tanto el valor serial numérico de fecha Excel (`source_excel_serial_datetime`) como el timestamp interpretado en hora local de Chile (`source_timestamp_local`).
8. WHEN el Ingestion_Module detecta anomalías de calidad de datos, THE ETL_Engine SHALL incluir un resumen de esas anomalías en el registro `etl_run` de la corrida sin abortar el procesamiento.

---

### Requirement 2 — Normalización

**User Story:** Como desarrollador del motor de disponibilidad, quiero los datos PCS en formato largo (1 fila por PCS×timestamp), para que el motor pueda procesar cada combinación sin depender del índice de columna frágil del Excel.

#### Acceptance Criteria

1. WHEN el Normalization_Module recibe los datos crudos de `RawData-PCS`, THE Normalization_Module SHALL transformar el formato ancho de 244 columnas (61 PCS × 4 columnas) en formato largo con una fila por `(sample_timestamp, pcs_number)`.
2. THE Normalization_Module SHALL extraer para cada PCS los cuatro campos: `fault_code_raw` (CURRENT FAULT), `status_raw` (STATUS), `warning_raw` (WARNING) y `modules_available` (NUMBER OF MODULES), usando el patrón de nombre `Arena - PCS XX - POWERELECTRONICS ...` donde `XX` va de `01` a `61`.
3. WHEN el Normalization_Module encuentra que el campo `NUMBER_OF_MODULES` está vacío para un `(sample_timestamp, pcs_number)`, THE Normalization_Module SHALL asignar `modules_available = 4` y marcar `modules_available_is_null = True`.
4. WHEN el Normalization_Module encuentra que el campo `NUMBER_OF_MODULES` contiene un valor numérico, THE Normalization_Module SHALL asignar ese valor a `modules_available` y marcar `modules_available_is_null = False`.
5. THE Normalization_Module SHALL detectar y reportar: PCS faltantes respecto al total configurado, columnas faltantes por PCS, columnas con nombres inesperados, y columnas en posición desplazada.
6. WHEN la normalización finaliza, THE Normalization_Module SHALL producir un conjunto de datos normalizado ordenado por `(sample_timestamp ASC, pcs_number ASC)`.

---

### Requirement 3 — Enriquecimiento

**User Story:** Como desarrollador del motor de disponibilidad, quiero unir los datos PCS normalizados con los factores de PlantActivity, para que el motor pueda aplicar los ponderadores operacional y excusable correctamente en cada intervalo.

#### Acceptance Criteria

1. WHEN el Enrichment_Module recibe los datos normalizados de PCS y los datos de PlantActivity, THE Enrichment_Module SHALL unir ambas fuentes por `sample_timestamp`.
2. THE Enrichment_Module SHALL asignar a cada registro de `raw_pcs_sample` el `operational_factor` (columna C de PlantActivity) y el `excused_factor` (columna D de PlantActivity) correspondientes al mismo timestamp.
3. IF un `sample_timestamp` de `raw_pcs_sample` no tiene correspondencia en `plant_activity_sample`, THEN THE Enrichment_Module SHALL registrar el hueco como advertencia y asignar valores por defecto: `operational_factor = 1`, `excused_factor = 1`.
4. WHEN el enriquecimiento finaliza, THE Enrichment_Module SHALL producir un conjunto de datos enriquecido con todos los campos necesarios para el Availability_Engine: `modules_available`, `modules_available_is_null`, `operational_factor` y `excused_factor`.

---

### Requirement 4 — Motor de disponibilidad

**User Story:** Como analista de disponibilidad, quiero que el motor Python calcule C12, C14 y el KPI del período reproduciéndose exactamente igual que la macro `cmdCalcAvailability`, para poder comparar los resultados antes de retirar el Excel como fuente oficial.

#### Acceptance Criteria

1. WHEN el Availability_Engine procesa un período, THE Availability_Engine SHALL incrementar `C12` en 1 por cada intervalo de 15 minutos válido en el rango de fechas, independientemente del número de PCS.
2. WHEN el Availability_Engine evalúa `modules_available` para un `(sample_timestamp, pcs_number)`, THE Availability_Engine SHALL considerar ese PCS como indisponible únicamente si `modules_available` es numérico y `modules_available < 4`.
3. WHEN el Availability_Engine identifica un PCS indisponible en un intervalo, THE Availability_Engine SHALL calcular `batteries_unavailable = batteries_per_pcs - modules_available`.
4. WHERE el parámetro de corrida `apply_excused_event = "Yes"`, THE Availability_Engine SHALL calcular `weighted_batteries_unavailable = batteries_unavailable * excused_factor`.
5. WHERE el parámetro de corrida `apply_excused_event = "No"`, THE Availability_Engine SHALL calcular `weighted_batteries_unavailable = batteries_unavailable`.
6. WHERE el parámetro de corrida `only_operational_time = "Yes"`, THE Availability_Engine SHALL acumular `C14 += racks_per_pcs * weighted_batteries_unavailable * operational_factor`.
7. WHERE el parámetro de corrida `only_operational_time = "No"`, THE Availability_Engine SHALL acumular `C14 += racks_per_pcs * weighted_batteries_unavailable`.
8. WHEN el Availability_Engine ha procesado todos los intervalos del período, THE Availability_Engine SHALL calcular `availability_period = 1 - C14 / (total_racks * C12)`.
9. IF `C12 = 0` o `total_racks = 0`, THEN THE Availability_Engine SHALL producir `availability_period = None` (equivalente al `IFERROR(..., "N/A")` del Excel).
10. THE Availability_Engine SHALL usar el factor `racks_per_pcs` proveniente de la configuración de la corrida (valor `12` durante la fase de paridad), sin asumir ese valor como constante del código.
11. THE Availability_Engine SHALL producir un registro en `availability_sample_result` por cada `(run_id, sample_timestamp, pcs_number)` incluyendo: `modules_available`, `modules_available_is_null`, `batteries_unavailable`, `excused_factor`, `operational_factor`, `weighted_batteries_unavailable` y `weighted_rack_impact`.

---

### Requirement 5 — Motor de eventos de falla

**User Story:** Como analista de disponibilidad, quiero que el motor Python genere la lista de eventos de falla reproduciéndose exactamente igual que la macro `mcoCreateList`, incluyendo la lógica de fallback en la descripción, para poder reconciliar la lista con el Excel evento por evento.

#### Acceptance Criteria

1. WHEN el Fault_Events_Engine procesa los datos normalizados, THE Fault_Events_Engine SHALL recorrer cada PCS por separado en orden temporal ascendente y agrupar secuencias consecutivas de intervalos donde `modules_available < 4` como un único evento de falla.
2. WHEN el Fault_Events_Engine detecta el inicio de un evento de falla, THE Fault_Events_Engine SHALL calcular `start_timestamp = current_timestamp - sampling_minutes / (24 * 60)` expresado como fracción de día (conservando la semántica exacta del VBA).
3. WHEN el Fault_Events_Engine detecta el fin de un evento de falla, THE Fault_Events_Engine SHALL cerrar el evento cuando el siguiente intervalo tiene `modules_available = 4`, está vacío, o su timestamp supera `period_end + 1 día`.
4. WHEN el Fault_Events_Engine determina la descripción de un evento de falla, THE Fault_Events_Engine SHALL aplicar la siguiente lógica en orden:
   - Si la descripción del intervalo actual es distinta de `"NO FAULTS"` y no está vacía → usar la descripción actual.
   - Si la descripción actual es `"NO FAULTS"` y el intervalo inmediatamente anterior tiene una descripción distinta de `"NO FAULTS"` → usar la descripción del intervalo anterior.
   - Si la descripción actual es `"NO FAULTS"` y el intervalo anterior también es `"NO FAULTS"` o no existe → asignar `"F13 NO MODULES"`.
   - Si la descripción resultante está vacía → asignar `"F1 Watchdog"`.
5. WHEN el Fault_Events_Engine aplica la descripción del intervalo anterior, THE Fault_Events_Engine SHALL marcar el evento con `fault_description_fallback = True`.
6. WHEN el Fault_Events_Engine calcula el código de falla, THE Fault_Events_Engine SHALL extraer el primer token antes del primer espacio en la descripción final; IF la descripción no contiene espacio, THEN THE Fault_Events_Engine SHALL usar `"F" + descripción_completa`.
7. WHEN el Fault_Events_Engine calcula la duración de un evento, THE Fault_Events_Engine SHALL usar `duration_hours = 24 * (end_timestamp - start_timestamp)`.
8. WHEN el Fault_Events_Engine calcula el impacto de un evento, THE Fault_Events_Engine SHALL usar `unavailability_rack_hours = racks_per_pcs * duration_hours * (sumablocks / numBlock)`, donde `racks_per_pcs = 12` proviene de la configuración.
9. WHERE el parámetro `apply_excused_event = "Yes"`, THE Fault_Events_Engine SHALL ponderar cada bloque del acumulador `sumablocks` con el `excused_factor` del timestamp correspondiente.
10. THE Fault_Events_Engine SHALL producir registros en la tabla `fault_event` con todos los campos del modelo de datos: `run_id`, `pcs_number`, `start_timestamp`, `end_timestamp`, `duration_hours`, `fault_code`, `fault_description`, `fault_description_fallback`, `average_batteries_involved` y `unavailable_rack_hours`.

---

### Requirement 6 — Agregaciones

**User Story:** Como analista de disponibilidad, quiero los KPI diarios, mensuales y el acumulado anual calculados con la misma lógica que el Excel (macros `mcoDailyAvailability` y hoja `Annual_AVA`), para comparar los resultados período a período.

#### Acceptance Criteria

1. WHEN el Aggregation_Module calcula el KPI diario para un día, THE Aggregation_Module SHALL sumar todos los valores `weighted_rack_impact` de `availability_sample_result` correspondientes a ese día y `run_id`.
2. WHEN el Aggregation_Module acumula el KPI diario, THE Aggregation_Module SHALL calcular `accumulated_unavailable_rack_blocks` como la suma desde el primer día del período hasta el día actual, inclusive.
3. WHEN el Aggregation_Module calcula la disponibilidad diaria acumulada para el día N del mes, THE Aggregation_Module SHALL usar: `availability = 1 - accumulated_unavailable_rack_blocks / (total_racks * 24 * 60 * N / sampling_minutes)`.
4. IF el denominador de la fórmula de disponibilidad diaria es cero, THEN THE Aggregation_Module SHALL producir `availability = None` para ese día.
5. WHEN el Aggregation_Module calcula la variación diaria, THE Aggregation_Module SHALL calcular `variation = availability_day_N - availability_day_N_minus_1` con manejo de error equivalente a `IFERROR`.
6. WHEN el Aggregation_Module calcula la disponibilidad acumulada anual para una corrida, THE Aggregation_Module SHALL usar la fórmula `availability_annual = 1 - C14 / (total_racks * (365 * 24 * 4))`, conservando el supuesto de 365 días y 4 bloques/hora del Excel durante la fase de paridad.
7. THE Aggregation_Module SHALL producir registros en `annual_availability` con: año, mes, días en el mes, bloques de muestreo, racks indisponibles, disponibilidad mensual, acumulados históricos y disponibilidad acumulada.
8. THE Aggregation_Module SHALL inicializar la acumulación anual desde el **08/Abril/2026** (fecha de inicio de operación de Arena BESS), sin acumular períodos anteriores a esa fecha.

---

### Requirement 7 — Persistencia SQL Server

**User Story:** Como administrador del sistema, quiero que todos los resultados y datos intermedios se almacenen en SQL Server de forma append-only por `run_id`, para que ninguna corrida nueva destruya los resultados históricos y cada corrida sea completamente auditable.

#### Acceptance Criteria

1. THE Persistence_Module SHALL operar en modo append-only: nunca eliminar ni sobreescribir registros existentes en ninguna tabla.
2. THE Persistence_Module SHALL crear un registro en `etl_run` al inicio de cada corrida con `status = "running"` y actualizarlo con `status = "success"` o `status = "failed"` al finalizar.
3. THE Persistence_Module SHALL insertar todos los registros de staging, resultados intermedios y KPI bajo el mismo `run_id` de la corrida activa.
4. THE Persistence_Module SHALL persistir las ocho tablas del modelo de datos: `etl_run`, `raw_pcs_sample`, `plant_activity_sample`, `availability_sample_result`, `availability_run_result`, `fault_event`, `daily_availability` y `annual_availability`.
5. WHEN el Persistence_Module inserta registros en `raw_pcs_sample`, THE Persistence_Module SHALL conservar el número de fila origen (`source_row_number`) y el nombre de las columnas origen (`source_columns`) para trazabilidad.
6. THE Persistence_Module SHALL almacenar en `etl_run` los parámetros completos de la corrida: `total_pcs`, `batteries_per_pcs`, `racks_per_pcs`, `total_racks`, `sampling_minutes`, `only_operational_time`, `apply_excused_event` y `algorithm_version`.
7. IF la conexión a SQL Server falla durante la inserción, THEN THE Persistence_Module SHALL registrar el error, marcar la corrida con `status = "failed"` y no dejar datos parciales sin el correspondiente `etl_run` de error.

---

### Requirement 8 — Reconciliación

**User Story:** Como ingeniero de datos, quiero comparar automáticamente una corrida Python contra una corrida Excel en cinco niveles de detalle, para poder declarar la paridad antes de retirar el Excel como fuente oficial.

#### Acceptance Criteria

1. THE Reconciliation_Module SHALL comparar corridas en cinco niveles: (1) input, (2) muestra individual, (3) acumulados, (4) KPI y (5) eventos.
2. WHEN el Reconciliation_Module ejecuta la comparación de Nivel 1 (input), THE Reconciliation_Module SHALL verificar que coincidan: cantidad total de filas, rango de timestamps, cantidad de PCS, cantidad de registros por PCS, frecuencia de muestreo, fecha inicial y fecha final.
3. WHEN el Reconciliation_Module ejecuta la comparación de Nivel 2 (muestra individual), THE Reconciliation_Module SHALL comparar por cada `(sample_timestamp, pcs_number)`: `modules_available`, `batteries_unavailable`, `excused_factor`, `operational_factor` y `weighted_batteries_unavailable`.
4. WHEN el Reconciliation_Module ejecuta la comparación de Nivel 3 (acumulados), THE Reconciliation_Module SHALL comparar `C12` y `C14` entre la corrida Python y los valores del Excel.
5. WHEN el Reconciliation_Module ejecuta la comparación de Nivel 4 (KPI), THE Reconciliation_Module SHALL comparar `C16` (disponibilidad del período) y `C19` (disponibilidad acumulada anual), así como los KPI diarios y mensuales.
6. WHEN el Reconciliation_Module ejecuta la comparación de Nivel 5 (eventos), THE Reconciliation_Module SHALL comparar campo por campo cada evento de `ListOfFaults`: `pcs_number`, `start_timestamp`, `end_timestamp`, `duration_hours`, `fault_code`, `average_batteries_involved` y `unavailable_rack_hours`.
7. WHEN el Reconciliation_Module compara valores numéricos de punto flotante, THE Reconciliation_Module SHALL usar una tolerancia absoluta de `1e-9` para los cálculos internos y reportar cualquier diferencia que supere ese umbral.
8. WHEN el Reconciliation_Module detecta que el período analizado supera 167 eventos, THE Reconciliation_Module SHALL documentar que el Excel solo ordena hasta 167 eventos y excluir la comparación de posición para los eventos adicionales.
9. WHEN el Reconciliation_Module finaliza, THE Reconciliation_Module SHALL producir un reporte estructurado que indique el resultado (pass/fail) por nivel y el detalle de las discrepancias encontradas.

---

### Requirement 9 — Auditoría y trazabilidad

**User Story:** Como auditor del KPI contractual, quiero poder rastrear cualquier valor de disponibilidad hasta su dato de entrada original, para poder responder preguntas sobre por qué se obtuvo un determinado resultado.

#### Acceptance Criteria

1. THE ETL_Engine SHALL garantizar trazabilidad completa desde el KPI final hasta el dato crudo: `availability_run_result` → `availability_sample_result` → `raw_pcs_sample` → registro del Excel original.
2. THE ETL_Engine SHALL identificar cada corrida con un `run_id` único y un `algorithm_version` que permita saber qué versión de la lógica produjo ese resultado.
3. WHEN la ETL_Engine almacena registros de `raw_pcs_sample`, THE ETL_Engine SHALL conservar el `source_row_number` de la fila en el Excel origen y los nombres de columna originales en `source_columns`.
4. THE ETL_Engine SHALL identificar y hacer consultables todos los intervalos con `modules_available_is_null = True` a lo largo de toda la historia del activo, sin suprimir esos registros.
5. THE ETL_Engine SHALL identificar y hacer consultables todos los eventos con `fault_description_fallback = True`, reportando la frecuencia de ese caso en cada corrida.
6. THE ETL_Engine SHALL incluir en cada corrida un reporte de calidad de datos que indique: cantidad de intervalos con `modules_available_is_null`, cantidad de eventos con `fault_description_fallback`, anomalías de timestamp detectadas, y PCS o columnas faltantes.

---

### Requirement 10 — Configuración y versionado

**User Story:** Como desarrollador, quiero que todos los parámetros de negocio provengan de la configuración de la corrida y no estén hardcodeados, para que el sistema pueda adaptarse a futuros cambios sin modificar el código del motor.

#### Acceptance Criteria

1. THE ETL_Engine SHALL leer todos los parámetros de negocio desde la configuración de la corrida: `total_pcs`, `batteries_per_pcs`, `racks_per_pcs`, `sampling_minutes`, `project_start_date`, `only_operational_time` y `apply_excused_event`.
2. THE ETL_Engine SHALL calcular `total_racks = total_pcs * batteries_per_pcs * racks_per_pcs` a partir de los parámetros de configuración, sin usar el valor 2.928 como constante.
3. THE ETL_Engine SHALL asignar a cada corrida un `algorithm_version` que permita distinguir la versión de la lógica de cálculo. El valor por defecto para la fase de paridad es `"availability-v1-excel-parity"`.
4. WHEN se modifica una regla de negocio en una versión futura del algoritmo, THE ETL_Engine SHALL crear una nueva versión del algoritmo y nunca sobreescribir resultados históricos calculados con versiones anteriores.
5. THE ETL_Engine SHALL proporcionar valores por defecto para la primera versión de paridad: `project_name = "Arena BESS"`, `project_start_date = "2026-04-08"`, `total_pcs = 61`, `batteries_per_pcs = 4`, `racks_per_pcs = 12`, `sampling_minutes = 15`.

---

### Requirement 11 — Tratamiento de timestamps y DST

**User Story:** Como ingeniero de datos, quiero que el sistema preserve los timestamps originales del Excel y maneje correctamente el cambio de horario de Chile, para que la paridad no se rompa en períodos que atraviesan un cambio de DST.

#### Acceptance Criteria

1. THE Ingestion_Module SHALL conservar para cada registro tanto el número serial de fecha Excel original (`source_excel_serial_datetime`) como el timestamp interpretado en hora local de Chile (`source_timestamp_local`).
2. THE ETL_Engine SHALL definir explícitamente la zona horaria de negocio como `America/Santiago` para la interpretación de timestamps.
3. THE ETL_Engine SHALL no convertir timestamps a UTC sin documentar el efecto de esa conversión en el reporte de la corrida.
4. THE Reconciliation_Module SHALL validar especialmente los períodos que atraviesan cambios de DST, comparando los timestamps de inicio y fin de eventos entre Excel y Python.
5. WHEN el Ingestion_Module lee timestamps de `RawData-PCS`, THE Ingestion_Module SHALL calcular la frecuencia de muestreo efectiva usando `ROUND((ts[i+1] - ts[i]) * 24 * 60, 2)` y reportar cualquier diferencia respecto a la frecuencia configurada.

---

### Requirement 12 — Golden Reference: comparacion obligatoria contra valores reales del Excel

**User Story:** Como ingeniero de datos, quiero que cada corrida Python sea validada automaticamente contra los valores reales calculados por el Excel (golden references), para garantizar paridad numerica exacta antes de declarar que el motor Python es equivalente al VBA.

#### Acceptance Criteria

1. THE ETL_Engine SHALL incluir un conjunto de golden references en `tests/golden/data/`, uno por cada mes con corrida VBA validada, en formato JSON con el esquema definido en `golden_index.json`.
2. WHEN se agrega un nuevo mes al golden index, THE Reconciliation_Module SHALL comparar automaticamente los resultados de la corrida Python contra ese golden reference usando las tolerancias definidas.
3. THE golden reference de cada mes SHALL contener como minimo: parametros de la corrida (total_pcs, batteries_per_pcs, total_racks, sampling_minutes, only_operational_time, apply_excused_event), KPI del periodo (C12, C14, C16, C19), disponibilidad diaria completa del mes, snapshot de Annual_AVA, y los primeros 20 eventos de ListOfFaults con el total de rack-hours.
4. WHEN el Reconciliation_Module compara KPI del periodo (C16, C19), THE Reconciliation_Module SHALL usar tolerancia absoluta de **1e-6**.
5. WHEN el Reconciliation_Module compara C12 (sample_blocks), THE Reconciliation_Module SHALL usar igualdad exacta de enteros.
6. WHEN el Reconciliation_Module compara C14 (unavailable_rack_blocks), THE Reconciliation_Module SHALL usar tolerancia absoluta de **1e-6**.
7. WHEN el Reconciliation_Module compara disponibilidad diaria (availability por dia), THE Reconciliation_Module SHALL usar tolerancia absoluta de **1e-6**.
8. WHEN el Reconciliation_Module compara total de rack-hours de ListOfFaults, THE Reconciliation_Module SHALL usar tolerancia absoluta de **1e-3**.
9. IF cualquier comparacion contra el golden reference falla, THEN THE ETL_Engine SHALL marcar la corrida con `status = "parity_failed"` y producir un reporte detallado con las discrepancias.
10. THE golden reference SHALL ser extraido del Excel usando el script `tests/golden/extract_golden.py`, que usa `openpyxl` con `data_only=True` para leer los valores calculados por la macro VBA, no las formulas.
11. THE golden index en `tests/golden/data/golden_index.json` SHALL ser la fuente de verdad sobre que meses tienen golden reference validado y cuales estan pendientes.

#### Golden References disponibles

| Mes | Archivo | C12 | C16 | C19 | Estado |
|---|---|---|---|---|---|
| Julio 2026  | `golden_2026_07_july.json`   | 2976 | 0.9503529130 | 0.9957833981 | ✅ verificado |
| Agosto 2026 | `golden_2026_08_august.json` | 2976 | 0.9394752689 | 0.9948595434 | ✅ verificado |
| Septiembre 2026 (parcial 1-21) | `golden_2026_09_september.json` | 1975 | 0.9819924099 | 0.9989850174 | ✅ verificado |

#### Tolerancias de comparacion

| Campo | Tolerancia | Justificacion |
|---|---|---|
| C12 (sample_blocks) | exacto (int) | Contador de enteros, no puede diferir |
| C14 (unavailable_rack_blocks) | abs <= 1e-6 | Suma de floats; diferencias por orden de operaciones son aceptables |
| C16 (availability_period) | abs <= 1e-6 | KPI presentado; precision suficiente para 6 decimales |
| C19 (accumulated_annual) | abs <= 1e-6 | Idem |
| daily availability | abs <= 1e-6 | Por dia; tolerancia igual al KPI del periodo |
| total rack-hours (ListOfFaults) | abs <= 1e-3 | Suma de muchos floats; margen mayor aceptable |
| event duration_hours | abs <= 0.01 | Diferencias de 36 segundos son irrelevantes operacionalmente |
| event rack_hours individual | abs <= 0.1 | Derivado de duration y avg_batteries |

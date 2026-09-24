# Implementation Plan: ETL Arena Availability

## Overview

Este plan implementa el sistema de reingeniería del cálculo de disponibilidad Arena BESS en Python + SQL Server, reproduciendo exactamente la lógica VBA del Excel para obtener paridad completa. El lenguaje de implementación es **Python 3.11+** con pandas, openpyxl, SQLAlchemy, pyodbc e Hypothesis.

Las tareas siguen el orden de las fases ETL: estructura base → ingesta → normalización → DDL → enriquecimiento y motores → agregaciones → persistencia → reconciliación → orquestador.

---

## Tasks

- [ ] 1. Estructura base y configuración del proyecto
  - Crear la estructura de directorios: `src/config/`, `src/ingestion/`, `src/staging/`, `src/normalization/`, `src/enrichment/`, `src/availability/`, `src/fault_events/`, `src/aggregation/`, `src/persistence/`, `src/reconciliation/`, `src/reporting/`, `tests/unit/`, `tests/property/`, `tests/integration/`, `sql/`
  - Crear `pyproject.toml` con dependencias exactas pinneadas: `pandas>=2.2`, `openpyxl>=3.1`, `sqlalchemy>=2.0`, `pyodbc>=5.0`, `pytz>=2024.1`, `hypothesis>=6.100`, `pytest>=8.0`, `pytest-cov>=5.0`
  - Crear `src/__init__.py` y `src/config/__init__.py`
  - Crear `src/config/models.py` con el dataclass `CalculationConfig` (campos: `run_id`, `algorithm_version`, `project_name`, `project_start_date`, `period_start`, `period_end`, `total_pcs`, `batteries_per_pcs`, `racks_per_pcs`, `total_racks`, `sampling_minutes`, `only_operational_time`, `apply_excused_event`, `source_file`, `timezone`)
  - Crear `src/config/defaults.py` con el diccionario `DEFAULT_CONFIG` para la fase de paridad: `total_pcs=61`, `batteries_per_pcs=4`, `racks_per_pcs=12`, `sampling_minutes=15`, `timezone="America/Santiago"`, `algorithm_version="availability-v1-excel-parity"`, `project_start_date=date(2026, 4, 8)`
  - Implementar la propiedad calculada `total_racks = total_pcs * batteries_per_pcs * racks_per_pcs` en `CalculationConfig.__post_init__`
  - _Requirements: 10.1, 10.2, 10.3, 10.5_

  - [ ]* 1.1 Tests de CalculationConfig
    - Verificar que `total_racks` se calcula correctamente
    - Verificar que los valores por defecto de paridad son los correctos
    - Verificar que no se puede pasar `total_racks` manualmente inconsistente
    - _Requirements: 10.2_

- [ ] 2. Módulo de ingesta
  - [ ] 2.1 Implementar ExcelIngestionService en `src/ingestion/excel_reader.py`
    - Método `read_raw_pcs(path)`: leer hoja `RawData-PCS` desde fila 2 con `openpyxl` en modo `read_only=True`, conservar todas las columnas, añadir columna `source_excel_serial_datetime` con el serial numérico de la celda de fecha antes de convertirla
    - Método `read_plant_activity(path)`: leer hoja `PlantActivity` desde fila 2
    - Retornar DataFrames de pandas; la columna de timestamp se convierte a `datetime` con zona `America/Santiago` usando `pytz`
    - Manejar `FileNotFoundError` y errores de hoja no encontrada lanzando excepciones descriptivas
    - _Requirements: 1.1, 1.2, 1.7, 11.1, 11.2_

  - [ ] 2.2 Implementar detección de anomalías en `src/ingestion/anomaly_detector.py`
    - Dataclass `AnomalyReport(anomaly_type, row_number, timestamp, detail)`
    - Función `detect_timestamp_anomalies(df, sampling_minutes)` que detecta: huecos (salto > sampling_minutes), duplicados, fuera de orden, intervalos con frecuencia distinta; calcula frecuencia efectiva con `ROUND((ts[i+1]-ts[i])*24*60, 2)`
    - No abortar el procesamiento; retornar lista de `AnomalyReport`
    - _Requirements: 1.5, 1.6, 11.5_

  - [ ]* 2.3 Tests de ingesta
    - Test de ejemplo: crear Excel de muestra pequeño, verificar que `read_raw_pcs` retorna las filas correctas con `source_excel_serial_datetime` correcto
    - Test de ejemplo: verificar que un hueco de 30 min en una secuencia de 15 min es detectado
    - Test de ejemplo: verificar que `FileNotFoundError` se propaga correctamente
    - _Requirements: 1.4, 1.5_

- [ ] 3. Módulo de normalización
  - [ ] 3.1 Implementar PCSNormalizer en `src/normalization/pcs_normalizer.py`
    - Método `normalize(raw_df, config)`: transformar formato ancho (244 cols) a formato largo con columnas `sample_timestamp`, `pcs_number`, `fault_description_raw`, `fault_code_raw`, `status_raw`, `warning_raw`, `modules_available`, `modules_available_is_null`, `source_row_number`
    - Usar regex `r"Arena - PCS (\d{2}) - POWERELECTRONICS (.*)"` para detectar columnas PCS
    - Aplicar regla de nulo: si `NUMBER_OF_MODULES` vacío → `modules_available=4`, `modules_available_is_null=True`
    - Ordenar salida por `(sample_timestamp ASC, pcs_number ASC)`
    - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.6_

  - [ ] 3.2 Implementar detección de anomalías de esquema en `src/normalization/schema_validator.py`
    - Función `detect_schema_anomalies(raw_df, config)`: detectar PCS faltantes respecto a `config.total_pcs`, columnas faltantes por PCS (de las 4 esperadas), nombres inesperados, columnas desplazadas
    - Retornar lista de `AnomalyReport`
    - _Requirements: 2.5_

  - [ ]* 3.3 Property test: normalización produce forma correcta y marca nulos consistentemente
    - **Property 1: Normalización produce forma correcta y marca nulos consistentemente**
    - **Validates: Requirements 2.1, 2.3, 2.4**
    - Usar `@given` de Hypothesis para generar DataFrames anchos con N timestamps (1–200) y P PCS (1–61)
    - Verificar: salida tiene N×P filas; para toda celda nula en origen `modules_available=4` y `modules_available_is_null=True`; para toda celda numérica `modules_available_is_null=False`
    - Tag: `# Feature: etl-arena-availability, Property 1`

  - [ ]* 3.4 Property test: detección de anomalías de timestamp es exhaustiva
    - **Property 2: Detección de anomalías de timestamp es exhaustiva**
    - **Validates: Requirements 1.5, 1.6**
    - Generar secuencias de timestamps con posiciones y tipos de anomalías aleatorias inyectadas
    - Verificar que cada anomalía inyectada aparece en el reporte de detección
    - Tag: `# Feature: etl-arena-availability, Property 2`

- [ ] 4. DDL SQL Server
  - [ ] 4.1 Crear scripts DDL en `sql/`
    - `sql/01_create_tables.sql`: DDL completo de las 10 tablas según el diseño, en este orden: (1) `tipo_detencion` con INSERT de los 68 códigos PCS-Fault, (2) `proyecto` con INSERT de los 4 proyectos Arena/Copiapó A/Luz del Norte/María Elena, (3) `etl_run` (con campo `id_proyecto`), (4) `raw_pcs_sample`, (5) `plant_activity_sample`, (6) `availability_sample_result`, (7) `availability_run_result`, (8) `fault_event`, (9) `detencion`, (10) `daily_availability`, (11) `annual_availability` — incluyendo índices definidos en el design.md
    - `sql/02_create_indexes.sql`: crear los índices adicionales `IX_raw_pcs_sample_run_ts_pcs`, `IX_plant_activity_run_ts`, `IX_avail_sample_run_ts_pcs`, `IX_fault_event_run_pcs`, `IX_daily_avail_run_day`, `IX_annual_avail_run_ym`
    - `sql/03_queries_audit.sql`: consultas de auditoría reutilizables: trazabilidad KPI→muestra→raw, listado de intervalos con `modules_available_is_null=1`, listado de eventos con `fault_description_fallback=1`
    - _Requirements: 7.4, 9.1_

  - [ ] 4.2 Crear script de inicialización de base de datos
    - `sql/00_create_database.sql`: script para crear la base de datos si no existe
    - `src/persistence/schema.py`: constantes con los nombres de tablas y funciones auxiliares para verificar que el esquema existe antes de la primera corrida
    - _Requirements: 7.1_

  - [ ] 4.3 Crear script de datos maestros en `sql/`
    - `sql/04_seed_tipo_detencion.sql`: INSERT con los 68 códigos del catálogo PCS-Fault (idempotente con MERGE o IF NOT EXISTS)
    - `sql/05_seed_proyectos.sql`: INSERT con los 4 proyectos: Arena (id=1, en_ejecucion, fecha_inicio=2026-04-08, num_pcs=61, batteries=4, racks=12), Copiapó A (id=2), Luz del Norte (id=3), María Elena (id=4) — idempotente
    - _Requirements: 13.1, 13.2, 13.3_

- [ ] 5. Checkpoint — Estructura base completa
  - Verificar que `pyproject.toml` instala correctamente todas las dependencias.
  - Verificar que los tests unitarios de `CalculationConfig` pasan.
  - Verificar que el DDL ejecuta sin errores en SQL Server de desarrollo.
  - Preguntar al usuario si hay ajustes antes de continuar con los motores.

- [ ] 6. Módulo de enriquecimiento y motor de disponibilidad
  - [ ] 6.1 Implementar EnrichmentService en `src/enrichment/enrichment_service.py`
    - Método `enrich(normalized_df, plant_activity_df)`: join por `sample_timestamp`
    - Extraer `operational_factor` (columna C de PlantActivity, valor 0/1) y `excused_factor` (columna D, valor 0/1)
    - Para timestamps sin correspondencia en PlantActivity: asignar `operational_factor=1`, `excused_factor=1`, registrar advertencia
    - _Requirements: 3.1, 3.2, 3.3, 3.4_

  - [ ] 6.2 Implementar AvailabilityEngine en `src/availability/availability_engine.py`
    - Dataclass `AvailabilityResult(c12, c14, availability_period, availability_annual, sample_results)`
    - Método `calculate(enriched_df, config)`: implementar el algoritmo exacto del VBA transcrito en el design.md
    - Iterar timestamps en orden ASC; incrementar `c12` una vez por timestamp
    - Condición de falla: `not is_null and modules < batteries_per_pcs`
    - Aplicar `excused_factor` si `config.apply_excused_event`
    - Aplicar `operational_factor` si `config.only_operational_time`
    - Acumular en `c14` usando `racks_per_pcs` de la configuración
    - Calcular `availability_period = 1 - c14 / (total_racks * c12)` con manejo de c12=0
    - Calcular `availability_annual = 1 - c14 / (total_racks * 365 * 24 * 4)` (supuesto del Excel)
    - Producir DataFrame `sample_results` con todos los campos de `availability_sample_result`
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9, 4.10, 4.11_

  - [ ]* 6.3 Property test: rack impact es correcto para toda combinación de parámetros
    - **Property 3: Cálculo de rack impact ponderado es correcto para toda combinación de parámetros**
    - **Validates: Requirements 4.3, 4.4, 4.5, 4.6, 4.7**
    - Generar combinaciones de `(modules_available ∈ [0,4], excused_factor ∈ {0,1}, operational_factor ∈ {0,1}, apply_excused_event, only_operational_time)`
    - Verificar que `weighted_rack_impact` coincide con la evaluación directa de la fórmula
    - Tag: `# Feature: etl-arena-availability, Property 3`

  - [ ]* 6.4 Property test: C12 siempre cuenta intervalos temporales, no PCS
    - **Property 4: C12 siempre cuenta intervalos temporales, no PCS**
    - **Validates: Requirements 4.1**
    - Generar N timestamps con P PCS; verificar que `result.c12 == N`
    - Tag: `# Feature: etl-arena-availability, Property 4`

  - [ ]* 6.5 Property test: C14 es igual a la suma de todos los weighted_rack_impact
    - **Property 5: C14 es igual a la suma de todos los weighted_rack_impact**
    - **Validates: Requirements 4.1, 4.6, 4.7, 4.11**
    - Verificar que `result.c14 == result.sample_results["weighted_rack_impact"].sum()`
    - Tag: `# Feature: etl-arena-availability, Property 5`

  - [ ]* 6.6 Tests unitarios del motor de disponibilidad
    - Test de ejemplo: 3 intervalos, 2 PCS con valores conocidos; verificar C12=3, C14 exacto, availability_period exacto
    - Test de ejemplo: todos los PCS disponibles; verificar C14=0, availability_period=1.0
    - Test de ejemplo: C12=0 → availability_period=None
    - _Requirements: 4.8, 4.9_

- [ ] 7. Motor de eventos de falla
  - [ ] 7.1 Implementar FaultEventsEngine en `src/fault_events/fault_events_engine.py`
    - Dataclass `FaultEvent(run_id, pcs_number, start_timestamp, end_timestamp, duration_hours, fault_code, fault_description, fault_description_fallback, average_batteries_involved, unavailable_rack_hours)`
    - Método `detect_events(enriched_df, config)`: implementar el algoritmo exacto del VBA transcrito en el design.md
    - Para cada PCS: iterar en orden temporal ASC; detectar inicio/fin de evento basado en `modules_available < batteries_per_pcs`
    - `start_timestamp = current_ts - timedelta(minutes=sampling_minutes)` (semántica exacta del VBA)
    - Aplicar árbol de decisión de descripción con fallback (4 casos según el design.md)
    - Marcar `fault_description_fallback=True` cuando se usa descripción del intervalo anterior
    - Calcular `duration_hours`, `average_batteries_involved`, `unavailable_rack_hours`
    - Implementar `_extract_fault_code(description)`: primer token antes del espacio, o `"F" + description` si no hay espacio
    - Cerrar eventos abiertos al final del período con `end_ts = period_end + timedelta(days=1)`
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7, 5.8, 5.9, 5.10_

  - [ ]* 7.2 Property test: eventos cubren exactamente las sub-secuencias con modules_available < 4
    - **Property 6: Los eventos de falla cubren exactamente las sub-secuencias con modules_available < 4**
    - **Validates: Requirements 5.1, 5.3**
    - Generar secuencias aleatorias de `modules_available` (0–4 + None); verificar cobertura exacta
    - Tag: `# Feature: etl-arena-availability, Property 6`

  - [ ]* 7.3 Property test: lógica de fallback de descripción es determinista y completa
    - **Property 7: La lógica de fallback de descripción es determinista y completa**
    - **Validates: Requirements 5.4, 5.5**
    - Generar pares `(current_description, previous_description)` con valores `"NO FAULTS"`, `""` y cadenas arbitrarias
    - Verificar resultado del árbol de decisión y que `fault_description_fallback` es correcto
    - Tag: `# Feature: etl-arena-availability, Property 7`

  - [ ]* 7.4 Tests unitarios del motor de eventos
    - Test de ejemplo: los cuatro casos del árbol de decisión de descripción con valores concretos
    - Test de ejemplo: evento abierto al final del período se cierra correctamente
    - Test de ejemplo: dos eventos separados en el mismo PCS
    - Test de ejemplo: `_extract_fault_code` con descripción con espacio y sin espacio
    - _Requirements: 5.4, 5.6_

- [ ] 8. Checkpoint — Motores de cálculo completos
  - Ejecutar todos los tests unitarios y de propiedad de los módulos 6 y 7.
  - Verificar con datos de muestra conocidos que C12, C14 y la lista de eventos coinciden con los valores del Excel.
  - Preguntar al usuario si hay ajustes antes de continuar con las agregaciones.

- [ ] 9. Módulo de agregaciones
  - [ ] 9.1 Implementar DailyAggregation en `src/aggregation/daily.py`
    - Método `calculate(sample_results, config)`: agrupar `weighted_rack_impact` por día
    - Calcular `accumulated_unavailable_rack_blocks` como suma acumulada desde el primer día
    - Fórmula: `availability = 1 - accum / (total_racks * 24 * 60 * day_N / sampling_minutes)` con manejo de denominador cero
    - Calcular `variation = availability_N - availability_(N-1)` con manejo de None
    - Retornar DataFrame con campos de `daily_availability`
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5_

  - [ ] 9.2 Implementar AnnualAggregation en `src/aggregation/annual.py`
    - Método `calculate(result, config, month, year)`: calcular KPI mensual y acumulado anual
    - Fórmula anual (conservada del Excel para paridad): `1 - C14 / (total_racks * 365 * 24 * 4)`
    - Inicializar acumulación desde `config.project_start_date` (2026-04-08)
    - Retornar dict/registro con campos de `annual_availability`
    - _Requirements: 6.6, 6.7, 6.8_

  - [ ]* 9.3 Property test: disponibilidad diaria acumulada es consistente con rack blocks
    - **Property 8: La disponibilidad diaria acumulada es consistente con la acumulación de rack blocks**
    - **Validates: Requirements 6.3, 6.4**
    - Generar secuencias de `accumulated_unavailable_rack_blocks` con denominador no nulo
    - Verificar que `0 <= availability <= 1` y que la fórmula es exacta
    - Tag: `# Feature: etl-arena-availability, Property 8`

  - [ ]* 9.4 Tests unitarios de agregaciones
    - Test de ejemplo: 5 días conocidos; verificar disponibilidad diaria y variación
    - Test de ejemplo: día_N=0 → availability=None
    - Test de ejemplo: fórmula anual con C14 y total_racks conocidos
    - _Requirements: 6.3, 6.6_

- [ ] 10. Módulo de persistencia
  - [ ] 10.1 Implementar PersistenceService en `src/persistence/persistence_service.py`
    - Usar `SQLAlchemy Core` con driver `pyodbc`; leer connection string de variable de entorno `ETL_ARENA_DB_URL`
    - Métodos: `begin_run`, `complete_run`, `fail_run`, `save_raw_pcs`, `save_plant_activity`, `save_sample_results`, `save_run_result`, `save_fault_events`, `save_daily`, `save_annual`
    - Añadir método `save_detenciones(run_id, id_proyecto, fault_events)`: convierte cada `FaultEvent` en un registro `detencion`, resuelve `id_tipo_detencion` buscando en `tipo_detencion` por `fault_code`, calcula `duracion_segundos = DATEDIFF(seconds, fecha_inicio, fecha_termino)`, inserta en bulk
    - Inserciones en bulk con `executemany`; nunca `DELETE` ni `UPDATE` sobre registros de corridas anteriores
    - En caso de error: hacer rollback, llamar `fail_run` antes de relanzar la excepción
    - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5, 7.6, 7.7_

  - [ ] 10.2 Implementar gestión de run_id en `src/persistence/run_manager.py`
    - Función `generate_run_id()`: genera UUID v4 como string
    - Función `create_run(config)`: llama `begin_run` y retorna el `run_id`
    - _Requirements: 7.2, 9.2_

  - [ ]* 10.3 Tests de integración de persistencia
    - Test de integración: dos corridas con `run_id` distintos; verificar que ambos conjuntos de registros coexisten (append-only)
    - Test de integración: simular fallo en `save_sample_results`; verificar que `etl_run.status = "failed"` y no quedan datos parciales huérfanos
    - _Requirements: 7.1, 7.7_

- [ ] 11. Módulo de reconciliación
  - [ ] 11.1 Implementar ReconciliationService en `src/reconciliation/reconciliation_service.py`
    - Dataclasses `LevelResult(level, name, passed, discrepancies)` y `ReconciliationReport(python_run_id, excel_source, tolerance, levels, overall_passed)`
    - Método `reconcile(python_run_id, excel_export_path, tolerance=1e-9)`: ejecutar los 5 niveles de comparación
    - _Requirements: 8.1_

  - [ ] 11.2 Implementar los 5 niveles de comparación en `src/reconciliation/levels.py`
    - `compare_level_1_input(python_run, excel_df)`: filas, timestamps, PCS count, frecuencia, rango de fechas
    - `compare_level_2_samples(python_samples, excel_samples, tolerance)`: por cada `(ts, pcs)` comparar `modules_available`, `batteries_unavailable`, factores, `weighted`
    - `compare_level_3_accumulators(python_result, excel_c12, excel_c14, tolerance)`: C12 y C14
    - `compare_level_4_kpi(python_result, excel_kpis, tolerance)`: C16, C19, KPI diario, KPI mensual
    - `compare_level_5_events(python_events, excel_events, tolerance, max_ordered=167)`: campo por campo; advertencia si supera 167 eventos
    - _Requirements: 8.2, 8.3, 8.4, 8.5, 8.6, 8.7, 8.8_

  - [ ]* 11.3 Tests de reconciliación
    - Test de ejemplo: reconciliación con datos idénticos → reporte overall_passed=True
    - Test de ejemplo: reconciliación con discrepancia conocida en C14 → reporte indica fail en nivel 3
    - Test de ejemplo: más de 167 eventos → advertencia registrada en nivel 5
    - _Requirements: 8.7, 8.8, 8.9_

- [ ] 12. Checkpoint — Pipeline completo
  - Ejecutar todos los tests unitarios, de propiedad e integración.
  - Verificar que todos los tests pasan antes de continuar.
  - Preguntar al usuario si hay ajustes antes del orquestador.

- [ ] 13. Pipeline orquestador
  - [ ] 13.1 Implementar script principal `run_etl.py` en la raíz del proyecto
    - Parsear argumentos CLI: `--excel-path`, `--period-start`, `--period-end`, `--only-operational-time`, `--apply-excused-event`, `--algorithm-version`
    - Orquestar las fases en orden: `config → ingestion → staging → normalization → enrichment → availability → fault_events → aggregation → persistence → (opcional: reconciliation)`
    - Registrar inicio y fin de cada fase en el log con timestamp
    - En caso de error: llamar `persistence.fail_run` antes de terminar con exit code 1
    - _Requirements: 7.2, 9.1, 9.2, 9.6_

  - [ ] 13.2 Implementar reporte de calidad de datos en `src/reporting/quality_report.py`
    - Función `generate_quality_report(run_id, anomalies, sample_results, fault_events)`: producir resumen con: cantidad de intervalos `modules_available_is_null`, cantidad de eventos `fault_description_fallback`, anomalías de timestamp, PCS o columnas faltantes
    - Imprimir el reporte en consola y persistir en `etl_run.error_message` (o en un campo dedicado si se agrega)
    - _Requirements: 9.6_

  - [ ]* 13.3 Tests de integración end-to-end
    - Test de integración: ejecutar el pipeline completo sobre un subconjunto de datos de muestra del Excel real (extraídos manualmente)
    - Verificar que los KPI producidos coinciden con los valores conocidos del Excel con tolerancia `1e-4`
    - Verificar que todos los registros tienen el mismo `run_id` y que `etl_run.status = "success"`
    - _Requirements: 9.1, 10.3_

- [ ] 13.4 Implementar golden tests en `tests/golden/`
    - Crear `tests/golden/conftest.py` con fixture `golden_data(month_id)` que carga el JSON correspondiente de `tests/golden/data/` usando el indice `golden_index.json`
    - Crear `tests/golden/test_golden_runner.py`: test parametrizado que itera sobre todos los meses en `golden_index.json` con status="verified" y verifica: C12 exacto (int), C14/C16/C19 con tolerancia 1e-6, disponibilidad diaria con tolerancia 1e-6 por dia, total rack-hours con tolerancia 1e-3
    - Crear `tests/golden/test_gt_july_2026.py`: test especifico para julio 2026 usando `tests/golden/data/golden_2026_07_july.json` — verifica C12=2976, C14=432611.2119999999, C16=0.9503529130126623, C19=0.9957833980914864, los 31 dias de disponibilidad diaria, y los primeros 10 eventos del PCS 1
    - Crear `tests/golden/test_gt_august_2026.py`: test especifico para agosto 2026 usando `tests/golden/data/golden_2026_08_august.json` — verifica C12=2976, C14=527396.0440000003, C16=0.9394752689090135, C19=0.9948595433867929, los 31 dias de disponibilidad diaria, y los primeros 10 eventos del PCS 1
    - Crear `tests/golden/test_gt_september_2026.py`: test especifico para septiembre 2026 (periodo parcial 01-21 sep) usando `tests/golden/data/golden_2026_09_september.json` — verifica C12=1975, C14=104134.2920000001, C16=0.9819924099052362, C19=0.9989850173961998, los 21 dias con datos reales y los 9 dias con daily=0, y los primeros 10 eventos del PCS 1
    - Crear `tests/golden/test_gt_decimal_modules.py`: verificar que el PCS 04 en el primer bloque (2026-04-08 00:15, modules=3.234...) produce batteries_unavailable correcto (caso de valor decimal en NUMBER_OF_MODULES — documentado en GT-7 del design.md)
    - El script `tests/golden/extract_golden.py` ya existe y debe usarse para agregar meses futuros: `python tests/golden/extract_golden.py --excel <ruta> --month <N> --year 2026 --label <nombre>`
    - _Requirements: 12.1, 12.2, 12.3, 12.4, 12.5, 12.6, 12.7, 12.8, 12.9_

- [ ] 14. Checkpoint final — Verificación de paridad
  - Ejecutar la suite completa de tests con `pytest --cov=src tests/`
  - Verificar cobertura de tests > 80% en los modulos `availability` y `fault_events`
  - Ejecutar el pipeline completo sobre el Excel de julio 2026 y verificar que todos los golden tests pasan
  - Ejecutar `pytest tests/golden/ -v` y confirmar que el golden runner reporta paridad para julio 2026
  - Documentar cualquier discrepancia encontrada en el golden index con campo "discrepancies"
  - Instruccion para agregar meses futuros: configurar el Excel para el mes deseado, ejecutar la macro VBA, luego `python tests/golden/extract_golden.py --excel <ruta> --month <N> --year 2026 --label <nombre>`, y agregar la entrada al `golden_index.json`

---

## Notes

- Las subtareas marcadas con `*` son opcionales y pueden omitirse para un MVP más rápido, pero son altamente recomendadas para garantizar la paridad con el Excel
- Los property-based tests usan **Hypothesis** con mínimo 100 iteraciones por propiedad
- Todas las constantes de negocio (`61`, `4`, `12`, `15`) deben provenir de `CalculationConfig`, nunca hardcodeadas en los motores
- SQL Server es append-only por `run_id`; nunca ejecutar `DELETE` ni `UPDATE` sobre datos históricos
- La tolerancia numérica para comparaciones en reconciliación es `1e-9`
- La fórmula anual con `365 * 24 * 4` se conserva sin modificar durante la fase de paridad
- Para la fase de paridad, `algorithm_version = "availability-v1-excel-parity"`

## Task Dependency Graph

```json
{
  "waves": [
    { "id": 0, "tasks": ["1.1"] },
    { "id": 1, "tasks": ["2.1", "2.2", "4.1", "4.2"] },
    { "id": 2, "tasks": ["2.3", "3.1", "3.2"] },
    { "id": 3, "tasks": ["3.3", "3.4"] },
    { "id": 4, "tasks": ["6.1"] },
    { "id": 5, "tasks": ["6.2"] },
    { "id": 6, "tasks": ["6.3", "6.4", "6.5", "6.6", "7.1"] },
    { "id": 7, "tasks": ["7.2", "7.3", "7.4"] },
    { "id": 8, "tasks": ["9.1", "9.2", "10.1", "10.2"] },
    { "id": 9, "tasks": ["9.3", "9.4", "10.3"] },
    { "id": 10, "tasks": ["11.1"] },
    { "id": 11, "tasks": ["11.2"] },
    { "id": 12, "tasks": ["11.3"] },
    { "id": 13, "tasks": ["13.1", "13.2"] },
    { "id": 14, "tasks": ["13.3"] }
  ]
}
```

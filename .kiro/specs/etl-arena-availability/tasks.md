# Implementation Plan: ETL Arena Availability

## Overview

Este plan implementa el sistema de reingeniería del cálculo de disponibilidad Arena BESS en Python + SQL Server, reproduciendo exactamente la lógica VBA del Excel para obtener paridad completa. El lenguaje de implementación es **Python 3.11+** con pandas, openpyxl, SQLAlchemy, pyodbc e Hypothesis.

Las tareas siguen el orden de las fases ETL: estructura base -> ingesta -> normalización -> DDL -> enriquecimiento y motores -> agregaciones -> persistencia -> reconciliación -> orquestador.

---

## Tasks

- [ ] 1. Estructura base y configuración del proyecto
  - Crear la estructura de directorios: `src/config/`, `src/ingestion/`, `src/staging/`, `src/normalization/`, `src/enrichment/`, `src/availability/`, `src/fault_events/`, `src/aggregation/`, `src/persistence/`, `src/reconciliation/`, `src/reporting/`, `tests/unit/`, `tests/property/`, `tests/integration/`, `sql/`
  - Crear `pyproject.toml` con dependencias exactas pinneadas: `pandas>=2.2`, `openpyxl>=3.1`, `sqlalchemy>=2.0`, `pyodbc>=5.0`, `pytz>=2024.1`, `hypothesis>=6.100`, `pytest>=8.0`, `pytest-cov>=5.0`
  - Crear `src/__init__.py` y `src/config/__init__.py`
  - Crear `src/config/models.py` con el dataclass `ConfiguracionCalculo` (campos: `id_corrida`, `version_algoritmo`, `nombre_proyecto`, `fecha_inicio_proyecto`, `inicio_periodo`, `fin_periodo`, `total_pcs`, `baterias_por_pcs`, `racks_por_pcs`, `total_racks`, `minutos_muestreo`, `solo_tiempo_operacional`, `aplicar_evento_excusable`, `archivo_origen`, `zona_horaria`)
  - Crear `src/config/defaults.py` con el diccionario `CONFIG_POR_DEFECTO` para la fase de paridad: `total_pcs=61`, `baterias_por_pcs=4`, `racks_por_pcs=12`, `minutos_muestreo=15`, `zona_horaria="America/Santiago"`, `version_algoritmo="availability-v1-excel-parity"`, `fecha_inicio_proyecto=date(2026, 4, 8)`
  - Implementar la propiedad calculada `total_racks = total_pcs * baterias_por_pcs * racks_por_pcs` en `ConfiguracionCalculo.__post_init__`
  - _Requirements: 10.1, 10.2, 10.3, 10.5_

  - [ ]* 1.1 Tests de ConfiguracionCalculo
    - Verificar que `total_racks` se calcula correctamente
    - Verificar que los valores por defecto de paridad son los correctos
    - Verificar que no se puede pasar `total_racks` manualmente inconsistente
    - _Requirements: 10.2_

- [ ] 2. Módulo de ingesta
  - [ ] 2.1 Implementar ServicioIngesta en `src/ingestion/excel_reader.py`
    - Método `leer_raw_pcs(ruta)`: leer hoja `RawData-PCS` desde fila 2 con `openpyxl` en modo `read_only=True`, conservar todas las columnas, añadir columna `SerialFechaExcelOrigen` con el serial numérico de la celda de fecha antes de convertirla
    - Método `leer_actividad_planta(ruta)`: leer hoja `PlantActivity` desde fila 2
    - Retornar DataFrames de pandas; la columna de timestamp se convierte a `datetime` con zona `America/Santiago` usando `pytz`
    - Manejar `FileNotFoundError` y errores de hoja no encontrada lanzando excepciones descriptivas
    - _Requirements: 1.1, 1.2, 1.7, 11.1, 11.2_

  - [ ] 2.2 Implementar detección de anomalías en `src/ingestion/anomaly_detector.py`
    - Dataclass `ReporteAnomalia(tipo_anomalia, numero_fila, timestamp, detalle)`
    - Función `detectar_anomalias_timestamp(df, minutos_muestreo)` que detecta: huecos (salto > minutos_muestreo), duplicados, fuera de orden, intervalos con frecuencia distinta; calcula frecuencia efectiva con `ROUND((ts[i+1]-ts[i])*24*60, 2)`
    - No abortar el procesamiento; retornar lista de `ReporteAnomalia`
    - _Requirements: 1.5, 1.6, 11.5_

  - [ ]* 2.3 Tests de ingesta
    - Test de ejemplo: crear Excel de muestra pequeño, verificar que `leer_raw_pcs` retorna las filas correctas con `SerialFechaExcelOrigen` correcto
    - Test de ejemplo: verificar que un hueco de 30 min en una secuencia de 15 min es detectado
    - Test de ejemplo: verificar que `FileNotFoundError` se propaga correctamente
    - _Requirements: 1.4, 1.5_

- [ ] 3. Módulo de normalización
  - [ ] 3.1 Implementar NormalizadorPCS en `src/normalization/pcs_normalizer.py`
    - Método `normalizar(df_crudo, config)`: transformar formato ancho (244 cols) a formato largo con columnas `MarcaTiempoMuestra`, `NumeroPCS`, `DescripcionFallaRaw`, `CodigoFallaRaw`, `EstadoRaw`, `AdvertenciaRaw`, `ModulosDisponibles`, `ModulosDisponiblesNulo`, `NumeroFilaOrigen`
    - Usar regex `r"Arena - PCS (\d{2}) - POWERELECTRONICS (.*)"` para detectar columnas PCS
    - Aplicar regla de nulo: si `NUMBER_OF_MODULES` vacío -> `ModulosDisponibles=4`, `ModulosDisponiblesNulo=True`
    - Ordenar salida por `(MarcaTiempoMuestra ASC, NumeroPCS ASC)`
    - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.6_

  - [ ] 3.2 Implementar detección de anomalías de esquema en `src/normalization/schema_validator.py`
    - Función `detectar_anomalias_esquema(df_crudo, config)`: detectar PCS faltantes respecto a `config.total_pcs`, columnas faltantes por PCS (de las 4 esperadas), nombres inesperados, columnas desplazadas
    - Retornar lista de `ReporteAnomalia`
    - _Requirements: 2.5_

  - [ ]* 3.3 Property test: normalización produce forma correcta y marca nulos consistentemente
    - **Property 1: Normalización produce forma correcta y marca nulos consistentemente**
    - **Validates: Requirements 2.1, 2.3, 2.4**
    - Usar `@given` de Hypothesis para generar DataFrames anchos con N timestamps (1-200) y P PCS (1-61)
    - Verificar: salida tiene N×P filas; para toda celda nula en origen `ModulosDisponibles=4` y `ModulosDisponiblesNulo=True`; para toda celda numérica `ModulosDisponiblesNulo=False`
    - Tag: `# Feature: etl-arena-availability, Property 1`

  - [ ]* 3.4 Property test: detección de anomalías de timestamp es exhaustiva
    - **Property 2: Detección de anomalías de timestamp es exhaustiva**
    - **Validates: Requirements 1.5, 1.6**
    - Generar secuencias de timestamps con posiciones y tipos de anomalías aleatorias inyectadas
    - Verificar que cada anomalía inyectada aparece en el reporte de detección
    - Tag: `# Feature: etl-arena-availability, Property 2`

- [ ] 4. DDL SQL Server
  - [ ] 4.1 Crear scripts DDL en `sql/`
    - `sql/01_create_tables.sql`: DDL completo de las 10 tablas según el diseño, en este orden: (1) `tipo_detencion` con INSERT de los 68 códigos PCS-Fault, (2) `proyecto` con INSERT de los 4 proyectos Arena/Copiapo A/Luz del Norte/Maria Elena, (3) `etl_run` (con campo `IdProyecto`), (4) `raw_pcs_sample`, (5) `plant_activity_sample`, (6) `availability_sample_result`, (7) `availability_run_result`, (8) `fault_event`, (9) `detencion`, (10) `daily_availability`, (11) `annual_availability` — incluyendo índices definidos en el design.md
    - `sql/02_create_indexes.sql`: crear los índices adicionales `IX_raw_pcs_sample_corrida_ts_pcs`, `IX_plant_activity_corrida_ts`, `IX_avail_sample_corrida_ts_pcs`, `IX_fault_event_corrida_pcs`, `IX_daily_avail_corrida_dia`, `IX_annual_avail_corrida_anio_mes`
    - `sql/03_queries_audit.sql`: consultas de auditoría reutilizables: trazabilidad KPI->muestra->raw, listado de intervalos con `ModulosDisponiblesNulo=1`, listado de eventos con `DescripcionFallaFallback=1`
    - _Requirements: 7.4, 9.1_

  - [ ] 4.2 Crear script de inicialización de base de datos
    - `sql/00_create_database.sql`: script para crear la base de datos si no existe
    - `src/persistence/schema.py`: constantes con los nombres de tablas y funciones auxiliares para verificar que el esquema existe antes de la primera corrida
    - _Requirements: 7.1_

  - [ ] 4.3 Crear script de datos maestros en `sql/`
    - `sql/04_seed_tipo_detencion.sql`: INSERT con los 68 códigos del catálogo PCS-Fault (idempotente con MERGE o IF NOT EXISTS)
    - `sql/05_seed_proyectos.sql`: INSERT con los 4 proyectos: Arena (IdProyecto=1, en_ejecucion, FechaInicio=2026-04-08, NumPCS=61, baterias=4, racks=12), Copiapo A (IdProyecto=2), Luz del Norte (IdProyecto=3), Maria Elena (IdProyecto=4) — idempotente
    - _Requirements: 13.1, 13.2, 13.3_

- [ ] 5. Checkpoint — Estructura base completa
  - Verificar que `pyproject.toml` instala correctamente todas las dependencias.
  - Verificar que los tests unitarios de `ConfiguracionCalculo` pasan.
  - Verificar que el DDL ejecuta sin errores en SQL Server de desarrollo.
  - Preguntar al usuario si hay ajustes antes de continuar con los motores.

- [ ] 6. Módulo de enriquecimiento y motor de disponibilidad
  - [ ] 6.1 Implementar ServicioEnriquecimiento en `src/enrichment/enrichment_service.py`
    - Método `enriquecer(df_normalizado, df_actividad_planta)`: join por `MarcaTiempoMuestra`
    - Extraer `FactorOperacional` (columna C de PlantActivity, valor 0/1) y `FactorExcusable` (columna D, valor 0/1)
    - Para timestamps sin correspondencia en PlantActivity: asignar `FactorOperacional=1`, `FactorExcusable=1`, registrar advertencia
    - _Requirements: 3.1, 3.2, 3.3, 3.4_

  - [ ] 6.2 Implementar MotorDisponibilidad en `src/availability/availability_engine.py`
    - Dataclass `ResultadoDisponibilidad(bloques_muestreo, bloques_racks_indisponibles, disponibilidad_periodo, disponibilidad_anual_acumulada, resultados_muestra)`
    - Método `calcular(df_enriquecido, config)`: implementar el algoritmo exacto del VBA transcrito en el design.md
    - Iterar timestamps en orden ASC; incrementar `bloques_muestreo` una vez por timestamp
    - Condición de falla: `not es_nulo and modulos < baterias_por_pcs`
    - Aplicar `FactorExcusable` si `config.aplicar_evento_excusable`
    - Aplicar `FactorOperacional` si `config.solo_tiempo_operacional`
    - Acumular en `bloques_racks_indisponibles` usando `racks_por_pcs` de la configuración
    - Calcular `disponibilidad_periodo = 1 - bloques_racks_indisponibles / (total_racks * bloques_muestreo)` con manejo de bloques_muestreo=0
    - Calcular `disponibilidad_anual_acumulada = 1 - bloques_racks_indisponibles / (total_racks * 365 * 24 * 4)` (supuesto del Excel)
    - Producir DataFrame `resultados_muestra` con todos los campos de `availability_sample_result`
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9, 4.10, 4.11_

  - [ ]* 6.3 Property test: ImpactoRackPonderado es correcto para toda combinación de parámetros
    - **Property 3: Cálculo de ImpactoRackPonderado es correcto para toda combinación de parámetros**
    - **Validates: Requirements 4.3, 4.4, 4.5, 4.6, 4.7**
    - Generar combinaciones de `(ModulosDisponibles en [0,4], FactorExcusable en {0,1}, FactorOperacional en {0,1}, aplicar_evento_excusable, solo_tiempo_operacional)`
    - Verificar que `ImpactoRackPonderado` coincide con la evaluación directa de la fórmula
    - Tag: `# Feature: etl-arena-availability, Property 3`

  - [ ]* 6.4 Property test: BloquesMuestreo siempre cuenta intervalos temporales, no PCS
    - **Property 4: BloquesMuestreo siempre cuenta intervalos temporales, no PCS**
    - **Validates: Requirements 4.1**
    - Generar N timestamps con P PCS; verificar que `resultado.bloques_muestreo == N`
    - Tag: `# Feature: etl-arena-availability, Property 4`

  - [ ]* 6.5 Property test: BloquesRacksIndisponibles es igual a la suma de todos los ImpactoRackPonderado
    - **Property 5: BloquesRacksIndisponibles es igual a la suma de todos los ImpactoRackPonderado**
    - **Validates: Requirements 4.1, 4.6, 4.7, 4.11**
    - Verificar que `resultado.bloques_racks_indisponibles == resultado.resultados_muestra["ImpactoRackPonderado"].sum()`
    - Tag: `# Feature: etl-arena-availability, Property 5`

  - [ ]* 6.6 Tests unitarios del motor de disponibilidad
    - Test de ejemplo: 3 intervalos, 2 PCS con valores conocidos; verificar `BloquesMuestreo=3`, `BloquesRacksIndisponibles` exacto, `DisponibilidadPeriodo` exacto
    - Test de ejemplo: todos los PCS disponibles; verificar `BloquesRacksIndisponibles=0`, `DisponibilidadPeriodo=1.0`
    - Test de ejemplo: `BloquesMuestreo=0` -> `DisponibilidadPeriodo=None`
    - _Requirements: 4.8, 4.9_

- [ ] 7. Motor de eventos de falla
  - [ ] 7.1 Implementar MotorEventosFalla en `src/fault_events/fault_events_engine.py`
    - Dataclass `EventoFalla(id_corrida, numero_pcs, marca_tiempo_inicio, marca_tiempo_fin, duracion_horas, codigo_falla, descripcion_falla, descripcion_falla_fallback, promedio_baterias_involucradas, horas_rack_indisponibles)`
    - Método `detectar_eventos(df_enriquecido, config)`: implementar el algoritmo exacto del VBA transcrito en el design.md
    - Para cada PCS: iterar en orden temporal ASC; detectar inicio/fin de evento basado en `ModulosDisponibles < baterias_por_pcs`
    - `MarcaTiempoInicio = ts_actual - timedelta(minutes=minutos_muestreo)` (semántica exacta del VBA)
    - Aplicar árbol de decisión de descripción con fallback (4 casos según el design.md)
    - Marcar `DescripcionFallaFallback=True` cuando se usa descripción del intervalo anterior
    - Calcular `DuracionHoras`, `PromedioBateriasInvolucradas`, `HorasRackIndisponibles`
    - Implementar `_extraer_codigo_falla(descripcion)`: primer token antes del espacio, o `"F" + descripcion` si no hay espacio
    - Cerrar eventos abiertos al final del período con `marca_fin = fin_periodo + timedelta(days=1)`
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7, 5.8, 5.9, 5.10_

  - [ ]* 7.2 Property test: eventos cubren exactamente las sub-secuencias con ModulosDisponibles < 4
    - **Property 6: Los eventos de falla cubren exactamente las sub-secuencias con ModulosDisponibles < 4**
    - **Validates: Requirements 5.1, 5.3**
    - Generar secuencias aleatorias de `ModulosDisponibles` (0-4 + None); verificar cobertura exacta
    - Tag: `# Feature: etl-arena-availability, Property 6`

  - [ ]* 7.3 Property test: lógica de fallback de descripción es determinista y completa
    - **Property 7: La lógica de fallback de descripción es determinista y completa**
    - **Validates: Requirements 5.4, 5.5**
    - Generar pares `(descripcion_actual, descripcion_anterior)` con valores `"NO FAULTS"`, `""` y cadenas arbitrarias
    - Verificar resultado del árbol de decisión y que `DescripcionFallaFallback` es correcto
    - Tag: `# Feature: etl-arena-availability, Property 7`

  - [ ]* 7.4 Tests unitarios del motor de eventos
    - Test de ejemplo: los cuatro casos del árbol de decisión de descripción con valores concretos
    - Test de ejemplo: evento abierto al final del período se cierra correctamente
    - Test de ejemplo: dos eventos separados en el mismo PCS
    - Test de ejemplo: `_extraer_codigo_falla` con descripción con espacio y sin espacio
    - _Requirements: 5.4, 5.6_

- [ ] 8. Checkpoint — Motores de cálculo completos
  - Ejecutar todos los tests unitarios y de propiedad de los módulos 6 y 7.
  - Verificar con datos de muestra conocidos que `BloquesMuestreo`, `BloquesRacksIndisponibles` y la lista de eventos coinciden con los valores del Excel.
  - Preguntar al usuario si hay ajustes antes de continuar con las agregaciones.

- [ ] 9. Módulo de agregaciones
  - [ ] 9.1 Implementar AgregacionDiaria en `src/aggregation/daily.py`
    - Método `calcular(resultados_muestra, config)`: agrupar `ImpactoRackPonderado` por día
    - Calcular `BloquesRacksIndisponiblesAcumulados` como suma acumulada desde el primer día
    - Fórmula: `Disponibilidad = 1 - acumulado / (TotalRacks * 24 * 60 * dia_N / MinutosMuestreo)` con manejo de denominador cero
    - Calcular `Variacion = Disponibilidad_N - Disponibilidad_(N-1)` con manejo de None
    - Retornar DataFrame con campos de `daily_availability`
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5_

  - [ ] 9.2 Implementar AgregacionAnual en `src/aggregation/annual.py`
    - Método `calcular(resultado, config, mes, anio)`: calcular KPI mensual y acumulado anual
    - Fórmula anual (conservada del Excel para paridad): `1 - BloquesRacksIndisponibles / (TotalRacks * 365 * 24 * 4)`
    - Inicializar acumulación desde `config.fecha_inicio_proyecto` (2026-04-08)
    - Retornar dict/registro con campos de `annual_availability`
    - _Requirements: 6.6, 6.7, 6.8_

  - [ ]* 9.3 Property test: disponibilidad diaria acumulada es consistente con rack blocks
    - **Property 8: La disponibilidad diaria acumulada es consistente con la acumulación de rack blocks**
    - **Validates: Requirements 6.3, 6.4**
    - Generar secuencias de `BloquesRacksIndisponiblesAcumulados` con denominador no nulo
    - Verificar que `0 <= Disponibilidad <= 1` y que la fórmula es exacta
    - Tag: `# Feature: etl-arena-availability, Property 8`

  - [ ]* 9.4 Tests unitarios de agregaciones
    - Test de ejemplo: 5 días conocidos; verificar `Disponibilidad` diaria y `Variacion`
    - Test de ejemplo: `dia_N=0` -> `Disponibilidad=None`
    - Test de ejemplo: fórmula anual con `BloquesRacksIndisponibles` y `TotalRacks` conocidos
    - _Requirements: 6.3, 6.6_

- [ ] 10. Módulo de persistencia
  - [ ] 10.1 Implementar ServicioPersistencia en `src/persistence/persistence_service.py`
    - Usar `SQLAlchemy Core` con driver `pyodbc`; leer connection string de variable de entorno `ETL_ARENA_DB_URL`
    - Métodos: `iniciar_corrida`, `completar_corrida`, `fallar_corrida`, `guardar_raw_pcs`, `guardar_actividad_planta`, `guardar_resultados_muestra`, `guardar_resultado_corrida`, `guardar_eventos_falla`, `guardar_diario`, `guardar_anual`
    - Añadir método `guardar_detenciones(id_corrida, id_proyecto, eventos_falla)`: convierte cada `EventoFalla` en un registro `detencion`, resuelve `IdTipoDetencion` buscando en `tipo_detencion` por `CodigoFalla`, calcula `DuracionSegundos = DATEDIFF(seconds, FechaInicio, FechaTermino)`, inserta en bulk
    - Inserciones en bulk con `executemany`; nunca `DELETE` ni `UPDATE` sobre registros de corridas anteriores
    - En caso de error: hacer rollback, llamar `fallar_corrida` antes de relanzar la excepción
    - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5, 7.6, 7.7_

  - [ ] 10.2 Implementar gestión de IdCorrida en `src/persistence/run_manager.py`
    - Función `generar_id_corrida()`: genera UUID v4 como string
    - Función `crear_corrida(config)`: llama `iniciar_corrida` y retorna el `IdCorrida`
    - _Requirements: 7.2, 9.2_

  - [ ]* 10.3 Tests de integración de persistencia
    - Test de integración: dos corridas con `IdCorrida` distintos; verificar que ambos conjuntos de registros coexisten (append-only)
    - Test de integración: simular fallo en `guardar_resultados_muestra`; verificar que `etl_run.Estado = "failed"` y no quedan datos parciales huérfanos
    - _Requirements: 7.1, 7.7_

- [ ] 11. Módulo de reconciliación
  - [ ] 11.1 Implementar ServicioReconciliacion en `src/reconciliation/reconciliation_service.py`
    - Dataclasses `ResultadoNivel(nivel, nombre, aprobado, discrepancias)` y `ReporteReconciliacion(id_corrida_python, archivo_excel_origen, tolerancia, niveles, aprobado_global)`
    - Método `reconciliar(id_corrida_python, ruta_excel, tolerancia=1e-9)`: ejecutar los 5 niveles de comparación
    - _Requirements: 8.1_

  - [ ] 11.2 Implementar los 5 niveles de comparación en `src/reconciliation/levels.py`
    - `comparar_nivel_1_input(corrida_python, df_excel)`: filas, timestamps, PCS count, frecuencia, rango de fechas
    - `comparar_nivel_2_muestras(muestras_python, muestras_excel, tolerancia)`: por cada `(MarcaTiempoMuestra, NumeroPCS)` comparar `ModulosDisponibles`, `BateriasIndisponibles`, factores, `BateriasIndisponiblesPonderadas`
    - `comparar_nivel_3_acumulados(resultado_python, excel_bloques_muestreo, excel_bloques_racks_indisp, tolerancia)`: BloquesMuestreo (C12) y BloquesRacksIndisponibles (C14)
    - `comparar_nivel_4_kpi(resultado_python, kpi_excel, tolerancia)`: DisponibilidadPeriodo (C16), DisponibilidadAnualAcumulada (C19), KPI diario, KPI mensual
    - `comparar_nivel_5_eventos(eventos_python, eventos_excel, tolerancia, max_ordenado=167)`: campo por campo; advertencia si supera 167 eventos
    - _Requirements: 8.2, 8.3, 8.4, 8.5, 8.6, 8.7, 8.8_

  - [ ]* 11.3 Tests de reconciliación
    - Test de ejemplo: reconciliación con datos idénticos -> reporte `aprobado_global=True`
    - Test de ejemplo: reconciliación con discrepancia conocida en `BloquesRacksIndisponibles` -> reporte indica fail en nivel 3
    - Test de ejemplo: más de 167 eventos -> advertencia registrada en nivel 5
    - _Requirements: 8.7, 8.8, 8.9_

- [ ] 12. Checkpoint — Pipeline completo
  - Ejecutar todos los tests unitarios, de propiedad e integración.
  - Verificar que todos los tests pasan antes de continuar.
  - Preguntar al usuario si hay ajustes antes del orquestador.

- [ ] 13. Pipeline orquestador
  - [ ] 13.1 Implementar script principal `ejecutar_etl.py` en la raíz del proyecto
    - Parsear argumentos CLI: `--excel-path`, `--period-start`, `--period-end`, `--solo-tiempo-operacional`, `--aplicar-evento-excusable`, `--version-algoritmo`
    - Orquestar las fases en orden: `config -> ingesta -> staging -> normalizacion -> enriquecimiento -> disponibilidad -> eventos_falla -> agregacion -> persistencia -> (opcional: reconciliacion)`
    - Registrar inicio y fin de cada fase en el log con timestamp
    - En caso de error: llamar `servicio_persistencia.fallar_corrida` antes de terminar con exit code 1
    - _Requirements: 7.2, 9.1, 9.2, 9.6_

  - [ ] 13.2 Implementar reporte de calidad de datos en `src/reporting/reporte_calidad.py`
    - Función `generar_reporte_calidad(id_corrida, anomalias, resultados_muestra, eventos_falla)`: producir resumen con: cantidad de intervalos `ModulosDisponiblesNulo`, cantidad de eventos `DescripcionFallaFallback`, anomalías de timestamp, PCS o columnas faltantes
    - Imprimir el reporte en consola y persistir en `etl_run.MensajeError` (o en un campo dedicado si se agrega)
    - _Requirements: 9.6_

  - [ ]* 13.3 Tests de integración end-to-end
    - Test de integración: ejecutar el pipeline completo sobre un subconjunto de datos de muestra del Excel real (extraídos manualmente)
    - Verificar que los KPI producidos coinciden con los valores conocidos del Excel con tolerancia `1e-4`
    - Verificar que todos los registros tienen el mismo `IdCorrida` y que `etl_run.Estado = "success"`
    - _Requirements: 9.1, 10.3_

- [ ] 13.4 Implementar golden tests en `tests/golden/`
    - Crear `tests/golden/conftest.py` con fixture `datos_golden(id_mes)` que carga el JSON correspondiente de `tests/golden/data/` usando el índice `golden_index.json`
    - Crear `tests/golden/test_golden_runner.py`: test parametrizado que itera sobre todos los meses en `golden_index.json` con status="verified" y verifica: `BloquesMuestreo` exacto (int), `BloquesRacksIndisponibles`/`DisponibilidadPeriodo`/`DisponibilidadAnualAcumulada` con tolerancia 1e-6, `Disponibilidad` diaria con tolerancia 1e-6 por dia, total rack-hours con tolerancia 1e-3
    - Crear `tests/golden/test_gt_july_2026.py`: test especifico para julio 2026 usando `tests/golden/data/golden_2026_07_july.json` — verifica `BloquesMuestreo=2976`, `BloquesRacksIndisponibles=432611.2119999999`, `DisponibilidadPeriodo=0.9503529130126623`, `DisponibilidadAnualAcumulada=0.9957833980914864`, los 31 dias de `Disponibilidad` diaria, y los primeros 10 eventos del PCS 1
    - Crear `tests/golden/test_gt_august_2026.py`: test especifico para agosto 2026 usando `tests/golden/data/golden_2026_08_august.json` — verifica `BloquesMuestreo=2976`, `BloquesRacksIndisponibles=527396.0440000003`, `DisponibilidadPeriodo=0.9394752689090135`, `DisponibilidadAnualAcumulada=0.9948595433867929`, los 31 dias de `Disponibilidad` diaria, y los primeros 10 eventos del PCS 1
    - Crear `tests/golden/test_gt_september_2026.py`: test especifico para septiembre 2026 (periodo parcial 01-21 sep) usando `tests/golden/data/golden_2026_09_september.json` — verifica `BloquesMuestreo=1975`, `BloquesRacksIndisponibles=104134.2920000001`, `DisponibilidadPeriodo=0.9819924099052362`, `DisponibilidadAnualAcumulada=0.9989850173961998`, los 21 dias con datos reales y los 9 dias con `Disponibilidad=0`, y los primeros 10 eventos del PCS 1
    - Crear `tests/golden/test_gt_decimal_modules.py`: verificar que el PCS 04 en el primer bloque (2026-04-08 00:15, modules=3.234...) produce `BateriasIndisponibles` correcto (caso de valor decimal en NUMBER_OF_MODULES — documentado en GT-7 del design.md)
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
- Todas las constantes de negocio (`61`, `4`, `12`, `15`) deben provenir de `ConfiguracionCalculo`, nunca hardcodeadas en los motores
- SQL Server es append-only por `IdCorrida`; nunca ejecutar `DELETE` ni `UPDATE` sobre datos históricos
- La tolerancia numérica para comparaciones en reconciliación es `1e-9`
- La fórmula anual con `365 * 24 * 4` se conserva sin modificar durante la fase de paridad
- Para la fase de paridad, `version_algoritmo = "availability-v1-excel-parity"`

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

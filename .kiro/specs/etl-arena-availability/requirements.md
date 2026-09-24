# Requirements Document

## Introduction

Este proyecto reemplaza progresivamente el cálculo de disponibilidad del activo **Arena BESS** (almacenamiento de energía), actualmente implementado en un libro Excel con macros VBA, por un proceso reproducible y auditable en **Python + SQL Server**.

El objetivo de la primera fase es obtener **paridad exacta** con el Excel: reproducir los mismos valores de `BloquesMuestreo` (C12), `BloquesRacksIndisponibles` (C14), `DisponibilidadPeriodo` (C16) y `DisponibilidadAnualAcumulada` (C19), la misma lista de eventos de `ListOfFaults`, y los mismos KPI diarios y acumulados anuales, usando los mismos datos de entrada.

**Parámetros del activo (Arena BESS):**

| Parámetro | Valor |
|---|---|
| Fecha de inicio de operación | 08/Abril/2026 |
| Total PCS | 61 |
| Módulos BEC por PCS | 4 |
| Racks por BAC | 12 |
| Total Racks | 2.928 |
| Frecuencia de muestreo | 15 minutos |
| Versión del algoritmo | `availability-v1-excel-parity` |

---

## Glossary

- **PCS**: Power Conversion System. Unidad de conversión de potencia. El activo tiene 61 unidades.
- **BEC**: Battery Energy Controller. Módulo de batería. Cada PCS tiene 4 módulos BEC.
- **Rack**: Unidad mínima de batería. Cada BAC contiene 12 racks. Total: 2.928 racks.
- **BAC**: Battery Array Container. Contenedor de baterías que agrupa 12 racks.
- **SOC**: State of Charge. Porcentaje de carga de la batería.
- **POI**: Point of Interconnection. Punto de interconexión a la red eléctrica.
- **DST**: Daylight Saving Time. Cambio de horario de verano/invierno de Chile. Aplica en septiembre 2026.
- **KPI**: Key Performance Indicator. Indicador clave de rendimiento.
- **MotorETL**: El sistema Python que reproduce la lógica VBA del Excel.
- **ModuloIngesta**: Componente responsable de leer los archivos origen (Excel o exportaciones CSV).
- **ModuloStaging**: Componente que almacena los datos crudos sin transformar en SQL Server.
- **ModuloNormalizacion**: Componente que convierte el formato ancho (244 columnas) a formato largo (1 fila por PCS×timestamp).
- **ModuloEnriquecimiento**: Componente que une los datos PCS con los factores de PlantActivity.
- **MotorDisponibilidad**: Componente que reproduce `cmdCalcAvailability` — calcula `BloquesMuestreo` (C12), `BloquesRacksIndisponibles` (C14) y el KPI del período.
- **MotorEventosFalla**: Componente que reproduce `mcoCreateList` — consolida intervalos consecutivos en eventos de falla discretos.
- **ModuloAgregacion**: Componente que calcula KPI diario, mensual y acumulado anual.
- **ModuloPersistencia**: Componente que escribe resultados en SQL Server de forma append-only.
- **ModuloReconciliacion**: Componente que compara una corrida Python contra una corrida Excel en 5 niveles.
- **IdCorrida**: Identificador único UUID de una ejecución del MotorETL. Clave de auditoría en todas las tablas.
- **VersionAlgoritmo**: Cadena que identifica la versión del algoritmo usada en una corrida (ej. `availability-v1-excel-parity`).
- **BloquesMuestreo**: Contador de bloques de 15 minutos procesados en el período. Equivale a `C12` en el Excel y a la columna `BloquesMuestreo` en SQL.
- **BloquesRacksIndisponibles**: Acumulado de `(racks indisponibles) × bloques`, opcionalmente ponderado por factores excusable y operacional. Equivale a `C14` en el Excel y a la columna `BloquesRacksIndisponibles` en SQL.
- **DisponibilidadPeriodo**: Disponibilidad del período = `1 - BloquesRacksIndisponibles / (TotalRacks × BloquesMuestreo)`. Equivale a `C16` en el Excel.
- **DisponibilidadAnualAcumulada**: Disponibilidad acumulada anual = `1 - BloquesRacksIndisponibles / (total_racks × 365 × 24 × 4)`. Equivale a `C19` en el Excel.
- **NUMBER_OF_MODULES**: Campo `Arena - PCS XX - POWERELECTRONICS HEM-k NUMBER OF MODULES` en `RawData-PCS`. Valor numérico 0–4. Valor < 4 indica indisponibilidad. Vacío se trata como 4 (disponible) con flag de auditoría.
- **ModulosDisponiblesNulo**: Flag booleano. `True` cuando `NUMBER_OF_MODULES` estaba vacío en el origen.
- **DescripcionFallaFallback**: Flag booleano. `True` cuando la descripción del evento fue tomada del intervalo temporal anterior.
- **FactorExcusable**: Factor del campo `Excused Event` de PlantActivity (columna D). Valor 0 o 1.
- **FactorOperacional**: Factor de actividad operacional de PlantActivity (columna C). Valor 0 (inactivo) o 1 (activo).
- **RawData-PCS**: Hoja del Excel con los registros crudos de estado y falla por intervalo de 15 min y por PCS (244 columnas de datos PCS).
- **PlantActivity**: Hoja del Excel con el factor operacional (columna C) y el indicador de evento excusable (columna D) por intervalo de 15 min.
- **Proyecto**: Entidad maestra que representa un activo BESS. Contiene parámetros de configuración como `NumPCS`, `NumBateriasPorPCS`, `NumRacksPorBAC`. Arena BESS es el proyecto `IdProyecto=1`, actualmente en ejecución. Copiapó A, Luz del Norte y María Elena están por implementar.
- **Detencion**: Evento de indisponibilidad discreto de un PCS, equivalente a una fila de `ListOfFaults`. Tiene `FechaInicio`, `FechaTermino`, `DuracionSegundos`, tipo de detención, flags de calidad y campos de estado para revisión operacional.
- **TipoDetencion**: Catálogo de códigos de falla extraído de la hoja `PCS-Fault` del Excel. 68 códigos (F0–F118). Identifica la causa técnica de cada detención.
- **EstadoRevision**: Campo de workflow en `detencion`. Valores: pendiente (por defecto), revisado, excluido. Permite marcar detenciones para análisis posterior.

---

## Requirements

### Requirement 1: Ingesta y staging

**User Story:** Como ingeniero de datos, quiero leer los archivos origen y conservar una copia inmutable en staging, para que cualquier reprocesamiento pueda partir del mismo dato original y sea completamente trazable.

#### Acceptance Criteria

1. WHEN el ModuloIngesta recibe la ruta de un archivo Excel válido, THE ModuloIngesta SHALL leer la hoja `RawData-PCS` comenzando en la fila 2 y conservando todas las columnas presentes.
2. WHEN el ModuloIngesta recibe la ruta de un archivo Excel válido, THE ModuloIngesta SHALL leer la hoja `PlantActivity` comenzando en la fila 2 y conservando todas las columnas presentes.
3. WHEN el ModuloIngesta termina la lectura, THE ModuloStaging SHALL persistir los registros exactamente como llegaron, sin ninguna transformación, asociados al `IdCorrida` de la corrida actual.
4. IF el archivo fuente no existe o no es legible, THEN THE ModuloIngesta SHALL registrar el error en `etl_run.MensajeError` y marcar la corrida con `Estado = "failed"`.
5. WHEN el ModuloIngesta lee `RawData-PCS`, THE ModuloIngesta SHALL detectar y reportar como advertencias: huecos en la secuencia de timestamps, timestamps duplicados, timestamps fuera de orden ascendente, e intervalos con frecuencia distinta a la configurada.
6. WHEN el ModuloIngesta encuentra una celda vacía en la columna A (timestamp) de `RawData-PCS`, THE ModuloIngesta SHALL registrar la posición del hueco como advertencia y continuar el procesamiento, en lugar de detenerlo silenciosamente como hace el VBA.
7. THE ModuloStaging SHALL conservar para cada registro de `RawData-PCS` tanto el valor serial numérico de fecha Excel (`SerialFechaExcelOrigen`) como el timestamp interpretado en hora local de Chile (`MarcaTiempoLocalOrigen`).
8. WHEN el ModuloIngesta detecta anomalías de calidad de datos, THE MotorETL SHALL incluir un resumen de esas anomalías en el registro `etl_run` de la corrida sin abortar el procesamiento.

---

### Requirement 2: Normalización

**User Story:** Como desarrollador del motor de disponibilidad, quiero los datos PCS en formato largo (1 fila por PCS×timestamp), para que el motor pueda procesar cada combinación sin depender del índice de columna frágil del Excel.

#### Acceptance Criteria

1. WHEN el ModuloNormalizacion recibe los datos crudos de `RawData-PCS`, THE ModuloNormalizacion SHALL transformar el formato ancho de 244 columnas (61 PCS × 4 columnas) en formato largo con una fila por `(MarcaTiempoMuestra, NumeroPCS)`.
2. THE ModuloNormalizacion SHALL extraer para cada PCS los cuatro campos: `CodigoFallaRaw` (CURRENT FAULT), `EstadoRaw` (CURRENT STATUS), `AdvertenciaRaw` (CURRENT WARNING) y `ModulosDisponibles` (NUMBER OF MODULES), usando el patrón de nombre `Arena - PCS XX - POWERELECTRONICS ...` donde `XX` va de `01` a `61`.
3. WHEN el ModuloNormalizacion encuentra que el campo `NUMBER_OF_MODULES` está vacío para un `(MarcaTiempoMuestra, NumeroPCS)`, THE ModuloNormalizacion SHALL asignar `ModulosDisponibles = 4` y marcar `ModulosDisponiblesNulo = True`.
4. WHEN el ModuloNormalizacion encuentra que el campo `NUMBER_OF_MODULES` contiene un valor numérico, THE ModuloNormalizacion SHALL asignar ese valor a `ModulosDisponibles` y marcar `ModulosDisponiblesNulo = False`.
5. THE ModuloNormalizacion SHALL detectar y reportar: PCS faltantes respecto al total configurado, columnas faltantes por PCS, columnas con nombres inesperados, y columnas en posición desplazada.
6. WHEN la normalización finaliza, THE ModuloNormalizacion SHALL producir un conjunto de datos normalizado ordenado por `(MarcaTiempoMuestra ASC, NumeroPCS ASC)`.

---

### Requirement 3: Enriquecimiento

**User Story:** Como desarrollador del motor de disponibilidad, quiero unir los datos PCS normalizados con los factores de PlantActivity, para que el motor pueda aplicar los ponderadores operacional y excusable correctamente en cada intervalo.

#### Acceptance Criteria

1. WHEN el ModuloEnriquecimiento recibe los datos normalizados de PCS y los datos de PlantActivity, THE ModuloEnriquecimiento SHALL unir ambas fuentes por `MarcaTiempoMuestra`.
2. THE ModuloEnriquecimiento SHALL asignar a cada registro de `raw_pcs_sample` el `FactorOperacional` (columna C de PlantActivity) y el `FactorExcusable` (columna D de PlantActivity) correspondientes al mismo timestamp.
3. IF un `MarcaTiempoMuestra` de `raw_pcs_sample` no tiene correspondencia en `plant_activity_sample`, THEN THE ModuloEnriquecimiento SHALL registrar el hueco como advertencia y asignar valores por defecto: `FactorOperacional = 1`, `FactorExcusable = 1`.
4. WHEN el enriquecimiento finaliza, THE ModuloEnriquecimiento SHALL producir un conjunto de datos enriquecido con todos los campos necesarios para el MotorDisponibilidad: `ModulosDisponibles`, `ModulosDisponiblesNulo`, `FactorOperacional` y `FactorExcusable`.

---

### Requirement 4: Motor de disponibilidad

**User Story:** Como analista de disponibilidad, quiero que el motor Python calcule `BloquesMuestreo` (C12), `BloquesRacksIndisponibles` (C14) y el KPI del período reproduciéndose exactamente igual que la macro `cmdCalcAvailability`, para poder comparar los resultados antes de retirar el Excel como fuente oficial.

#### Acceptance Criteria

1. WHEN el MotorDisponibilidad procesa un período, THE MotorDisponibilidad SHALL incrementar `BloquesMuestreo` en 1 por cada intervalo de 15 minutos válido en el rango de fechas, independientemente del número de PCS.
2. WHEN el MotorDisponibilidad evalúa `ModulosDisponibles` para un `(MarcaTiempoMuestra, NumeroPCS)`, THE MotorDisponibilidad SHALL considerar ese PCS como indisponible únicamente si `ModulosDisponibles` es numérico y `ModulosDisponibles < 4`.
3. WHEN el MotorDisponibilidad identifica un PCS indisponible en un intervalo, THE MotorDisponibilidad SHALL calcular `BateriasIndisponibles = baterias_por_pcs - ModulosDisponibles`.
4. WHERE el parámetro de corrida `aplicar_evento_excusable = "Yes"`, THE MotorDisponibilidad SHALL calcular `BateriasIndisponiblesPonderadas = BateriasIndisponibles * FactorExcusable`.
5. WHERE el parámetro de corrida `aplicar_evento_excusable = "No"`, THE MotorDisponibilidad SHALL calcular `BateriasIndisponiblesPonderadas = BateriasIndisponibles`.
6. WHERE el parámetro de corrida `solo_tiempo_operacional = "Yes"`, THE MotorDisponibilidad SHALL acumular `BloquesRacksIndisponibles += racks_por_pcs * BateriasIndisponiblesPonderadas * FactorOperacional`.
7. WHERE el parámetro de corrida `solo_tiempo_operacional = "No"`, THE MotorDisponibilidad SHALL acumular `BloquesRacksIndisponibles += racks_por_pcs * BateriasIndisponiblesPonderadas`.
8. WHEN el MotorDisponibilidad ha procesado todos los intervalos del período, THE MotorDisponibilidad SHALL calcular `DisponibilidadPeriodo = 1 - BloquesRacksIndisponibles / (TotalRacks * BloquesMuestreo)`.
9. IF `BloquesMuestreo = 0` o `TotalRacks = 0`, THEN THE MotorDisponibilidad SHALL producir `DisponibilidadPeriodo = None` (equivalente al `IFERROR(..., "N/A")` del Excel).
10. THE MotorDisponibilidad SHALL usar el factor `racks_por_pcs` proveniente de la configuración de la corrida (valor `12` durante la fase de paridad), sin asumir ese valor como constante del código.
11. THE MotorDisponibilidad SHALL producir un registro en `availability_sample_result` por cada `(IdCorrida, MarcaTiempoMuestra, NumeroPCS)` incluyendo: `ModulosDisponibles`, `ModulosDisponiblesNulo`, `BateriasIndisponibles`, `FactorExcusable`, `FactorOperacional`, `BateriasIndisponiblesPonderadas` y `ImpactoRackPonderado`.

---

### Requirement 5: Motor de eventos de falla

**User Story:** Como analista de disponibilidad, quiero que el motor Python genere la lista de eventos de falla reproduciéndose exactamente igual que la macro `mcoCreateList`, incluyendo la lógica de fallback en la descripción, para poder reconciliar la lista con el Excel evento por evento.

#### Acceptance Criteria

1. WHEN el MotorEventosFalla procesa los datos normalizados, THE MotorEventosFalla SHALL recorrer cada PCS por separado en orden temporal ascendente y agrupar secuencias consecutivas de intervalos donde `ModulosDisponibles < 4` como un único evento de falla.
2. WHEN el MotorEventosFalla detecta el inicio de un evento de falla, THE MotorEventosFalla SHALL calcular `MarcaTiempoInicio = marca_tiempo_actual - minutos_muestreo / (24 * 60)` expresado como fracción de día (conservando la semántica exacta del VBA).
3. WHEN el MotorEventosFalla detecta el fin de un evento de falla, THE MotorEventosFalla SHALL cerrar el evento cuando el siguiente intervalo tiene `ModulosDisponibles = 4`, está vacío, o su timestamp supera `fin_periodo + 1 día`.
4. WHEN el MotorEventosFalla determina la descripción de un evento de falla, THE MotorEventosFalla SHALL aplicar la siguiente lógica en orden:
   - Si la descripción del intervalo actual es distinta de `"NO FAULTS"` y no está vacía: usar la descripción actual.
   - Si la descripción actual es `"NO FAULTS"` y el intervalo inmediatamente anterior tiene una descripción distinta de `"NO FAULTS"`: usar la descripción del intervalo anterior.
   - Si la descripción actual es `"NO FAULTS"` y el intervalo anterior también es `"NO FAULTS"` o no existe: asignar `"F13 NO MODULES"`.
   - Si la descripción resultante está vacía: asignar `"F1 Watchdog"`.
5. WHEN el MotorEventosFalla aplica la descripción del intervalo anterior, THE MotorEventosFalla SHALL marcar el evento con `DescripcionFallaFallback = True`.
6. WHEN el MotorEventosFalla calcula el código de falla, THE MotorEventosFalla SHALL extraer el primer token antes del primer espacio en la descripción final; IF la descripción no contiene espacio, THEN THE MotorEventosFalla SHALL usar `"F" + descripcion_completa`.
7. WHEN el MotorEventosFalla calcula la duración de un evento, THE MotorEventosFalla SHALL usar `DuracionHoras = 24 * (MarcaTiempoFin - MarcaTiempoInicio)`.
8. WHEN el MotorEventosFalla calcula el impacto de un evento, THE MotorEventosFalla SHALL usar `HorasRackIndisponibles = racks_por_pcs * DuracionHoras * (sumablocks / numBlock)`, donde `racks_por_pcs = 12` proviene de la configuración.
9. WHERE el parámetro `aplicar_evento_excusable = "Yes"`, THE MotorEventosFalla SHALL ponderar cada bloque del acumulador `sumablocks` con el `FactorExcusable` del timestamp correspondiente.
10. THE MotorEventosFalla SHALL producir registros en la tabla `fault_event` con todos los campos del modelo de datos: `IdCorrida`, `NumeroPCS`, `MarcaTiempoInicio`, `MarcaTiempoFin`, `DuracionHoras`, `CodigoFalla`, `DescripcionFalla`, `DescripcionFallaFallback`, `PromedioBateriasInvolucradas` y `HorasRackIndisponibles`.

---

### Requirement 6: Agregaciones

**User Story:** Como analista de disponibilidad, quiero los KPI diarios, mensuales y el acumulado anual calculados con la misma lógica que el Excel (macros `mcoDailyAvailability` y hoja `Annual_AVA`), para comparar los resultados período a período.

#### Acceptance Criteria

1. WHEN el ModuloAgregacion calcula el KPI diario para un día, THE ModuloAgregacion SHALL sumar todos los valores `ImpactoRackPonderado` de `availability_sample_result` correspondientes a ese día e `IdCorrida`.
2. WHEN el ModuloAgregacion acumula el KPI diario, THE ModuloAgregacion SHALL calcular `BloquesRacksIndisponiblesAcumulados` como la suma desde el primer día del período hasta el día actual, inclusive.
3. WHEN el ModuloAgregacion calcula la disponibilidad diaria acumulada para el día N del mes, THE ModuloAgregacion SHALL usar: `Disponibilidad = 1 - BloquesRacksIndisponiblesAcumulados / (TotalRacks * 24 * 60 * N / MinutosMuestreo)`.
4. IF el denominador de la fórmula de disponibilidad diaria es cero, THEN THE ModuloAgregacion SHALL producir `Disponibilidad = None` para ese día.
5. WHEN el ModuloAgregacion calcula la variación diaria, THE ModuloAgregacion SHALL calcular `Variacion = Disponibilidad_dia_N - Disponibilidad_dia_N_menos_1` con manejo de error equivalente a `IFERROR`.
6. WHEN el ModuloAgregacion calcula la disponibilidad acumulada anual para una corrida, THE ModuloAgregacion SHALL usar la fórmula `DisponibilidadAnualAcumulada = 1 - BloquesRacksIndisponibles / (TotalRacks * (365 * 24 * 4))`, conservando el supuesto de 365 días y 4 bloques/hora del Excel durante la fase de paridad.
7. THE ModuloAgregacion SHALL producir registros en `annual_availability` con: año, mes, días en el mes, bloques de muestreo, racks indisponibles, disponibilidad mensual, acumulados históricos y disponibilidad acumulada.
8. THE ModuloAgregacion SHALL inicializar la acumulación anual desde el **08/Abril/2026** (fecha de inicio de operación de Arena BESS), sin acumular períodos anteriores a esa fecha.

---

### Requirement 7: Persistencia SQL Server

**User Story:** Como administrador del sistema, quiero que todos los resultados y datos intermedios se almacenen en SQL Server de forma append-only por `IdCorrida`, para que ninguna corrida nueva destruya los resultados históricos y cada corrida sea completamente auditable.

#### Acceptance Criteria

1. THE ModuloPersistencia SHALL operar en modo append-only: nunca eliminar ni sobreescribir registros existentes en ninguna tabla.
2. THE ModuloPersistencia SHALL crear un registro en `etl_run` al inicio de cada corrida con `Estado = "running"` y actualizarlo con `Estado = "success"` o `Estado = "failed"` al finalizar.
3. THE ModuloPersistencia SHALL insertar todos los registros de staging, resultados intermedios y KPI bajo el mismo `IdCorrida` de la corrida activa.
4. THE ModuloPersistencia SHALL persistir las diez tablas del modelo de datos: `proyecto`, `tipo_detencion`, `etl_run`, `raw_pcs_sample`, `plant_activity_sample`, `availability_sample_result`, `availability_run_result`, `detencion`, `daily_availability` y `annual_availability`.
5. WHEN el ModuloPersistencia inserta registros en `raw_pcs_sample`, THE ModuloPersistencia SHALL conservar el número de fila origen (`NumeroFilaOrigen`) y el nombre de las columnas origen (`ColumnasOrigen`) para trazabilidad.
6. THE ModuloPersistencia SHALL almacenar en `etl_run` los parámetros completos de la corrida: `TotalPCS`, `BateriasPorPCS`, `RacksPorPCS`, `TotalRacks`, `MinutosMuestreo`, `SoloTiempoOperacional`, `AplicarEventoExcusable` y `VersionAlgoritmo`.
7. IF la conexión a SQL Server falla durante la inserción, THEN THE ModuloPersistencia SHALL registrar el error, marcar la corrida con `Estado = "failed"` y no dejar datos parciales sin el correspondiente `etl_run` de error.

---

### Requirement 8: Reconciliación

**User Story:** Como ingeniero de datos, quiero comparar automáticamente una corrida Python contra una corrida Excel en cinco niveles de detalle, para poder declarar la paridad antes de retirar el Excel como fuente oficial.

#### Acceptance Criteria

1. THE ModuloReconciliacion SHALL comparar corridas en cinco niveles: (1) input, (2) muestra individual, (3) acumulados, (4) KPI y (5) eventos.
2. WHEN el ModuloReconciliacion ejecuta la comparación de Nivel 1 (input), THE ModuloReconciliacion SHALL verificar que coincidan: cantidad total de filas, rango de timestamps, cantidad de PCS, cantidad de registros por PCS, frecuencia de muestreo, fecha inicial y fecha final.
3. WHEN el ModuloReconciliacion ejecuta la comparación de Nivel 2 (muestra individual), THE ModuloReconciliacion SHALL comparar por cada `(MarcaTiempoMuestra, NumeroPCS)`: `ModulosDisponibles`, `BateriasIndisponibles`, `FactorExcusable`, `FactorOperacional` y `BateriasIndisponiblesPonderadas`.
4. WHEN el ModuloReconciliacion ejecuta la comparación de Nivel 3 (acumulados), THE ModuloReconciliacion SHALL comparar `BloquesMuestreo` (C12) y `BloquesRacksIndisponibles` (C14) entre la corrida Python y los valores del Excel.
5. WHEN el ModuloReconciliacion ejecuta la comparación de Nivel 4 (KPI), THE ModuloReconciliacion SHALL comparar `DisponibilidadPeriodo` (C16) y `DisponibilidadAnualAcumulada` (C19), así como los KPI diarios y mensuales.
6. WHEN el ModuloReconciliacion ejecuta la comparación de Nivel 5 (eventos), THE ModuloReconciliacion SHALL comparar campo por campo cada evento de `ListOfFaults`: `NumeroPCS`, `MarcaTiempoInicio`, `MarcaTiempoFin`, `DuracionHoras`, `CodigoFalla`, `PromedioBateriasInvolucradas` y `HorasRackIndisponibles`.
7. WHEN el ModuloReconciliacion compara valores numéricos de punto flotante, THE ModuloReconciliacion SHALL usar una tolerancia absoluta de `1e-9` para los cálculos internos y reportar cualquier diferencia que supere ese umbral.
8. WHEN el ModuloReconciliacion detecta que el período analizado supera 167 eventos, THE ModuloReconciliacion SHALL documentar que el Excel solo ordena hasta 167 eventos y excluir la comparación de posición para los eventos adicionales.
9. WHEN el ModuloReconciliacion finaliza, THE ModuloReconciliacion SHALL producir un `ReporteReconciliacion` estructurado que indique el resultado (pass/fail) por nivel y el detalle de las discrepancias encontradas.

---

### Requirement 9: Auditoría y trazabilidad

**User Story:** Como auditor del KPI contractual, quiero poder rastrear cualquier valor de disponibilidad hasta su dato de entrada original, para poder responder preguntas sobre por qué se obtuvo un determinado resultado.

#### Acceptance Criteria

1. THE MotorETL SHALL garantizar trazabilidad completa desde el KPI final hasta el dato crudo: `availability_run_result` -> `availability_sample_result` -> `raw_pcs_sample` -> registro del Excel original.
2. THE MotorETL SHALL identificar cada corrida con un `IdCorrida` único y una `VersionAlgoritmo` que permita saber qué versión de la lógica produjo ese resultado.
3. WHEN el MotorETL almacena registros de `raw_pcs_sample`, THE MotorETL SHALL conservar el `NumeroFilaOrigen` de la fila en el Excel origen y los nombres de columna originales en `ColumnasOrigen`.
4. THE MotorETL SHALL identificar y hacer consultables todos los intervalos con `ModulosDisponiblesNulo = True` a lo largo de toda la historia del activo, sin suprimir esos registros.
5. THE MotorETL SHALL identificar y hacer consultables todos los eventos con `DescripcionFallaFallback = True`, reportando la frecuencia de ese caso en cada corrida.
6. THE MotorETL SHALL incluir en cada corrida un reporte de calidad de datos que indique: cantidad de intervalos con `ModulosDisponiblesNulo`, cantidad de eventos con `DescripcionFallaFallback`, anomalías de timestamp detectadas, y PCS o columnas faltantes.

---

### Requirement 10: Configuración y versionado

**User Story:** Como desarrollador, quiero que todos los parámetros de negocio provengan de la configuración de la corrida y no estén hardcodeados, para que el sistema pueda adaptarse a futuros cambios sin modificar el código del motor.

#### Acceptance Criteria

1. THE MotorETL SHALL leer todos los parámetros de negocio desde la `ConfiguracionCalculo` de la corrida: `total_pcs`, `baterias_por_pcs`, `racks_por_pcs`, `minutos_muestreo`, `fecha_inicio_proyecto`, `solo_tiempo_operacional` y `aplicar_evento_excusable`.
2. THE MotorETL SHALL calcular `total_racks = total_pcs * baterias_por_pcs * racks_por_pcs` a partir de los parámetros de configuración, sin usar el valor 2.928 como constante.
3. THE MotorETL SHALL asignar a cada corrida una `VersionAlgoritmo` que permita distinguir la versión de la lógica de cálculo. El valor por defecto para la fase de paridad es `"availability-v1-excel-parity"`.
4. WHEN se modifica una regla de negocio en una versión futura del algoritmo, THE MotorETL SHALL crear una nueva `VersionAlgoritmo` y nunca sobreescribir resultados históricos calculados con versiones anteriores.
5. THE MotorETL SHALL proporcionar valores por defecto para la primera versión de paridad: `nombre_proyecto = "Arena BESS"`, `fecha_inicio_proyecto = "2026-04-08"`, `total_pcs = 61`, `baterias_por_pcs = 4`, `racks_por_pcs = 12`, `minutos_muestreo = 15`.

---

### Requirement 11: Tratamiento de timestamps y DST

**User Story:** Como ingeniero de datos, quiero que el sistema preserve los timestamps originales del Excel y maneje correctamente el cambio de horario de Chile, para que la paridad no se rompa en períodos que atraviesan un cambio de DST.

#### Acceptance Criteria

1. THE ModuloIngesta SHALL conservar para cada registro tanto el número serial de fecha Excel original (`SerialFechaExcelOrigen`) como el timestamp interpretado en hora local de Chile (`MarcaTiempoLocalOrigen`).
2. THE MotorETL SHALL definir explícitamente la zona horaria de negocio como `America/Santiago` para la interpretación de timestamps.
3. THE MotorETL SHALL no convertir timestamps a UTC sin documentar el efecto de esa conversión en el reporte de la corrida.
4. THE ModuloReconciliacion SHALL validar especialmente los períodos que atraviesan cambios de DST, comparando los timestamps de inicio y fin de eventos entre Excel y Python.
5. WHEN el ModuloIngesta lee timestamps de `RawData-PCS`, THE ModuloIngesta SHALL calcular la frecuencia de muestreo efectiva usando `ROUND((ts[i+1] - ts[i]) * 24 * 60, 2)` y reportar cualquier diferencia respecto a la frecuencia configurada.

---

### Requirement 12: Golden Reference: comparacion obligatoria contra valores reales del Excel

**User Story:** Como ingeniero de datos, quiero que cada corrida Python sea validada automaticamente contra los valores reales calculados por el Excel (golden references), para garantizar paridad numerica exacta antes de declarar que el motor Python es equivalente al VBA.

#### Acceptance Criteria

1. THE MotorETL SHALL incluir un conjunto de golden references en `tests/golden/data/`, uno por cada mes con corrida VBA validada, en formato JSON con el esquema definido en `golden_index.json`.
2. WHEN se agrega un nuevo mes al golden index, THE ModuloReconciliacion SHALL comparar automaticamente los resultados de la corrida Python contra ese golden reference usando las tolerancias definidas.
3. THE golden reference de cada mes SHALL contener como minimo: parametros de la corrida (`TotalPCS`, `BateriasPorPCS`, `TotalRacks`, `MinutosMuestreo`, `SoloTiempoOperacional`, `AplicarEventoExcusable`), KPI del periodo (`BloquesMuestreo` (C12), `BloquesRacksIndisponibles` (C14), `DisponibilidadPeriodo` (C16), `DisponibilidadAnualAcumulada` (C19)), disponibilidad diaria completa del mes, snapshot de Annual_AVA, y los primeros 20 eventos de ListOfFaults con el total de rack-hours.
4. WHEN el ModuloReconciliacion compara KPI del periodo (`DisponibilidadPeriodo` (C16), `DisponibilidadAnualAcumulada` (C19)), THE ModuloReconciliacion SHALL usar tolerancia absoluta de **1e-6**.
5. WHEN el ModuloReconciliacion compara `BloquesMuestreo` (C12), THE ModuloReconciliacion SHALL usar igualdad exacta de enteros.
6. WHEN el ModuloReconciliacion compara `BloquesRacksIndisponibles` (C14), THE ModuloReconciliacion SHALL usar tolerancia absoluta de **1e-6**.
7. WHEN el ModuloReconciliacion compara disponibilidad diaria (`Disponibilidad` por dia), THE ModuloReconciliacion SHALL usar tolerancia absoluta de **1e-6**.
8. WHEN el ModuloReconciliacion compara total de rack-hours de ListOfFaults, THE ModuloReconciliacion SHALL usar tolerancia absoluta de **1e-3**.
9. IF cualquier comparacion contra el golden reference falla, THEN THE MotorETL SHALL marcar la corrida con `Estado = "parity_failed"` y producir un reporte detallado con las discrepancias.
10. THE golden reference SHALL ser extraido del Excel usando el script `tests/golden/extract_golden.py`, que usa `openpyxl` con `data_only=True` para leer los valores calculados por la macro VBA, no las formulas.
11. THE golden index en `tests/golden/data/golden_index.json` SHALL ser la fuente de verdad sobre que meses tienen golden reference validado y cuales estan pendientes.

#### Golden References disponibles

| Mes | Archivo | BloquesMuestreo (C12) | DisponibilidadPeriodo (C16) | DisponibilidadAnualAcumulada (C19) | Estado |
|---|---|---|---|---|---|
| Julio 2026  | `golden_2026_07_july.json`   | 2976 | 0.9503529130 | 0.9957833981 | verificado |
| Agosto 2026 | `golden_2026_08_august.json` | 2976 | 0.9394752689 | 0.9948595434 | verificado |
| Septiembre 2026 (parcial 1-21) | `golden_2026_09_september.json` | 1975 | 0.9819924099 | 0.9989850174 | verificado |

#### Tolerancias de comparacion

| Campo | Tolerancia | Justificacion |
|---|---|---|
| BloquesMuestreo (C12) | exacto (int) | Contador de enteros, no puede diferir |
| BloquesRacksIndisponibles (C14) | abs <= 1e-6 | Suma de floats; diferencias por orden de operaciones son aceptables |
| DisponibilidadPeriodo (C16) | abs <= 1e-6 | KPI presentado; precision suficiente para 6 decimales |
| DisponibilidadAnualAcumulada (C19) | abs <= 1e-6 | Idem |
| Disponibilidad diaria | abs <= 1e-6 | Por dia; tolerancia igual al KPI del periodo |
| total rack-hours (ListOfFaults) | abs <= 1e-3 | Suma de muchos floats; margen mayor aceptable |
| DuracionHoras por evento | abs <= 0.01 | Diferencias de 36 segundos son irrelevantes operacionalmente |
| HorasRackIndisponibles por evento | abs <= 0.1 | Derivado de DuracionHoras y PromedioBateriasInvolucradas |

---

### Requirement 13: Modelo de proyectos y detenciones

**User Story:** Como analista de disponibilidad, quiero que la base de datos soporte múltiples proyectos BESS y almacene las detenciones de forma estructurada con su tipo, duración y estado de revisión, para poder consultar y analizar fallas entre proyectos y a lo largo del tiempo.

#### Acceptance Criteria

1. THE ModuloPersistencia SHALL mantener una tabla `proyecto` con un registro por cada activo BESS gestionado: Arena (`IdProyecto=1`, en ejecución), Copiapo A (`IdProyecto=2`), Luz del Norte (`IdProyecto=3`) y Maria Elena (`IdProyecto=4`), con `Estado = 'por_implementar'` para los tres últimos.
2. THE tabla `proyecto` SHALL almacenar los parámetros de configuración del activo: `NumPCS`, `NumBateriasPorPCS`, `NumRacksPorBAC`, `TotalRacks` (calculado), `MinutosMuestreo`, `ZonaHoraria` y `FechaInicio`.
3. THE ModuloPersistencia SHALL mantener una tabla `tipo_detencion` con el catálogo completo de 68 códigos de falla extraído de la hoja `PCS-Fault` del Excel, con campos `CodigoFalla` (ej: F55), `DescripcionFallaPE` (ej: Fallo externo) y `CodigoDescripcion` (ej: F55 Fallo externo).
4. WHEN el MotorEventosFalla persiste un evento de falla, THE ModuloPersistencia SHALL insertar un registro en `detencion` con: `IdProyecto`, `IdCorrida`, `NumeroPCS`, `FechaInicio`, `FechaTermino`, `DuracionSegundos` (calculado como DATEDIFF en segundos), `IdTipoDetencion` (FK al catálogo), `CodigoFalla`, `DescripcionFalla`, `DescripcionFallaFallback`, `PromedioBateriasInvolucradas` y `HorasRackIndisponibles`.
5. THE tabla `detencion` SHALL incluir campos de calidad de datos: `ModulosDisponiblesNulo` (BIT) y `EsExcusable` (BIT).
6. THE tabla `detencion` SHALL incluir campos de workflow operacional: `Observacion` (NVARCHAR(500), nullable) para anotaciones libres y `EstadoRevision` (NVARCHAR(20)) con valor por defecto pendiente y valores permitidos: pendiente, revisado, excluido.
7. WHEN el ModuloPersistencia resuelve `IdTipoDetencion` para una detención, THE ModuloPersistencia SHALL buscar en `tipo_detencion` por `CodigoFalla`; IF el código no existe en el catálogo, THEN THE ModuloPersistencia SHALL insertar la detención con `IdTipoDetencion = NULL` y registrar una advertencia.
8. THE tabla `detencion` es la vista operacional de `fault_event`: ambas tablas coexisten. `fault_event` es la tabla técnica de auditoría (append-only por `IdCorrida`), `detencion` es la tabla de negocio consultada por analistas y sistemas de reporting.
9. THE tabla `etl_run` SHALL incluir el campo `IdProyecto` (INT FK -> proyecto) para asociar cada corrida con su activo correspondiente.

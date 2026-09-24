# Requirements Document

> Revisión 2 — 2026-09-24. Corregida contra el VBA y las fórmulas reales del libro. Ver `audit.md` para la evidencia (hallazgos `F-xx`) y las decisiones abiertas (`D-xx`).

## Introduction

Este proyecto reemplaza progresivamente el cálculo de disponibilidad del activo **Arena BESS**, implementado hoy en un libro Excel con macros VBA, por un proceso reproducible y auditable en **Python + SQL Server**.

La primera versión (`availability-v1-excel-parity`) busca **paridad exacta** con el Excel: mismos `BloquesMuestreo` (C12), `BloquesRacksIndisponibles` (C14), `DisponibilidadPeriodo` (C16), `DisponibilidadAnualAcumulada` (C19), misma lista de `ListOfFaults` y mismos KPI diarios, con el mismo input. **Paridad significa reproducir también los defectos del VBA** (marcados con flags de auditoría); corregirlos es una versión posterior.

**Flujo semanal (Fase S):** cada lunes se exporta `RawData-PCS` desde el server SCADA a `data/inbox/` (TeamViewer hoy), se carga en una copia de trabajo del `.xlsm`, se ejecutan las macros de referencia vía COM en la PC local, el pipeline Python escribe en SQL Server con un `IdCorrida` nuevo, se reconcilia contra la referencia Excel y se notifica a Power BI (owner Misael).

| Parámetro del activo | Valor |
|---|---|
| Fecha de inicio de operación | 08/Abril/2026 (primer dato: 2026-04-08 00:15) |
| Total PCS | 61 |
| Módulos BEC por PCS | 4 |
| Racks por BAC | 12 |
| Total Racks | 2.928 |
| Frecuencia de muestreo | 15 minutos |
| Versión del algoritmo | `availability-v1-excel-parity` |

### Mapa de numeración (revisión 1 → revisión 2)

| Rev. 1 | Rev. 2 | | Rev. 1 | Rev. 2 |
|---|---|---|---|---|
| 10 | 1 | | 6 (diaria) | 9 |
| 14 | 2, 3 | | 6 (anual) | 10 |
| 1 | 4 | | 7 | 11 |
| 2 | 5 | | 13 | 12 |
| 3 | 6 | | 8 | 13 |
| 4 | 7 | | 12 | 14 |
| 5 | 8 | | 9 | 15 |
| 11 | 16 | | 15 | 18 |
| — | 17 (nuevo) | | — | 19 (nuevo) |

---

## Glossary

- **PCS**: Power Conversion System. 61 unidades.
- **BEC**: Battery Energy Controller (módulo de batería). 4 por PCS.
- **BAC**: Battery Array Container. 12 racks por BAC.
- **Rack**: unidad mínima de batería. Total 2.928.
- **Fila origen** (`NumeroFilaOrigen`): número de fila de la hoja `RawData-PCS` (la primera fila de datos es la 2). Es la **clave de unión** con `PlantActivity` en modo paridad.
- **Serial Excel** (`SerialFechaExcelOrigen`): valor `double` crudo de la celda de fecha (días desde 1899-12-30). Los motores de paridad comparan y restan seriales, no `datetime`.
- **Semántica Excel**: reglas de comparación y conversión del VBA/Excel que el Python debe emular (vacío, texto vs número, `"Yes"`/`"No"` exactos, `FIND`/`MID`/`IFERROR`, conversión número→texto).
- **Parámetros KPI**: `C5` (inicio), `C7` (fin), `C21` (solo tiempo operacional), `C31` (evento excusable) de `Calculation-Availability`.
- **Parámetros de eventos**: `L2` (inicio), `L4` (fin), `L14` (evento excusable) de `ListOfFaults`. Son **independientes** de los parámetros KPI.
- **Fin Daily**: `Daily!D5`, fin del rango de días de la hoja `Daily` (máx. 31 días desde `C5`).
- **MotorETL**: el sistema Python completo.
- **ModuloAdquisicion** (Fase S): espera/valida el export SCADA en `data/inbox`.
- **ModuloLibroTrabajo** (Fase T/M): carga el export en la copia de trabajo del `.xlsm` y ejecuta las macros vía COM.
- **OrquestadorLunes**: `scripts/run_lunes.py` con etapas `acquire-wait`, `prepare-workbook`, `run-macros`, `run-etl`, `reconcile`, `notify-bi`, `all`.
- **ModuloIngesta**, **ModuloNormalizacion**, **ModuloEnriquecimiento**, **MotorDisponibilidad** (`cmdCalcAvailability`), **MotorEventosFalla** (`mcoCreateList`), **ModuloAgregacion** (`mcoDailyAvailability` + `Annual_AVA`), **ModuloPersistencia**, **ModuloReconciliacion**: componentes del pipeline (ver design).
- **IdCorrida**: UUID de una ejecución del MotorETL. Clave de auditoría en todas las tablas.
- **VersionAlgoritmo**: versión de la lógica (`availability-v1-excel-parity`).
- **Referencia Excel**: valores producidos por las macros en el libro de trabajo y extraídos por el runner COM (C12/C14/C16/C19, tabla de resultados, ListOfFaults, Daily).
- **Golden reference**: referencia Excel congelada en `tests/golden/data/` para tests.
- **BloquesMuestreo (C12)**: cantidad de **filas** de `RawData-PCS` dentro del período.
- **BloquesRacksIndisponibles (C14)**: Σ `racks_por_pcs × BateriasIndisponiblesPonderadas × [FactorOperacional]`.
- **DisponibilidadPeriodo (C16)**: `1 - C14 / (TotalRacks × C12)`.
- **DisponibilidadAnualAcumulada (C19)**: `1 - C14 / (TotalRacks × 365 × 24 × 4)`.
- **DisponibilidadAcumuladaAnual (Annual_AVA!J)**: `1 - Σ indisponibles mensuales / (TotalRacks × Σ bloques mensuales)`. Métrica distinta de C19.
- **ModulosDisponiblesNulo**: `True` si `NUMBER_OF_MODULES` estaba vacío.
- **DescripcionFallaFallback**: `True` si la descripción del evento se tomó de la fila anterior.
- **EventoArrastradoExcel**: `True` si el evento reproduce el defecto VBA de arrastre entre PCS (F-06).
- **FactorOperacional**: `PlantActivity!C` (1 activo / 0 inactivo). **FactorExcusable**: `PlantActivity!D` (1 = no excusado, 0 = excusado).
- **Proyecto**, **Detencion**, **TipoDetencion**, **EstadoRevision**: modelo de negocio multi-proyecto (Req 12).

---

## Requirements

### Requirement 1: Configuración y versionado

**User Story:** Como desarrollador, quiero que todos los parámetros de negocio provengan de la configuración de la corrida, para adaptar el sistema sin tocar los motores.

#### Acceptance Criteria

1. THE MotorETL SHALL leer los parámetros desde `ConfiguracionCalculo`: `total_pcs`, `baterias_por_pcs`, `racks_por_pcs`, `minutos_muestreo`, `fecha_inicio_proyecto`, parámetros KPI (`inicio_periodo`, `fin_periodo`, `solo_tiempo_operacional`, `aplicar_evento_excusable`), parámetros de eventos (`inicio_periodo_eventos`, `fin_periodo_eventos`, `aplicar_evento_excusable_eventos`), `fin_diario`, `modo_huecos` y `version_algoritmo`.
2. THE MotorETL SHALL calcular `total_racks = total_pcs × baterias_por_pcs × racks_por_pcs`; SHALL NOT usar 2.928, 61, 4, 12 ni 15 como constantes en los motores.
3. THE MotorETL SHALL proveer valores por defecto de paridad: `nombre_proyecto="Arena BESS"`, `fecha_inicio_proyecto=2026-04-08`, `total_pcs=61`, `baterias_por_pcs=4`, `racks_por_pcs=12`, `minutos_muestreo=15`, `modo_huecos="excel"`, `version_algoritmo="availability-v1-excel-parity"`.
4. WHERE los parámetros de eventos no se informan, THE MotorETL SHALL usar los parámetros KPI equivalentes y registrar en `etl_run` los valores efectivamente usados.
5. WHEN la configuración se construye desde valores de celdas Excel, THE MotorETL SHALL interpretar `solo_tiempo_operacional = (C21 <> "No")` y `aplicar_evento_excusable = (C31 = "Yes")`, `aplicar_evento_excusable_eventos = (L14 = "Yes")`, con comparación exacta sensible a mayúsculas (F-12).
6. WHEN se modifica una regla de negocio, THE MotorETL SHALL crear una nueva `VersionAlgoritmo` y nunca sobrescribir resultados de versiones anteriores.

---

### Requirement 2: Adquisición SCADA semanal (Fase S)

**User Story:** Como operador, quiero bajar cada lunes el export SCADA a `data/inbox` y validarlo de forma reproducible, sin procesar en el server SCADA.

#### Acceptance Criteria

1. WHEN se ejecuta `acquire-wait`, THE ModuloAdquisicion SHALL esperar (con timeout configurable) un archivo en `data/inbox` que cumpla la convención de nombre, validar que no está vacío y es legible, calcular su sha256 y registrarlo en log.
2. IF no hay archivo al vencer el timeout, THEN THE ModuloAdquisicion SHALL fallar con un mensaje accionable (riesgo J: sin API/UNC).
3. THE ModuloAdquisicion SHALL mover el original a `data/processed/<corte>/` junto a su `.sha256`; SHALL NOT modificar ese archivo después.
4. THE ModuloAdquisicion SHALL validar el data contract acordado en P1 (columnas = encabezados de `RawData-PCS`, delimitador, encoding, formato de fecha origen) y SHALL fallar sin tocar el libro de trabajo si no se cumple.
5. THE ModuloAdquisicion SHALL validar que el rango de fechas del reporte cumple la convención (`DESDE` = primer dato o `01-01-2026`; `HASTA` = último domingo o rango configurado).
6. THE Sistema SHALL NOT ejecutar macros ni el pipeline ETL en el server SCADA.

---

### Requirement 3: Libro de trabajo y macros de referencia (Fases T/M)

**User Story:** Como ingeniero de datos, quiero que las macros Excel corran automáticamente sobre el mismo input que el pipeline Python, para tener una referencia de paridad por corrida.

#### Acceptance Criteria

1. WHEN se ejecuta `prepare-workbook`, THE ModuloLibroTrabajo SHALL crear una copia de trabajo del `.xlsm` maestro en `data/work/<corte>/` y un backup; SHALL NOT editar el maestro.
2. THE ModuloLibroTrabajo SHALL escribir `RawData-PCS` en la copia de trabajo **vía COM** (no `openpyxl`), con la columna A como fecha/serial Excel real; el formato `dd-mm-aaaa hh:mm:ss` SHALL ser solo formato de visualización (F-21, F-22).
3. THE ModuloLibroTrabajo SHALL convertir fechas origen `mm-dd-aaaa hh:mm:ss` con un parser estricto; IF alguna fecha no parsea, THEN SHALL fallar la etapa.
4. WHEN termina la carga, THE ModuloLibroTrabajo SHALL verificar la alineación por fila con `PlantActivity` (timestamp de `PlantActivity!B` = `RawData-PCS!A` en cada fila donde exista) y reportar las filas desalineadas o sin timestamp.
5. WHEN se ejecuta `run-macros`, THE ModuloLibroTrabajo SHALL escribir `C5`, `C7`, `C21`, `C31`, `ListOfFaults!L2`, `L4`, `L14` y `Daily!D5` desde la configuración, y ejecutar en orden `cmdCalcAvailability` → `mcoCreateList` → `mcoDailyAvailability` → `Graphupdate` en Excel local vía COM.
6. WHEN las macros terminan, THE ModuloLibroTrabajo SHALL extraer como referencia: `C12`, `C14`, `C16`, `C19`, `C23`, `ListOfFaults!L10`, la tabla `Calculation-Availability!E4:BO<n>`, la lista completa `ListOfFaults!B6:I<n>` y `Daily!B9:G39`, y guardarla en `data/work/<corte>/referencia_excel.json` y en SQL (tablas `excel_reference_*`).
7. IF una macro falla o excede el timeout, THEN THE ModuloLibroTrabajo SHALL cerrar solo la instancia de Excel que creó, guardar el log y marcar la referencia como ausente; el pipeline Python SHALL poder continuar con advertencia.

---

### Requirement 4: Ingesta y staging

**User Story:** Como ingeniero de datos, quiero leer el libro de trabajo conservando el dato original, para que cualquier reproceso sea trazable.

#### Acceptance Criteria

1. WHEN el ModuloIngesta recibe la ruta de un libro válido, THE ModuloIngesta SHALL leer `RawData-PCS` y `PlantActivity` desde la fila 2 conservando todas las columnas y el número de fila origen.
2. THE ModuloIngesta SHALL obtener `SerialFechaExcelOrigen` del valor numérico crudo de la celda (sin reconvertir desde `datetime`) y `MarcaTiempoLocalOrigen` como `datetime` local naive.
3. IF el archivo o una hoja requerida no existe o no es legible, THEN THE ModuloIngesta SHALL marcar la corrida `Estado="failed"` con el error en `etl_run.MensajeError`.
4. THE ModuloIngesta SHALL detectar y reportar sin abortar: huecos (salto > `minutos_muestreo`), timestamps duplicados, timestamps fuera de orden, frecuencia efectiva distinta (`ROUND(Δserial × 1440, 2)`), celdas A vacías, filas de `PlantActivity` sin timestamp o desalineadas.
5. WHERE `modo_huecos = "excel"`, WHEN la columna A tiene la primera celda vacía, THE ModuloIngesta SHALL ignorar esa fila y todas las siguientes (como el VBA) y reportar cuántas filas se descartaron. WHERE `modo_huecos = "continuar"`, SHALL omitir solo las filas vacías.
6. IF una celda `NUMBER_OF_MODULES` contiene texto no numérico, THEN THE ModuloIngesta SHALL rechazar la corrida en modo paridad (el VBA fallaría con *Type mismatch*) indicando fila y PCS (F-16).
7. THE ModuloIngesta SHALL conservar el tipo crudo de `CURRENT FAULT` (texto, número o vacío).
8. THE ModuloStaging SHALL persistir los registros sin transformaciones de negocio, asociados al `IdCorrida`.

---

### Requirement 5: Normalización

**User Story:** Como desarrollador del motor, quiero los datos PCS en formato largo, para no depender del índice de columna frágil del Excel.

#### Acceptance Criteria

1. THE ModuloNormalizacion SHALL transformar el formato ancho (61 × 4 columnas) a una fila por `(NumeroFilaOrigen, NumeroPCS)`, conservando `SerialFechaExcelOrigen` y `MarcaTiempoMuestra`.
2. THE ModuloNormalizacion SHALL identificar las columnas por el patrón `Arena - PCS XX - POWERELECTRONICS <campo>` con los campos exactos `GEN3 HEx CURRENT FAULT`, `GEN3 HEx CURRENT STATUS`, `GEN3 HEx CURRENT WARNING`, `HEM-k NUMBER OF MODULES`.
3. WHEN `NUMBER_OF_MODULES` está vacío, THE ModuloNormalizacion SHALL asignar `ModulosDisponibles = baterias_por_pcs`, `ModulosDisponiblesNulo = True` y `ModulosRaw = NULL`.
4. WHEN `NUMBER_OF_MODULES` es numérico, THE ModuloNormalizacion SHALL conservar el valor como `float` **sin redondear ni truncar** (hay valores fraccionarios, F-02) y `ModulosDisponiblesNulo = False`.
5. THE ModuloNormalizacion SHALL detectar y reportar PCS faltantes, columnas faltantes por PCS, nombres inesperados y columnas desplazadas respecto del orden `PCS k → columnas 4k-2 … 4k+1`.
6. THE ModuloNormalizacion SHALL conservar el **orden de fila origen**; SHALL NOT reordenar ni deduplicar por timestamp en modo paridad (F-07).

---

### Requirement 6: Enriquecimiento

**User Story:** Como desarrollador del motor, quiero asociar los factores de PlantActivity a cada fila exactamente como lo hace el VBA.

#### Acceptance Criteria

1. THE ModuloEnriquecimiento SHALL asociar a cada fila de `RawData-PCS` los valores de `PlantActivity!C` (FactorOperacional) y `PlantActivity!D` (FactorExcusable) de la **misma fila origen** (F-01).
2. IF la celda de `PlantActivity` está vacía o la fila no existe, THEN THE ModuloEnriquecimiento SHALL usar `0` (semántica de celda vacía en Excel) y registrar la anomalía.
3. WHEN la fila de `PlantActivity` tiene timestamp distinto del de `RawData-PCS`, THE ModuloEnriquecimiento SHALL registrar la desalineación con ambos valores.
4. THE ModuloEnriquecimiento SHALL producir para cada `(NumeroFilaOrigen, NumeroPCS)`: `ModulosDisponibles`, `ModulosDisponiblesNulo`, `FallaRaw`, `FactorOperacional`, `FactorExcusable`.

---

### Requirement 7: Motor de disponibilidad (`cmdCalcAvailability`)

**User Story:** Como analista, quiero que el motor Python calcule C12, C14, C16 y C19 exactamente como la macro, para compararlos antes de retirar el Excel.

#### Acceptance Criteria

1. THE MotorDisponibilidad SHALL recorrer las filas en orden de fila origen y procesar una fila si `serial ≥ serial(inicio_periodo)` AND `serial < serial(fin_periodo) + 1`.
2. THE MotorDisponibilidad SHALL incrementar `BloquesMuestreo` en 1 por **cada fila** procesada, independientemente del número de PCS y aunque existan timestamps repetidos.
3. THE MotorDisponibilidad SHALL considerar un PCS indisponible en una fila únicamente si `ModulosDisponiblesNulo = False` AND `ModulosDisponibles < 4` (umbral literal del VBA, parametrizado como `baterias_por_pcs`).
4. WHEN un PCS es indisponible, THE MotorDisponibilidad SHALL calcular `BateriasIndisponibles = baterias_por_pcs − ModulosDisponibles` como `float`.
5. WHERE `aplicar_evento_excusable`, THE MotorDisponibilidad SHALL calcular `BateriasIndisponiblesPonderadas = BateriasIndisponibles × FactorExcusable`; en otro caso `= BateriasIndisponibles`.
6. WHERE `solo_tiempo_operacional`, THE MotorDisponibilidad SHALL acumular `C14 += racks_por_pcs × BateriasIndisponiblesPonderadas × FactorOperacional`; en otro caso `C14 += racks_por_pcs × BateriasIndisponiblesPonderadas`.
7. THE MotorDisponibilidad SHALL acumular C14 en el mismo orden que el VBA (fila, luego PCS 1..N) y sin redondeos intermedios.
8. THE MotorDisponibilidad SHALL calcular `DisponibilidadPeriodo = 1 − C14 / (TotalRacks × C12)` y `DisponibilidadAnualAcumulada = 1 − C14 / (TotalRacks × 365 × 24 × 4)`; IF el denominador es 0, THEN SHALL producir `None` (equivalente a `IFERROR(…,"N/A")`).
9. THE MotorDisponibilidad SHALL derivar `MinutosMuestreoDerivado = ROUND((serial₂ − serial₁) × 1440, 2)` de las dos primeras filas procesadas (C23) y reportar si difiere de `minutos_muestreo` (F-13).
10. THE MotorDisponibilidad SHALL producir una fila de `availability_sample_result` por `(NumeroFilaOrigen, NumeroPCS)` procesado con: `ModulosDisponibles`, `ModulosDisponiblesNulo`, `BateriasIndisponibles`, `FactorExcusable`, `FactorOperacional`, `BateriasIndisponiblesPonderadas`, `ImpactoRackPonderado`.

---

### Requirement 8: Motor de eventos de falla (`mcoCreateList`)

**User Story:** Como analista, quiero la lista de eventos idéntica a `ListOfFaults`, incluyendo el fallback de descripción y los defectos conocidos, para reconciliar evento por evento.

#### Acceptance Criteria

1. THE MotorEventosFalla SHALL recorrer los PCS en el orden de los bloques de encabezado de `RawData-PCS` y, para cada PCS, todas las filas en orden de fila origen, usando los **parámetros de eventos** (`L2`, `L4`, `L14`).
2. THE MotorEventosFalla SHALL considerar una fila si `serial ≥ L2` AND `serial < L4 + 1` AND `ModulosDisponiblesNulo = False` AND `ModulosDisponibles < 4`.
3. WHEN considera una fila, THE MotorEventosFalla SHALL acumular `sumablocks += (4 − Modulos) × FactorExcusable` si `aplicar_evento_excusable_eventos`, o `sumablocks = (sumablocks + 4) − Modulos` en otro caso (orden de evaluación del VBA), y `numBlock += 1`.
4. THE MotorEventosFalla SHALL iniciar un registro de evento cuando la fila **anterior de la hoja completa** tiene módulos `= 4`, está vacía, o su serial es `< L2`. El inicio SHALL ser `serial_actual − MinutosMuestreoDerivado / 1440` y `NumeroPCS` el del bloque.
5. THE MotorEventosFalla SHALL cerrar el evento cuando la **fila siguiente** tiene módulos `= 4`, está vacía (o no existe), o su serial es **estrictamente** `> L4 + 1`. El fin SHALL ser el serial de la **fila actual** (última en falla) (F-03).
6. WHEN cierra, THE MotorEventosFalla SHALL calcular `DuracionHoras = 24 × (fin − inicio)` en seriales, `PromedioBateriasInvolucradas = sumablocks / numBlock`, `HorasRackIndisponibles = racks_por_pcs × DuracionHoras × Promedio`, acumular `L10`, y reiniciar `sumablocks` y `numBlock`.
7. THE MotorEventosFalla SHALL mantener `sumablocks`, `numBlock` y el registro abierto **entre PCS** como el VBA; WHEN un evento quedó abierto y otro PCS lo sobrescribe o cierra, SHALL marcar `EventoArrastradoExcel = True` (F-06).
8. WHEN determina la descripción al iniciar, THE MotorEventosFalla SHALL aplicar en orden: (a) `G = CURRENT FAULT` de la fila actual; (b) si `G = "NO FAULTS"`: si la fila anterior `≠ "NO FAULTS"` entonces `G = valor de la fila anterior` y `DescripcionFallaFallback = True`, si no `G = "F13 NO MODULES"`; (c) si `G` es vacío entonces `G = "F1 Watchdog"`. Una fila anterior vacía produce por tanto `"F1 Watchdog"` (F-04).
9. THE MotorEventosFalla SHALL calcular `CodigoFalla` emulando `IFERROR(MID(G,1,FIND(" ",G,1)-1), "F"&G)`: texto hasta el primer espacio (cadena vacía si el espacio está en la posición 1); `"F" & texto_excel(G)` si no hay espacio; números convertidos a texto como Excel (`55` → `"F55"`) (F-14).
10. IF un evento debe iniciar en la primera fila de datos (la anterior es el encabezado, donde el VBA falla con *Type mismatch*), THEN THE MotorEventosFalla SHALL iniciar el evento y marcarlo `ExcelHabriaFallado = True` (F-11, D-08).
11. THE MotorEventosFalla SHALL producir registros `fault_event` con: `NumeroPCS`, `OrdenExcel`, `SerialInicio`, `SerialFin`, `MarcaTiempoInicio`, `MarcaTiempoFin`, `DuracionHoras`, `CodigoFalla`, `DescripcionFalla`, `DescripcionFallaFallback`, `NumeroBloques`, `PromedioBateriasInvolucradas`, `HorasRackIndisponibles`, `EventoArrastradoExcel`, `ExcelHabriaFallado`.
12. THE MotorEventosFalla SHALL producir el resumen por código de falla equivalente a `ListOfFaults!N:Q` (Σ rack-hours por código, % sobre L10), ordenado de mayor a menor. No existe límite de eventos (F-18).

---

### Requirement 9: Disponibilidad diaria (`mcoDailyAvailability`)

**User Story:** Como analista, quiero el KPI diario acumulado idéntico a la hoja `Daily`.

#### Acceptance Criteria

1. THE ModuloAgregacion SHALL generar un día por fecha desde `inicio_periodo` hasta `fin_diario` inclusive (default `fin_periodo`), con un máximo de 31 días; `DiaN` = 1 para `inicio_periodo` (F-09).
2. THE ModuloAgregacion SHALL calcular `BloquesRacksIndisponiblesDiarios` = Σ `racks_por_pcs × BateriasIndisponiblesPonderadas` de las filas procesadas por el MotorDisponibilidad cuyo serial cae en `[día, día+1)`, **sin** aplicar `FactorOperacional` (F-08).
3. THE ModuloAgregacion SHALL calcular `BloquesRacksIndisponiblesAcumulados(N) = Diario(N) + Acumulado(N−1)`.
4. THE ModuloAgregacion SHALL calcular `Disponibilidad(N) = 1 − Acumulado(N) / (TotalRacks × 24 × 60 × DiaN / MinutosMuestreoDerivado)`; IF el denominador es 0, THEN `None`.
5. THE ModuloAgregacion SHALL calcular `Variacion(1) = 0` y `Variacion(N) = Disponibilidad(N) − Disponibilidad(N−1)`, `None` si alguno es `None`.
6. THE ModuloAgregacion SHALL incluir días sin datos dentro del rango con valor diario 0 (el denominador sigue creciendo, como en Excel).

---

### Requirement 10: KPI mensual y acumulado anual (`Annual_AVA`)

**User Story:** Como analista, quiero el KPI mensual y el acumulado anual con la misma fórmula que la hoja `Annual_AVA`, sabiendo que esa hoja se mantiene a mano.

#### Acceptance Criteria

1. THE ModuloAgregacion SHALL mantener una tabla `monthly_official_kpi` con una fila vigente por `(IdProyecto, Anio, Mes)`: `DiasMes`, `BloquesMuestreo`, `BloquesRacksIndisponibles`, `Origen` (`corrida` | `excel_manual`) e `IdCorrida` (nullable).
2. WHEN una corrida marcada como oficial cubre un mes, THE ModuloAgregacion SHALL registrar para ese mes `BloquesRacksIndisponibles = C14` y `DiasMes = serial(última fila en rango) − serial(primer día del mes)` si el mes está incompleto, o los días calendario si está completo; `BloquesMuestreo = DiasMes × 24 × 60 / minutos_muestreo`.
3. THE ModuloAgregacion SHALL calcular por mes `DisponibilidadMensual = 1 − Indisp / (TotalRacks × Bloques)` y los acumulados `BloquesMuestreoAcumulados`, `BloquesIndisponiblesAcumulados` y `DisponibilidadAcumulada = 1 − IndispAcum / (TotalRacks × BloquesAcum)` desde `mes_inicio_acumulado` (D-06).
4. THE ModuloAgregacion SHALL permitir importar como `Origen = excel_manual` las filas históricas de `Annual_AVA` que no se pueden reproducir (jul/ago 2026) (F-20).
5. THE ModuloAgregacion SHALL registrar `DisponibilidadContractual` (0,98 por defecto) por mes.
6. THE MotorETL SHALL documentar que C19 (fórmula 365 días sobre el período) y `DisponibilidadAcumulada` (Annual_AVA!J) son métricas distintas y persistir ambas.

---

### Requirement 11: Persistencia SQL Server

**User Story:** Como administrador, quiero que resultados e intermedios se guarden append-only por `IdCorrida`.

#### Acceptance Criteria

1. THE ModuloPersistencia SHALL operar append-only sobre las tablas de corrida: nunca `DELETE` ni `UPDATE` de filas de corridas anteriores (excepto la transición de estado de su propio `etl_run`).
2. THE ModuloPersistencia SHALL crear `etl_run` con `Estado="running"` al inicio y terminar en `success`, `failed` o `parity_failed`.
3. THE ModuloPersistencia SHALL insertar todos los datos de la corrida en **una transacción**; IF falla, THEN SHALL hacer rollback completo y registrar `failed` en una transacción separada.
4. THE ModuloPersistencia SHALL persistir en `etl_run` los parámetros efectivos (KPI, eventos, `fin_diario`, `modo_huecos`), `ArchivoOrigen`, `HashArchivoOrigen` (sha256), `MinutosMuestreoDerivado`, `TipoCorrida`, `EsOficial`, `VersionAlgoritmo` e `IdProyecto`.
5. THE ModuloPersistencia SHALL usar inserción masiva (`fast_executemany` con ODBC Driver 18 o `BULK INSERT`) para las tablas de muestras.
6. THE ModuloPersistencia SHALL almacenar módulos, baterías y factores como `FLOAT`.

---

### Requirement 12: Proyectos, catálogo y detenciones

**User Story:** Como analista, quiero consultar detenciones entre proyectos y en el tiempo, con revisión operacional que sobreviva a los reprocesos semanales.

#### Acceptance Criteria

1. THE ModuloPersistencia SHALL mantener `proyecto` con Arena (`IdProyecto=1`, `en_ejecucion`), Copiapó A (2), Luz del Norte (3) y María Elena (4) (`por_implementar`).
2. THE tabla `proyecto` SHALL almacenar `NumPCS`, `NumBateriasPorPCS`, `NumRacksPorBAC`, `TotalRacks` (calculado), `MinutosMuestreo`, `ZonaHoraria`, `FechaInicio`.
3. THE ModuloPersistencia SHALL cargar `tipo_detencion` desde la hoja `PCS-Fault` completa (**167 códigos**, F0…F257) con `CodigoFalla`, `DescripcionFallaPE`, `CodigoDescripcion`, `Significado` y `Operativo` (F-19); el seed SHALL ser idempotente y generado por script.
4. WHEN se persiste un evento, THE ModuloPersistencia SHALL insertar una fila en `detencion` con `IdProyecto`, `IdCorrida`, `NumeroPCS`, `FechaInicio`, `FechaTermino`, `DuracionSegundos`, `IdTipoDetencion` (NULL + advertencia si el código no existe), `CodigoFalla`, `DescripcionFalla`, flags de calidad y rack-hours.
5. THE Sistema SHALL guardar el workflow (`EstadoRevision` ∈ {pendiente, revisado, excluido}, `Observacion`, `RevisadoPor`, `RevisadoEn`) en `detencion_revision`, con clave de negocio `(IdProyecto, NumeroPCS, FechaInicio)`, para que sobreviva a nuevas corridas (F-27).
6. THE Sistema SHALL exponer `v_detencion_vigente`: detenciones de la última corrida oficial por período, unidas con su revisión.
7. THE tabla `etl_run` SHALL incluir `IdProyecto` (FK a `proyecto`).

---

### Requirement 13: Reconciliación

**User Story:** Como ingeniero de datos, quiero comparar automáticamente cada corrida Python contra la referencia Excel en cinco niveles.

#### Acceptance Criteria

1. THE ModuloReconciliacion SHALL comparar en cinco niveles: (1) input, (2) muestra individual, (3) acumulados, (4) KPI, (5) eventos.
2. Nivel 1 SHALL comparar: filas procesadas, primer/último serial, cantidad de PCS, `MinutosMuestreoDerivado` (C23) y parámetros efectivos (C5/C7/C21/C31/L2/L4/L14/D5).
3. Nivel 2 SHALL comparar por `(NumeroFilaOrigen, NumeroPCS)` las `BateriasIndisponiblesPonderadas` contra la tabla de resultados de `Calculation-Availability` (celda vacía = 0).
4. Nivel 3 SHALL comparar C12 (exacto) y C14.
5. Nivel 4 SHALL comparar C16, C19 y la serie diaria completa (diario, acumulado, disponibilidad, variación).
6. Nivel 5 SHALL comparar la lista de eventos **en orden de escritura** (`OrdenExcel`) campo a campo, `L10`, el conteo y el resumen por código.
7. THE ModuloReconciliacion SHALL usar las tolerancias de la tabla única de `design.md` §Tolerancias.
8. THE ModuloReconciliacion SHALL evaluar los invariantes cruzados: `C14 = 4 × L10` (cuando períodos y flags de ambos motores coinciden, `solo_tiempo_operacional` es falso y no hay eventos arrastrados) y por PCS `Σ ponderadas × racks / 4 = Σ rack-hours`.
9. WHEN finaliza, THE ModuloReconciliacion SHALL producir un `ReporteReconciliacion` (pass/fail por nivel y discrepancias), persistirlo en `reconciliation_result` y, si falla algún nivel 3–5, marcar la corrida `parity_failed`.
10. IF no hay referencia Excel para la corrida, THEN THE ModuloReconciliacion SHALL registrar `sin_referencia` sin fallar la corrida.

---

### Requirement 14: Golden references

**User Story:** Como ingeniero de datos, quiero validar el motor contra valores congelados del Excel en CI.

#### Acceptance Criteria

1. THE repositorio SHALL contener goldens en `tests/golden/data/` indexados por `golden_index.json` (fuente de verdad del estado `pending` / `verified`).
2. THE golden de cada mes SHALL contener: parámetros efectivos (incl. L2/L4/L14, Daily!D5, C23), KPI (C12/C14/C16/C19, L10), Daily completo, snapshot `Annual_AVA`, **todos** los eventos de `ListOfFaults` en orden de escritura, el resumen por código y la tabla de resultados de `Calculation-Availability` (archivo comprimido aparte si excede 1 MB).
3. THE extractor `tests/golden/extract_golden.py` SHALL leer valores con `openpyxl data_only=True`, validar invariantes y fallar si no se cumplen; el invariante `Σ daily = C14` SHALL aplicarse solo si `C21 = "No"`.
4. THE suite SHALL ejecutar el golden runner sobre todos los meses `verified` y el test de integridad sobre todos los meses.
5. IF una comparación contra un golden `verified` falla, THEN el test SHALL fallar con el detalle de la discrepancia.
6. Los valores de `Annual_AVA` de meses no vinculados a C14 SHALL tratarse como históricos, nunca como esperado de paridad (GT-5).

---

### Requirement 15: Auditoría, trazabilidad y calidad de datos

**User Story:** Como auditor del KPI contractual, quiero rastrear cualquier valor hasta la celda de origen.

#### Acceptance Criteria

1. THE MotorETL SHALL garantizar la trazabilidad `availability_run_result → availability_sample_result → raw_pcs_sample → (archivo, hash, fila, columna)`.
2. THE MotorETL SHALL persistir cada anomalía en `data_quality_issue` (tipo, fila, PCS, timestamp, detalle) y un resumen JSON en `etl_run.ResumenCalidad`.
3. THE resumen de calidad SHALL incluir: filas con `ModulosDisponiblesNulo`, eventos con `DescripcionFallaFallback`, eventos `EventoArrastradoExcel` y `ExcelHabriaFallado`, anomalías de timestamp, filas truncadas por `modo_huecos`, desalineaciones y vacíos de PlantActivity, PCS/columnas faltantes, diferencia C23 vs config y descripciones numéricas.
4. THE MotorETL SHALL hacer consultables en toda la historia los intervalos `ModulosDisponiblesNulo = True`, sin suprimirlos.

---

### Requirement 16: Timestamps y DST

**User Story:** Como ingeniero de datos, quiero preservar los timestamps originales y no romper la paridad en los cambios de horario de Chile.

#### Acceptance Criteria

1. THE MotorETL SHALL tratar los timestamps como hora local de Chile **naive** (zona de negocio `America/Santiago`, documentada) y SHALL NOT convertirlos a UTC ni localizarlos en los motores.
2. THE motores de paridad SHALL filtrar, comparar y restar sobre `SerialFechaExcelOrigen` (F-15).
3. THE ModuloIngesta SHALL reportar los saltos (p. ej. 2026-09-06 00:00 → 01:00) y repeticiones (retroceso de abril) de hora local como anomalías DST informativas.
4. THE ModuloReconciliacion SHALL incluir al menos un período que cruce un cambio de horario en los goldens verificados (septiembre 2026).

---

### Requirement 17: Fidelidad a la semántica Excel/VBA

**User Story:** Como responsable de la paridad, quiero que las reglas implícitas de Excel/VBA estén centralizadas y probadas, para que los motores no las reinterpreten.

#### Acceptance Criteria

1. THE MotorETL SHALL centralizar en un módulo `excel_semantics` las funciones: `es_vacio`, `igual_numero` (celda `= 4`), `texto_excel` (número→texto como Excel), `codigo_falla_excel` (`IFERROR/MID/FIND`), `flag_si`/`flag_no` (comparación exacta) y conversión serial↔datetime.
2. THE módulo `excel_semantics` SHALL tener tests unitarios por cada regla con casos tomados del libro real.
3. THE motores SHALL NOT implementar comparaciones de celdas fuera de `excel_semantics`.

---

### Requirement 18: Notificación a Power BI

**User Story:** Como owner de reporting (Misael), quiero una notificación cuando SQL esté listo para refrescar el dashboard.

#### Acceptance Criteria

1. WHEN `run-etl` y `reconcile` terminan, THE OrquestadorLunes SHALL emitir una notificación con `IdCorrida`, período, estado y resumen de reconciliación y de calidad.
2. THE destino SHALL configurarse por variable de entorno (webhook); IF no está configurado, THEN THE OrquestadorLunes SHALL escribir la notificación en `data/work/<corte>/notificacion.md` y en consola.
3. THE Sistema SHALL NOT intentar refresh vía API de Power BI hasta que exista service principal (fuera de alcance v1).
4. THE Sistema SHALL exponer vistas SQL estables para Power BI (`v_kpi_vigente`, `v_daily_vigente`, `v_detencion_vigente`, `v_calidad_corrida`) de modo que el dashboard no dependa de `IdCorrida`.

---

### Requirement 19: Orquestador semanal

**User Story:** Como operador, quiero ejecutar el flujo del lunes por etapas o completo, reanudable ante fallas.

#### Acceptance Criteria

1. THE OrquestadorLunes SHALL ofrecer `--stage acquire-wait | prepare-workbook | run-macros | run-etl | reconcile | notify-bi | all` y los argumentos `--inbox`, `--work`, `--period-start`, `--period-end`, `--oficial`.
2. THE OrquestadorLunes SHALL guardar el estado de cada etapa en `data/work/<corte>/run_state.json` y permitir reanudar desde la etapa fallida sin repetir las exitosas.
3. WHEN no se informa el período, THE OrquestadorLunes SHALL calcular el período por defecto según D-07 (mes en curso hasta el último domingo; el primer lunes del mes, además, el cierre del mes anterior).
4. THE OrquestadorLunes SHALL registrar logs estructurados por etapa con timestamps y terminar con código de salida ≠ 0 ante error.

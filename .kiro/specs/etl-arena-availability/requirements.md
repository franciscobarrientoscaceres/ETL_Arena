# Requirements Document

> Revisión 2 — 2026-09-24. Corregida contra el VBA y las fórmulas reales del libro. Ver `audit.md` para la evidencia (hallazgos `F-xx`) y las decisiones abiertas (`D-xx`).
>
> **Revisión 3 (aceptada 2026-09-26, ADR-12, D-20…D-27):** la persistencia pasa de copias por `IdCorrida` a un **estado vigente por mes** que se reemplaza al recargar. Cambian R10, R11, R12, R15, R18 y R19 (marcados *rev. 3*). Los motores y la paridad no cambian. Plan: `docs/plan-estado-vigente.md`; tareas: Fase 6 de `tasks.md`.

## Introduction

Este proyecto reemplaza progresivamente el cálculo de disponibilidad del activo **Arena BESS**, implementado hoy en un libro Excel con macros VBA, por un proceso reproducible y auditable en **Python + SQL Server**.

La primera versión (`availability-v1-excel-parity`) buscó **paridad exacta**; desde 2026-09-24 la versión vigente es `availability-v1.1-exclusion-matrix` (F-37): igual a v1 salvo que la exclusión viene de `Exclusion_Matrix`. Sin esa hoja (libros hasta septiembre 2026) ambas dan resultados idénticos. La paridad busca con el Excel: mismos `BloquesMuestreo` (C12), `BloquesRacksIndisponibles` (C14), `DisponibilidadPeriodo` (C16), `DisponibilidadAnualAcumulada` (C19), misma lista de `ListOfFaults` y mismos KPI diarios, con el mismo input. **Paridad significa reproducir también los defectos del VBA** (marcados con flags de auditoría); corregirlos es una versión posterior.

**Flujo semanal (Fase S):** cada lunes se exporta `RawData-PCS` desde el server SCADA a `data/inbox/` (TeamViewer hoy), se carga en una copia de trabajo del `.xlsm`, se ejecutan las macros de referencia vía COM en la PC local, el pipeline Python calcula y se reconcilia contra la referencia Excel, **reemplaza en SQL Server los datos vigentes del mes** (rev. 3; cada ejecución queda registrada en `etl_run` con su `NumCorrida`) y se notifica a Power BI (owner Misael).

| Parámetro del activo | Valor |
|---|---|
| Fecha de inicio de operación | 08/Abril/2026 (primer dato: 2026-04-08 00:15) |
| Total PCS | 61 |
| Módulos BEC por PCS | 4 |
| Racks por BAC | 12 |
| Total Racks | 2.928 |
| Frecuencia de muestreo | 15 minutos |
| Versión del algoritmo | `availability-v1.1-exclusion-matrix` (F-37; antes `availability-v1-excel-parity`) |

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
- **VersionAlgoritmo**: versión de la lógica (`availability-v1.1-exclusion-matrix`; v1 = `availability-v1-excel-parity`).
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
- **FactorOperacional**: `PlantActivity!C` (1 activo / 0 inactivo); solo pondera con "Only Operational Time?" (C21).
- **Exclusion_Matrix / ValorExclusion** (F-37): hoja con una columna por PCS alineada por fila con `RawData-PCS`; valor por celda `0`/vacío = sin evento de exclusión (EE), `1` = EE con todos los módulos considerados en operación, `2` = EE con los módulos en operación **antes** del inicio del evento. Reemplaza a `PlantActivity!D` como fuente de la exclusión; `PlantActivity!D` queda solo como espejo de `Calculation-Availability!BO`.
- **BateriasPrevias**: para celdas con valor 2, baterías indisponibles consideradas en la fila anterior al primer 2 del tramo.
- **Proyecto**, **Detencion**, **TipoDetencion**, **EstadoRevision**: modelo de negocio multi-proyecto (Req 12).

---

## Requirements

### Requirement 1: Configuración y versionado

**User Story:** Como desarrollador, quiero que todos los parámetros de negocio provengan de la configuración de la corrida, para adaptar el sistema sin tocar los motores.

#### Acceptance Criteria

1. THE MotorETL SHALL leer los parámetros desde `ConfiguracionCalculo`: `total_pcs`, `baterias_por_pcs`, `racks_por_pcs`, `minutos_muestreo`, `fecha_inicio_proyecto`, parámetros KPI (`inicio_periodo`, `fin_periodo`, `solo_tiempo_operacional`, `aplicar_evento_excusable`), parámetros de eventos (`inicio_periodo_eventos`, `fin_periodo_eventos`, `aplicar_evento_excusable_eventos`), `fin_diario`, `modo_huecos` y `version_algoritmo`.
2. THE MotorETL SHALL calcular `total_racks = total_pcs × baterias_por_pcs × racks_por_pcs`; SHALL NOT usar 2.928, 61, 4, 12 ni 15 como constantes en los motores.
3. THE MotorETL SHALL proveer valores por defecto de paridad: `nombre_proyecto="Arena BESS"`, `fecha_inicio_proyecto=2026-04-08`, `total_pcs=61`, `baterias_por_pcs=4`, `racks_por_pcs=12`, `minutos_muestreo=15`, `modo_huecos="excel"`, `version_algoritmo="availability-v1.1-exclusion-matrix"`, `solo_tiempo_operacional=False`, `aplicar_evento_excusable=True` y `aplicar_evento_excusable_eventos=True` (D-03: el KPI oficial descuenta eventos excusables).
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
5. THE ModuloAdquisicion SHALL validar la continuidad del export incremental (D-07): el primer dato del export SHALL ser el siguiente al último dato de `RawData-PCS` en el libro base (último + `minutos_muestreo`, admitiendo el salto DST, que se reporta); IF hay hueco, THEN SHALL reportarlo y exigir confirmación (`--aceptar-hueco`); IF hay solape con valores distintos a los ya cargados, THEN SHALL rechazar la carga (D-13); el solape con valores idénticos SHALL descartarse y reportarse.
6. THE Sistema SHALL NOT ejecutar macros ni el pipeline ETL en el server SCADA.

---

### Requirement 3: Libro de trabajo y macros de referencia (Fases T/M)

**User Story:** Como ingeniero de datos, quiero que las macros Excel corran automáticamente sobre el mismo input que el pipeline Python, para tener una referencia de paridad por corrida.

#### Acceptance Criteria

1. WHEN se ejecuta `prepare-workbook`, THE ModuloLibroTrabajo SHALL crear una copia de trabajo del **libro base** en `data/work/<corte>/` y un backup; el libro base es el libro de trabajo de la última corrida oficial exitosa (o el maestro de `data/` en la primera corrida); SHALL NOT editar el libro base.
2. THE ModuloLibroTrabajo SHALL **agregar** las filas del export a continuación de la última fila con dato de `RawData-PCS` en la copia de trabajo, **vía COM** (no `openpyxl`), con la columna A como fecha/serial Excel real; el formato `dd-mm-aaaa hh:mm:ss` SHALL ser solo formato de visualización (F-21, F-22). Las filas existentes SHALL NOT modificarse (D-07, D-13).
3. THE ModuloLibroTrabajo SHALL convertir fechas origen `mm-dd-aaaa hh:mm:ss` con un parser estricto; IF alguna fecha no parsea, THEN SHALL fallar la etapa.
4. WHEN termina la carga, THE ModuloLibroTrabajo SHALL verificar la alineación por fila con `PlantActivity` (timestamp de `PlantActivity!B` = `RawData-PCS!A` en cada fila donde exista) y reportar las filas desalineadas o sin timestamp.
5. WHEN se ejecuta `run-macros`, THE ModuloLibroTrabajo SHALL escribir `C5`, `C7`, `C21`, `C31`, `ListOfFaults!L2`, `L4`, `L14` y `Daily!D5` desde la configuración, y ejecutar en orden `cmdCalcAvailability` → `mcoCreateList` → `mcoDailyAvailability` → `Graphupdate` en Excel local vía COM.
6. WHEN las macros terminan, THE ModuloLibroTrabajo SHALL extraer como referencia: `C12`, `C14`, `C16`, `C19`, `C23`, `ListOfFaults!L10`, la tabla `Calculation-Availability!E4:BO<n>`, la lista completa `ListOfFaults!B6:I<n>` y `Daily!B9:G39`, y guardarla en `data/work/<corte>/referencia_excel.json` y en SQL (tablas `excel_reference_*`).
7. IF una macro falla o excede el timeout, THEN THE ModuloLibroTrabajo SHALL cerrar solo la instancia de Excel que creó, guardar el log y marcar la referencia como ausente; el pipeline Python SHALL poder continuar con advertencia.
8. WHEN se ejecuta `load-exclusion-matrix` con la `Exclusion_Matrix` del mes (y opcionalmente PlantActivity; D-12, D-17), THE ModuloLibroTrabajo SHALL escribir las columnas por PCS, `Excused Event` y `Comments` (y B/C/D/E:I de PlantActivity si vienen) en la copia de trabajo **en la fila de `RawData-PCS` con el mismo timestamp** (el Excel une por fila, F-01), SHALL rechazar timestamps que no existan en `RawData-PCS`, y SHALL registrar en `correccion_dato` cada celda cuyo valor cambie respecto al libro base.
9. WHEN se ejecuta una corrida `reproceso` con un tramo corregido de RawData-PCS (D-13), THE ModuloLibroTrabajo SHALL sobrescribir ese tramo solo en la nueva copia de trabajo, SHALL registrar cada celda cambiada (fila, timestamp, PCS, campo, valor anterior, valor nuevo, archivo origen) en `correccion_dato` y en `data/work/<corte>/cambios.csv`, y SHALL NOT modificar el libro base ni resultados de corridas anteriores.
10. THE libro maestro de referencia SHALL ser la versión **v1.1** (D-19): libro de septiembre + hoja `Exclusion_Matrix` + regla de la matriz en `cmdCalcAvailability` y `mcoCreateList` según `design.md §Libro maestro v1.1`; el resto de macros, hojas y celdas de septiembre SHALL quedar sin cambios. El pipeline SHALL NOT modificar VBA.

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

**User Story:** Como desarrollador del motor, quiero asociar a cada fila el factor operacional de PlantActivity y los eventos de exclusión de `Exclusion_Matrix`, alineados por fila como lo hace el VBA.

#### Acceptance Criteria

1. THE ModuloEnriquecimiento SHALL asociar a cada fila de `RawData-PCS` el valor de `PlantActivity!C` (FactorOperacional) de la **misma fila origen** (F-01). `PlantActivity!D` SHALL conservarse solo como espejo de `Calc!BO` y SHALL NOT ponderar (F-37).
2. IF la celda de `PlantActivity` está vacía o la fila no existe, THEN THE ModuloEnriquecimiento SHALL usar `0` (semántica de celda vacía en Excel) y registrar la anomalía.
3. WHEN la fila de `PlantActivity` tiene timestamp distinto del de `RawData-PCS`, THE ModuloEnriquecimiento SHALL registrar la desalineación con ambos valores.
4. THE ModuloEnriquecimiento SHALL producir para cada `(NumeroFilaOrigen, NumeroPCS)`: `ModulosDisponibles`, `ModulosDisponiblesNulo`, `FallaRaw`, `FactorOperacional`, `ValorExclusion`, `BateriasPrevias`.
5. THE ModuloEnriquecimiento SHALL asociar `Exclusion_Matrix` **por fila** (misma fila Excel que `RawData-PCS`, columna `k+1` = PCS `k`); IF la hoja no existe, THEN todos los valores SHALL ser 0 con anomalía `exclusion_matrix_ausente` (info) (F-37).
6. IF una celda por PCS de `Exclusion_Matrix` contiene un valor distinto de 0, 1, 2 o vacío, o el encabezado no es `Date/time`, `PCS01`…`PCSnn`, THEN THE ModuloEnriquecimiento SHALL rechazar la corrida indicando fila y PCS.
7. THE ModuloEnriquecimiento SHALL calcular `BateriasPrevias` por tramo contiguo de valor 2 de cada PCS usando la fila anterior al tramo en la hoja completa: `C3 − M` si esa fila tenía `M < C3` y valor 0; `0` si tenía todos los módulos, estaba vacía o tenía valor 1; y SHALL reportar `ee2_sin_fila_previa` y `ee2_difiere_macro_agosto` (F-37).
8. THE ModuloEnriquecimiento SHALL conservar las columnas `Excused Event` (resumen por fila) y `Comments` (causa del EE) para trazabilidad, y reportar filas de la matriz desalineadas o con marcas sin timestamp.

---

### Requirement 7: Motor de disponibilidad (`cmdCalcAvailability`)

**User Story:** Como analista, quiero que el motor Python calcule C12, C14, C16 y C19 exactamente como la macro, para compararlos antes de retirar el Excel.

#### Acceptance Criteria

1. THE MotorDisponibilidad SHALL recorrer las filas en orden de fila origen y procesar una fila si `serial ≥ serial(inicio_periodo)` AND `serial < serial(fin_periodo) + 1`.
2. THE MotorDisponibilidad SHALL incrementar `BloquesMuestreo` en 1 por **cada fila** procesada, independientemente del número de PCS y aunque existan timestamps repetidos.
3. THE MotorDisponibilidad SHALL considerar un PCS indisponible en una fila únicamente si `ModulosDisponiblesNulo = False` AND `ModulosDisponibles < 4` (umbral literal del VBA, parametrizado como `baterias_por_pcs`).
4. WHEN un PCS es indisponible, THE MotorDisponibilidad SHALL calcular `BateriasIndisponibles = baterias_por_pcs − ModulosDisponibles` como `float`.
5. WHERE `aplicar_evento_excusable` (C31 = "Yes"), THE MotorDisponibilidad SHALL calcular `BateriasIndisponiblesPonderadas` según `Exclusion_Matrix` (F-37): `BateriasPrevias` si `ValorExclusion = 2`; `BateriasIndisponibles × (1 − ValorExclusion)` si es 0 o 1; en otro caso `= BateriasIndisponibles`.
6. WHERE `solo_tiempo_operacional`, THE MotorDisponibilidad SHALL acumular `C14 += racks_por_pcs × BateriasIndisponiblesPonderadas × FactorOperacional`; en otro caso `C14 += racks_por_pcs × BateriasIndisponiblesPonderadas`.
7. THE MotorDisponibilidad SHALL acumular C14 en el mismo orden que el VBA (fila, luego PCS 1..N) y sin redondeos intermedios.
8. THE MotorDisponibilidad SHALL calcular `DisponibilidadPeriodo = 1 − C14 / (TotalRacks × C12)` y `DisponibilidadAnualAcumulada = 1 − C14 / (TotalRacks × 365 × 24 × 4)`; IF el denominador es 0, THEN SHALL producir `None` (equivalente a `IFERROR(…,"N/A")`).
9. THE MotorDisponibilidad SHALL derivar `MinutosMuestreoDerivado = ROUND((serial₂ − serial₁) × 1440, 2)` de las dos primeras filas procesadas (C23) y reportar si difiere de `minutos_muestreo` (F-13).
10. THE MotorDisponibilidad SHALL producir una fila de `availability_sample_result` por `(NumeroFilaOrigen, NumeroPCS)` procesado con: `ModulosDisponibles`, `ModulosDisponiblesNulo`, `BateriasIndisponibles`, `ValorExclusion`, `FactorOperacional`, `BateriasIndisponiblesPonderadas`, `ImpactoRackPonderado`.

---

### Requirement 8: Motor de eventos de falla (`mcoCreateList`)

**User Story:** Como analista, quiero la lista de eventos idéntica a `ListOfFaults`, incluyendo el fallback de descripción y los defectos conocidos, para reconciliar evento por evento.

#### Acceptance Criteria

1. THE MotorEventosFalla SHALL recorrer los PCS en el orden de los bloques de encabezado de `RawData-PCS` y, para cada PCS, todas las filas en orden de fila origen, usando los **parámetros de eventos** (`L2`, `L4`, `L14`).
2. THE MotorEventosFalla SHALL considerar una fila si `serial ≥ L2` AND `serial < L4 + 1` AND `ModulosDisponiblesNulo = False` AND `ModulosDisponibles < 4`.
3. WHEN considera una fila, THE MotorEventosFalla SHALL acumular en `sumablocks` las baterías ponderadas por `Exclusion_Matrix` con la misma regla del KPI (R7.5) si `aplicar_evento_excusable_eventos` (L14 = "Yes"; D-16), o `sumablocks = (sumablocks + 4) − Modulos` en otro caso (orden de evaluación del VBA), y `numBlock += 1`.
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

1. *(rev. 3)* THE ModuloAgregacion SHALL mantener `disponibilidad_mensual` con **una sola fila** por `(IdProyecto, Anio, Mes)`: `BloquesMuestreo` (C12), `BloquesRacksIndisponibles` (C14), `DisponibilidadMensual` (C16), `DisponibilidadAnualAcumulada` (C19), `HorasRackEventos` (L10), `DiasMes`, `UltimoDato`, `MesCompleto`, `EstadoExclusiones`, `TipoCorrida`, `Origen` (`corrida` | `excel_manual`), `DisponibilidadContractual` y `NumCorrida` (NULL si `excel_manual`), reemplazada en cada carga del mes.
2. WHEN una corrida marcada como oficial cubre un mes, THE ModuloAgregacion SHALL registrar para ese mes `BloquesRacksIndisponibles = C14` y `BloquesMuestreo = C12` (**intervalos existentes**, D-07); `DiasMes` (informativo) = `serial(última fila en rango) − serial(primer día del mes)` si el mes está incompleto, o los días calendario si está completo. Esto se aparta a propósito de `Annual_AVA!E` (bloques de calendario: 1977 vs `C12` = 1975 en sep-2026) y hace que `DisponibilidadMensual` coincida con `C16`.
3. THE ModuloAgregacion SHALL calcular por mes `DisponibilidadMensual = 1 − Indisp / (TotalRacks × Bloques)` y los acumulados `BloquesMuestreoAcumulados`, `BloquesIndisponiblesAcumulados` y `DisponibilidadAcumulada = 1 − IndispAcum / (TotalRacks × BloquesAcum)` desde `mes_inicio_acumulado` (D-06). *(rev. 3)* En SQL el acumulado se expone como la vista `v_disponibilidad_anual` sobre `disponibilidad_mensual` (sin tabla por corrida); el motor Python lo sigue calculando para reconciliar con `Annual_AVA`.
4. THE ModuloAgregacion SHALL permitir importar como `Origen = excel_manual` las filas históricas de `Annual_AVA` que no se pueden reproducir (jul/ago 2026) (F-20).
5. THE ModuloAgregacion SHALL registrar `DisponibilidadContractual` (0,98 por defecto) por mes.
6. THE MotorETL SHALL documentar que C19 (fórmula 365 días sobre el período) y `DisponibilidadAcumulada` (Annual_AVA!J) son métricas distintas y persistir ambas.

---

### Requirement 11: Persistencia SQL Server — estado vigente por mes *(rev. 3, ADR-12)*

**User Story:** Como administrador y como consumidor de Power BI, quiero que cada mes exista **una sola vez** en la base: al volver a cargarlo, sus datos se reemplazan en vez de duplicarse, y cada ejecución queda registrada.

#### Acceptance Criteria

1. THE ModuloPersistencia SHALL guardar el estado vigente en `muestra_pcs`, `muestra_planta`, `detencion`, `disponibilidad_diaria`, `disponibilidad_mensual` y `calidad_dato`, sin `IdCorrida` en sus claves; cada fila SHALL guardar `NumCorrida` de la carga que la escribió (D-26).
2. WHEN se carga un período, THE ModuloPersistencia SHALL procesarlo **por mes calendario**: el inicio SHALL ajustarse al día 1 (con aviso) y un rango de varios meses SHALL cargarse mes a mes (D-20).
3. WHEN se carga un mes, THE ModuloPersistencia SHALL reemplazar, en **una transacción** por mes y con bloqueo por proyecto, todo lo vigente de ese mes: SHALL comparar con lo vigente y registrar en `correccion_dato` cada dato crudo (SCADA o `Exclusion_Matrix`) que cambió, SHALL borrar las filas del mes y SHALL insertar las nuevas; IF algo falla, THEN SHALL hacer rollback completo y el estado anterior SHALL quedar intacto.
4. THE ModuloPersistencia SHALL actualizar `detencion` por su clave de negocio `(IdProyecto, NumeroPCS, FechaInicio, Ocurrencia)`: SHALL conservar `IdDetencion` de las detenciones que siguen existiendo, insertar las nuevas y eliminar las del mes que ya no aparecen (D-25).
5. THE ModuloPersistencia SHALL publicar (reemplazar el estado) solo cargas que terminen `success` (incluye `sin_referencia`); una carga `parity_failed` SHALL quedar registrada en `etl_run` y `reconciliation_result` sin tocar el estado, salvo `--publicar-aunque-no-cuadre`; las corridas `golden` y `--sin-bd` SHALL NOT publicar (D-22).
6. IF el último dato de la carga es anterior al último dato vigente del mes, THEN THE ModuloPersistencia SHALL detenerse salvo `--permitir-recorte` (D-21); IF el mes es `excel_manual`, THEN SHALL detenerse salvo `--reemplazar-manual`; IF el mes está vigente "Con Exclusiones" y la carga es "Sin Exclusiones", THEN SHALL detenerse salvo `--forzar-sin-exclusiones` (D-23).
7. THE ModuloPersistencia SHALL registrar cada ejecución en `etl_run` (append-only: `IdCorrida`, `NumCorrida`, `Estado` `running` → `success` | `failed` | `parity_failed`, parámetros efectivos, `ArchivoOrigen`, `HashArchivoOrigen`, `MinutosMuestreoDerivado`, `TipoCorrida`, `VersionAlgoritmo`, `IdProyecto`, meses publicados y conteo de filas reemplazadas, resumen de calidad y de reconciliación).
8. THE ModuloPersistencia SHALL usar inserción masiva (`fast_executemany`, ODBC Driver 18) y almacenar módulos, baterías y factores como `FLOAT`.
9. THE tablas de registro (`etl_run`, `correccion_dato`, `exclusion_matrix_carga`, `detencion_revision`, `excel_reference_run`, `reconciliation_result`) SHALL ser append-only; el rol `etl_writer` SHALL tener `DELETE`/`UPDATE` solo sobre las tablas de estado (y el `UPDATE` de estado de `etl_run` y de `exclusion_matrix_carga.IdCorridaCierre`).
10. THE claves de tiempo SHALL incluir `Ocurrencia` (1, 2…) para distinguir timestamps repetidos por el cambio de hora de abril (F-32).

---

### Requirement 12: Proyectos, catálogo y detenciones

**User Story:** Como analista, quiero consultar detenciones entre proyectos y en el tiempo, con revisión operacional que sobreviva a los reprocesos semanales.

#### Acceptance Criteria

1. THE ModuloPersistencia SHALL mantener `proyecto` con Arena (`IdProyecto=1`, `en_ejecucion`), Copiapó A (2), Luz del Norte (3) y María Elena (4) (`por_implementar`).
2. THE tabla `proyecto` SHALL almacenar `NumPCS`, `NumBateriasPorPCS`, `NumRacksPorBAC`, `TotalRacks` (calculado), `MinutosMuestreo`, `ZonaHoraria`, `FechaInicio`.
3. THE ModuloPersistencia SHALL cargar `tipo_detencion` desde la hoja `PCS-Fault` completa (**167 códigos**, F0…F257) con `CodigoFalla`, `DescripcionFallaPE`, `CodigoDescripcion`, `Significado` y `Operativo` (F-19); el seed SHALL ser idempotente y generado por script.
4. *(rev. 3)* WHEN se persiste un evento, THE ModuloPersistencia SHALL insertar o actualizar (R11.4) una fila en `detencion` —que reúne los campos de `ListOfFaults` (antes `fault_event`)— con `IdProyecto`, `NumCorrida`, `NumeroPCS`, `FechaInicio`, `FechaTermino`, `DuracionSegundos`, `DuracionHoras` (misma duración en horas, = `fault_event.DuracionHoras`), `IdTipoDetencion` (NULL + advertencia si el código no existe), `CodigoFalla`, `DescripcionFalla`, flags de calidad y rack-hours.
5. THE Sistema SHALL guardar el workflow (`EstadoRevision` ∈ {pendiente, revisado, excluido}, `Observacion`, `RevisadoPor`, `RevisadoEn`) en `detencion_revision`, con clave de negocio `(IdProyecto, NumeroPCS, FechaInicio)`, para que sobreviva a nuevas corridas (F-27).
6. *(rev. 3)* THE Sistema SHALL exponer `v_detencion`: las detenciones vigentes unidas con su última revisión, y `v_resumen_codigo_mensual`: horas-rack por código de falla y mes (antes `fault_code_summary`).
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

1. *(rev. 3)* THE MotorETL SHALL garantizar la trazabilidad `disponibilidad_mensual → muestra_pcs (aporte y dato crudo, NumeroFilaOrigen) → etl_run (archivo, hash) de la carga vigente`; los valores reemplazados SHALL quedar en `correccion_dato`.
2. *(rev. 3)* THE MotorETL SHALL persistir las anomalías vigentes del mes en `calidad_dato` (tipo, severidad, fila, PCS, timestamp, detalle; reemplazadas en cada carga del mes) y un resumen JSON por ejecución en `etl_run.ResumenCalidad`.
3. THE resumen de calidad SHALL incluir: filas con `ModulosDisponiblesNulo`, eventos con `DescripcionFallaFallback`, eventos `EventoArrastradoExcel` y `ExcelHabriaFallado`, anomalías de timestamp, filas truncadas por `modo_huecos`, desalineaciones y vacíos de PlantActivity, PCS/columnas faltantes, diferencia C23 vs config y descripciones numéricas.
4. THE MotorETL SHALL hacer consultables en todo el estado vigente los intervalos `ModulosDisponiblesNulo = True` (`muestra_pcs`), sin suprimirlos.

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

1. WHEN `run-etl` y `reconcile` terminan, THE OrquestadorLunes SHALL emitir una notificación con `NumCorrida`, meses publicados, período, estado, etiqueta de exclusiones ("Sin Exclusiones" / "Con Exclusiones", R19.5) y resumen de reconciliación y de calidad.
2. THE destino SHALL configurarse por variable de entorno (webhook); IF no está configurado, THEN THE OrquestadorLunes SHALL escribir la notificación en `data/work/<corte>/notificacion.md` y en consola.
3. THE Sistema SHALL NOT intentar refresh vía API de Power BI hasta que exista service principal (fuera de alcance v1).
4. *(rev. 3)* THE Sistema SHALL ofrecer a Power BI las tablas de estado (`disponibilidad_mensual`, `disponibilidad_diaria`, `detencion`) y las vistas `v_disponibilidad_anual`, `v_resumen_codigo_mensual` y `v_detencion`, con una fila por mes, día o detención, sin exponer corridas.

---

### Requirement 19: Orquestador semanal

**User Story:** Como operador, quiero ejecutar el flujo del lunes por etapas o completo, reanudable ante fallas.

#### Acceptance Criteria

1. *(rev. 3)* THE OrquestadorLunes SHALL ofrecer `--stage acquire-wait | prepare-workbook | run-macros | run-etl | reconcile | notify-bi | all` (semanal), `--stage cierre-mensual` y `--stage load-exclusion-matrix` (mensuales, D-12, D-17) y los argumentos `--inbox`, `--work`, `--desde`, `--hasta`, `--entorno`, `--permitir-recorte`, `--reemplazar-manual`, `--forzar-sin-exclusiones`, `--publicar-aunque-no-cuadre`; `--oficial` SHALL aceptarse sin efecto (con aviso) y `--reproceso` SHALL reemplazarse por la recarga del mes (D-27).
2. THE OrquestadorLunes SHALL guardar el estado de cada etapa en `data/work/<corte>/run_state.json` y permitir reanudar desde la etapa fallida sin repetir las exitosas.
3. WHEN no se informa el período, THE OrquestadorLunes SHALL calcular el período por defecto según D-07: corrida `semanal` con `C5` = día 1 del mes del último dato cargado y `C7` = fecha del último dato (`L2`/`L4`/`Daily!D5` iguales; `C21 = "No"`, `C31 = L14 = "Yes"`); y, si los datos cargados ya cubren el último bloque de un mes (último día 23:45) sin corrida oficial `cierre_mensual`, SHALL encolar además la corrida `cierre_mensual` de ese mes (día 1 → último día).
4. THE OrquestadorLunes SHALL registrar logs estructurados por etapa con timestamps y terminar con código de salida ≠ 0 ante error.
5. THE OrquestadorLunes SHALL registrar en cada corrida `EstadoExclusiones`: `sin_exclusiones` si la `Exclusion_Matrix` del mes del período aún no se cargó (`exclusion_matrix_carga`), o `con_exclusiones` si ya se cargó (D-17). Las corridas `semanal` "Sin Exclusiones" SHALL ser **oficiales** (no preliminares) y la notificación SHALL mostrar la etiqueta "Sin Exclusiones" / "Con Exclusiones".
6. WHEN se ejecuta `load-exclusion-matrix` para un mes (Alex la entrega una vez al mes, al final), THE OrquestadorLunes SHALL registrar la carga en `exclusion_matrix_carga` y ejecutar el `cierre_mensual` oficial `con_exclusiones`; THE OrquestadorLunes SHALL NOT ejecutar un `cierre_mensual` oficial sin la matriz del mes cargada y alineada por fila con `RawData-PCS!A`, salvo `--sin-exclusiones` explícito (F-37, D-17).
7. *(rev. 3)* THE `disponibilidad_mensual` SHALL reflejar la última carga publicada del mes; una carga "Sin Exclusiones" SHALL NOT reemplazar un mes vigente "Con Exclusiones" (R11.6).
8. *(rev. 3)* WHEN se informa un período, THE OrquestadorLunes SHALL ajustar el inicio al día 1 del mes y, si abarca varios meses, SHALL ejecutar la cadena (macros, ETL, reconciliación, publicación) mes a mes, informando cada mes publicado (D-20).

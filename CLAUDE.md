# CLAUDE.md

Este archivo entrega contexto a los agentes de IA (Claude Code, Kiro, etc.) para trabajar en este repositorio.

---

## Estado actual del proyecto

Este repositorio tiene como libro fuente de cálculo:

```
data/AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm
```

La data cruda semanal llega desde el **server SCADA** a `data/inbox/` (hoy vía TeamViewer); ver §4.2 Flujo semanal.

**El objetivo ya no es mantener el Excel** — el objetivo es reemplazarlo progresivamente por un proceso reproducible en **Python + SQL Server**.

El plan completo de reingeniería está en [`AGENTS.md`](./AGENTS.md). Ese documento describe el plan de implementación; leerlo antes de cualquier tarea de desarrollo.

El spec ejecutable (revisión 2, auditado contra el VBA real el 2026-09-24) está en `.kiro/specs/etl-arena-availability/`: `audit.md` (hallazgos `F-xx` y decisiones `D-xx`), `requirements.md`, `design.md` y `tasks.md` (fases, olas y agente asignado a cada tarea). Ante diferencias, manda el spec revisión 2.

---

## Qué hace el Excel hoy

Herramienta de cálculo de KPI (equipo Trina Solar Chile / TS-ESD) que calcula la **disponibilidad** de unidades PCS (Power Conversion System) y racks de baterías a partir de registros crudos de fallas/estado, usando macros VBA sobre fórmulas de Excel.

### Contexto del activo

| Dato | Valor |
|---|---|
| Nombre del proyecto | Arena BESS |
| Fecha de inicio de operación | 08/Abril/2026 |
| Total PCS | 61 |
| Módulos de batería BEC por PCS | 4 |
| Total de baterías BEC | 61 × 4 = 244 |
| Racks por BAC | 12 |
| Total de racks | 61 × 4 × 12 = 2.928 |

### Parámetros clave

| Parámetro | Celda/nombre Excel | Valor observado |
|---|---|---:|
| Total PCS | `C2` / `Total_PCS` | 61 |
| Baterías por PCS | `C3` / `Total_Batteries_per_PCS` | 4 |
| Total racks | `C11` / `Total_Racks` | 2928 |
| Frecuencia de muestreo | `C23` / `Frecuencia_de_muestreo__min` | 15 min |
| Fecha inicio | `C5` | fecha seleccionada |
| Fecha fin | `C7` | fecha seleccionada |
| Solo tiempo operacional | `C21` (`solo_tiempo_operacional`) | Yes/No |
| Aplicar evento excusable | `C31` (`aplicar_evento_excusable`) | Yes/No — activa `Exclusion_Matrix` (F-37) |

### Fórmula del KPI principal

```
DisponibilidadPeriodo = 1 - BloquesRacksIndisponibles / (TotalRacks * BloquesMuestreo)
```

donde:
- `BloquesMuestreo` (C12) = cantidad de bloques de 15 min en el período seleccionado
- `BloquesRacksIndisponibles` (C14) = acumulado de `(racks indisponibles) × bloques`, opcionalmente ponderado por `Exclusion_Matrix` (C31) y `FactorOperacional` (C21)
- `DisponibilidadAnualAcumulada` (C19) = `1 - BloquesRacksIndisponibles / (TotalRacks * 365 * 24 * 4)`

---

## Arquitectura del libro Excel

### Hojas (11 visibles)

| codeName VBA      | Nombre de pestaña          | Rol                                                         |
|-------------------|----------------------------|-------------------------------------------------------------|
| `cmdAvailability` | Calculation-Availability   | Panel de control: parámetros + tabla de resultados          |
| `Sheet1`          | DateFormat_Correction      | Normalización de fechas/horas (~5 MB de XML)                |
| `Sheet2`          | RawData-PCS                | Registros crudos de estado/falla por intervalo y PCS        |
| `Sheet3`          | ListOfFaults               | Lista de eventos generada por `mcoCreateList`               |
| `Sheet4`          | PlantActivity              | `FactorOperacional` (col C) por intervalo; col D ya no pondera (F-37) |
| —                 | Exclusion_Matrix           | Eventos de exclusión 0/1/2 por intervalo y PCS (desde agosto 2026, F-37) |
| `Sheet5`          | PCS-Fault                  | Catálogo de 167 códigos de falla F0…F257 (TipoDetencion)   |
| `Sheet6`          | PCS-Status                 | Catálogo de estados de PCS                                  |
| `Sheet7`          | Verificación               | Controles de calidad                                        |
| `Sheet8`          | Graph                      | Datos fuente de gráficos + salida de `Graphupdate`          |
| `Sheet9`          | Daily                      | `Disponibilidad` acumulada diaria                           |
| `Sheet10`         | Annual_AVA                 | `DisponibilidadAcumulada` anual                             |

### Módulos VBA

**Module1**
- `cmdCalcAvailability` — macro principal. Limpia resultados anteriores, lee rango de fechas y parámetros, recorre `RawData-PCS` fila por fila y acumula `ImpactoRackPonderado` en `BloquesRacksIndisponibles` (C14) y contador de `BloquesMuestreo` (C12).
- `mcoCleanTable` — limpia `Calculation-Availability!E4:BO10000` y `C25:C29` antes de cada corrida.
- `mcoCreateList` — construye `ListOfFaults`: agrupa intervalos consecutivos con `NUMBER_OF_MODULES < 4` en `EventoFalla` discretos (inicio/fin/`DuracionHoras`/`CodigoFalla`/`DescripcionFalla`/`HorasRackIndisponibles`).
- `mcoDailyAvailability` — acumula la tabla de resultados en totales diarios en `Daily` (`Disponibilidad`, `Variacion`).

**Module2**
- `mcoCleanList` — limpia `ListOfFaults` antes de reconstruirla.
- `mcoOrder` — ordena la tabla resumen por código de falla `ListOfFaults!N5:Q172` (columna P = Σ rack-hours por código) de mayor a menor. La lista de eventos `B:I` no se ordena.
- `mcoTestFormulae`, `mcoTests2` — macros de prueba, no son parte del flujo productivo.

**Module3**
- `Graphupdate` — capa de presentación; ordena y copia datos hacia áreas fijas para los gráficos. No forma parte del MotorDisponibilidad.

**Orden típico de ejecución:**
```
parámetros -> cmdCalcAvailability -> mcoCreateList -> mcoDailyAvailability -> Graphupdate
```

### Estructura de columnas en RawData-PCS

Bloques de 4 columnas por PCS. El campo que gobierna el cálculo es `NUMBER_OF_MODULES` (columna `4 * pcs_number + 1`):

```
PCS 1 -> columna E (índice 5)
PCS 2 -> columna I (índice 9)
PCS 3 -> columna M (índice 13)
...
```

Los nombres de columna siguen el patrón `Arena - PCS XX - ...` donde `XX` va de `01` a `61`.

Nombres exactos de las 4 columnas por PCS:
```
Arena - PCS XX - POWERELECTRONICS GEN3 HEx CURRENT FAULT
Arena - PCS XX - POWERELECTRONICS GEN3 HEx CURRENT STATUS
Arena - PCS XX - POWERELECTRONICS GEN3 HEx CURRENT WARNING
Arena - PCS XX - POWERELECTRONICS HEM-k NUMBER OF MODULES
```

Campo clave: **`Arena - PCS XX - POWERELECTRONICS HEM-k NUMBER OF MODULES`** -> `ModulosDisponibles`

Un PCS es indisponible cuando `ModulosDisponibles < 4`. El valor puede ser **fraccionario** (p. ej. 3,2347): se conserva como `float`, nunca se trunca. Si `NUMBER_OF_MODULES` está **vacío**, el Excel lo ignora (trata el PCS como disponible); Python debe replicar ese comportamiento y marcar el registro con `ModulosDisponiblesNulo = True` para auditoría.

Trampas de paridad confirmadas en el VBA (detalle en `audit.md`):
- `PlantActivity` se une **por número de fila**, no por timestamp; celda vacía = 0.
- `mcoCreateList` usa sus propios parámetros `ListOfFaults!L2`/`L4`/`L14`, independientes de `C5`/`C7`/`C31`.
- El fin de un evento es la última fila en falla; con fila anterior vacía el fallback da `"F1 Watchdog"`.
- `mcoDailyAvailability` no aplica el factor operacional.
- `C12` cuenta filas (no timestamps únicos); los motores operan sobre seriales Excel en orden de fila origen.

**Eventos de exclusión (`Exclusion_Matrix`, definición de negocio 2026-09-24 — F-37, ADR-11):** una columna por PCS, unida por fila. Con `C31`/`L14 = "Yes"` y `M < 4`: `0`/vacío → `4 − M`; `1` → 0 (se consideran los 4 módulos); `2` → baterías indisponibles previas al inicio del evento. Reemplaza a `PlantActivity!D`; PlantActivity queda solo para "Only Operational Time?" (C21). Macros, hojas y celdas oficiales: las de septiembre, más la regla de la matriz en `cmdCalcAvailability` y `mcoCreateList` (maestro **v1.1**, D-19; ver `design.md §Libro maestro v1.1`). El pipeline nunca edita VBA.

No usar el código de falla (`CURRENT FAULT`) como criterio de indisponibilidad — el comentario VBA es impreciso; la lógica ejecutable usa `NUMBER_OF_MODULES`.

---

## Cómo inspeccionar el archivo Excel desde línea de comandos

Extraer código VBA (requiere `oletools`):

```bash
pip install oletools
python -m oletools.olevba "data/AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm"
```

Inspeccionar estructura XML sin Excel:

```bash
unzip -l  "data/AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm"
unzip -p  "data/AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm" xl/workbook.xml
```

> Las hojas `DateFormat_Correction` y `PlantActivity` pesan ~5 MB de XML cada una.
> Usar `unzip -p ... | grep` en lugar de leerlas completas.

---

## 4.2 Flujo semanal (lunes) — SCADA → ETL → SQL → PBI

La fuente cruda de `RawData-PCS` es el **server SCADA**. Cada lunes se exporta un reporte y se procesa en la PC local (no en el server).

```text
SCADA ~03:00 AM (solo extract)
  → TeamViewer → data/inbox/raw_pcs_<corte>.<ext>
  → acquire-wait + scada_adapter (fechas mm-dd → serial Excel, nunca texto; mapping RawData-PCS)
  → copia de trabajo del .xlsm (backup)
  → macros COM: cmdCalcAvailability → mcoCreateList → mcoDailyAvailability → Graphupdate
  → extrae C12/C14/C16/C19 (referencia Excel)
  → pipeline Python → SQL Server (IdCorrida nuevo)
  → reconcile Python vs referencia
  → notificar a Misael: refresh Power BI
```

Base de datos (ADR-10, 2026-09-24): **Azure SQL Database serverless, oferta gratuita** — servidor `trina-etl.database.windows.net` (Brazil South), autenticación Microsoft Entra ID (`ETL_ARENA_DB_AUTH=entra`, por defecto). **Dos ambientes fijos en `src/etl_arena/ambientes.py`:** **PROD** = `trina_etl` (por defecto) y **TEST/QA** = `trina_etl_prueba` (`--entorno prueba|qa|test` o `ETL_ARENA_ENTORNO`; usa además `data/work/_prueba`). Sin `.env` se usan esas direcciones; `ETL_ARENA_DB_URL` / `ETL_ARENA_DB_URL_PRUEBA` solo las reemplazan. El código nunca crea ni borra bases en Azure. Los tests de integración usan su propia base (`ETL_ARENA_TEST_DB_URL`, con "test" en el nombre); `reiniciar_esquema` se niega a tocar `trina_etl` y `trina_etl_prueba` (incidente 2026-09-25: los tests vaciaron TEST/QA).

Orquestador: `scripts/run_lunes.py` con etapas `acquire-wait`, `prepare-workbook`, `run-macros`, `run-etl`, `reconcile`, `notify-bi`, más `cierre-mensual` y `load-exclusion-matrix` (mensuales). Entre cortes (`etl_arena.orquestacion`): libro base promovido tras cada corrida oficial exitosa y cola de cierres mensuales. Fases P0–P9 y detalle en `AGENTS.md` §14 Fase S. Runbook: `docs/runbook-lunes.md`. Power BI: `docs/pbi-handoff.md`.

**Macros y exclusiones (F-44):** las macros de septiembre con C31/L14 = "Yes" excusan con `PlantActivity!D`; por eso `workbook.macros` escribe "No" salvo que el VBA lea `Exclusion_Matrix` (maestro v1.1, detectado con `workbook.vba`), y si el período tiene exclusiones la referencia Excel se omite.

**Velocidad de las macros (F-45):** `workbook.macros` corre por defecto en modo rápido: pantalla sin refrescar y recálculo manual, con `Calculate` antes de cada macro y al final (resultado idéntico, ~2× más rápido; el VBA no se toca). `run_lunes.py --macros-sin-optimizar` vuelve al modo antiguo.

**Reglas:** export **incremental** (desde el dato siguiente al último cargado hasta el último dato del lunes; continuidad validada al recibir; D-07); KPI semanal = mes en curso hasta el último dato; cierre mensual = mes completo; KPI oficial con `C31 = L14 = "Yes"` (D-03); transporte hoy solo TeamViewer (sin UNC/API); solo `RawData-PCS` desde SCADA; `Exclusion_Matrix` la entrega Alex a fin de mes: KPI semanal oficial "Sin Exclusiones", cierre mensual oficial "Con Exclusiones" (D-17, F-37); PlantActivity solo para C21 (D-12); correcciones solo por `reproceso` con registro de celdas cambiadas (D-13); macros y ETL solo en PC local; Power BI modo notificación (owner Misael) hasta service principal.

---

## Arquitectura Python (ver AGENTS.md §18)

Paquete `etl_arena` con layout *src* (`src/etl_arena/<módulo>/`, ADR-01); se agregan `excel_semantics/` (reglas VBA/Excel), `model/` y `workbook/` (COM). Detalle en `design.md`.

```
src/etl_arena/
  config/           <- ConfiguracionCalculo, defaults, desde_excel
  excel_semantics/  <- reglas de celda/texto/fechas/redondeo de VBA-Excel
  model/            <- MatrizPCS, DatosActividad, Anomalia
  acquisition/      <- acquire_wait, contrato_scada, lector_scada (Fase S)
  workbook/         <- sesión COM, preparar, macros, referencia (solo PC local)
  ingestion/        <- lectura streaming del libro + anomalías de timestamp
  normalization/    <- validador de esquema, ancho→matriz→largo
  enrichment/       <- actividad_planta (join por fila)
  availability/     <- MotorDisponibilidad
  fault_events/     <- MotorEventosFalla, resumen por código
  aggregation/      <- diaria, mensual_anual
  persistence/      <- conexión, repositorio, corrida
  reconciliation/   <- niveles, invariantes, servicio, reporte
  reporting/        <- calidad, notificacion
  pipeline.py
scripts/  ejecutar_etl.py  run_lunes.py  shadow_log.py  verificar_entorno.py  crear_base.py  generar_seed_tipo_detencion.py
sql/      00_database.sql … 07_audit_queries.sql
tests/    unit/ property/ golden/ integration/ com/ fixtures/
data/     inbox/ processed/ work/
docs/     glosario.md  modelo-datos.md  runbook-lunes.md  instalacion.md  pbi-handoff.md  shadow-log.md  go-no-go.md  data-contract-*.md  adr/
```

Scripts principales: `scripts/run_lunes.py` (orquestador semanal) y `scripts/ejecutar_etl.py` (pipeline de corrida).

---

## Tablas SQL Server (ver AGENTS.md §13)

| Tabla | Columnas clave |
|---|---|
| `etl_run` | `IdCorrida` (llave UUID), `NumCorrida` (correlativo 1, 2, 3… para personas; `IDENTITY_CACHE = OFF`), `IdProyecto`, parámetros de corrida (período, flags, `TotalPCS`…), `Status`, `MensajeError`, `VersionAlgoritmo`. Los KPI C12/C14 viven en `availability_run_result`, no aquí |
| `proyecto` | `IdProyecto`, `NumPCS`, `NumBateriasPorPCS`, `NumRacksPorBAC`, `TotalRacks`, `MinutosMuestreo`, `ZonaHoraria` |
| `tipo_detencion` | `IdTipoDetencion`, `CodigoFalla`, `DescripcionFallaPE`, `CodigoDescripcion`, `Significado`, `Operativo` (167 códigos de `PCS-Fault`) |
| `raw_pcs_sample` | `IdCorrida`, `MarcaTiempoMuestra`, `NumeroPCS`, `ModulosDisponibles`, `ModulosDisponiblesNulo`, `SerialFechaExcelOrigen` |
| `plant_activity_sample` | `IdCorrida`, `MarcaTiempoMuestra`, `EsOperacional`, `PorcentajeSOC` |
| `exclusion_matrix_sample` | `IdCorrida`, `NumeroFilaOrigen`, `NumeroPCS`, `ValorExclusion` (1/2), `BateriasPrevias`, `Comentario` (F-37) |
| `availability_sample_result` | `IdCorrida`, `MarcaTiempoMuestra`, `NumeroPCS`, `BateriasIndisponibles`, `ValorExclusion`, `FactorOperacional`, `ImpactoRackPonderado` |
| `availability_run_result` | `IdCorrida`, `BloquesMuestreo`, `BloquesRacksIndisponibles`, `DisponibilidadPeriodo`, `DisponibilidadAnualAcumulada` |
| `fault_event` | `IdCorrida`, `NumeroPCS`, `MarcaTiempoInicio`, `MarcaTiempoFin`, `DuracionHoras`, `CodigoFalla`, `DescripcionFallaFallback`, `HorasRackIndisponibles` |
| `detencion` | `IdDetencion`, `IdProyecto`, `IdCorrida`, `FechaInicio`, `FechaTermino`, `DuracionSegundos`, `DuracionHoras`, `IdTipoDetencion`, `EstadoRevision`, `Observacion` |
| `daily_availability` | `IdCorrida`, `Dia`, `BloquesRacksIndisponiblesDiarios`, `BloquesRacksIndisponiblesAcumulados`, `Disponibilidad`, `Variacion` |
| `annual_availability` | `IdCorrida`, `Anio`, `Mes`, `BloquesMuestreo`, `DisponibilidadMensual`, `DisponibilidadAcumulada` |

---

## Reglas importantes para agentes

1. **No modificar el algoritmo de disponibilidad** durante la fase de paridad — reproducirlo exactamente.
2. **No hardcodear** `61`, `4`, `12`, `15` — provienen de `ConfiguracionCalculo`.
3. **No destruir resultados históricos en SQL** — usar `IdCorrida` como clave; SQL es append-only.
4. **El campo que gobierna la indisponibilidad es `NUMBER_OF_MODULES`** (`ModulosDisponibles`), no el código de falla.
5. **Conservar timestamps originales** (`SerialFechaExcelOrigen` + `MarcaTiempoLocalOrigen`) — no convertir a UTC sin documentar.
6. **`ModulosDisponiblesNulo`**: intervalos con `NUMBER_OF_MODULES` vacío se tratan como disponibles (= 4) pero se marcan para auditoría.
7. **`DescripcionFallaFallback`**: eventos donde la descripción fue tomada del intervalo anterior se marcan en `fault_event` y `detencion`.
8. **`VersionAlgoritmo`**: toda corrida registra `"availability-v1.1-exclusion-matrix"` (F-37); sin `Exclusion_Matrix` equivale a `"availability-v1-excel-parity"`.
9. **No procesar en el server SCADA** — solo exportar/copy; macros y ETL solo en PC local.
10. **`data/processed` inmutable** — el archivo de origen con sha256 no se edita; el `.xlsm` de trabajo se copia/backup antes de macros o adapter.
11. La lógica de negocio detallada y el plan por etapas están en [`AGENTS.md`](./AGENTS.md); las reglas ejecutables, tolerancias y tareas vigentes están en `.kiro/specs/etl-arena-availability/` (revisión 2), que manda ante diferencias.

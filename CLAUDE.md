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

El plan completo de reingeniería está en [`AGENTS.md`](./AGENTS.md). Ese documento es la **fuente de verdad** del plan de implementación. Leerlo antes de cualquier tarea de desarrollo.

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
| Aplicar evento excusable | `C31` (`aplicar_evento_excusable`) | Yes/No |

### Fórmula del KPI principal

```
DisponibilidadPeriodo = 1 - BloquesRacksIndisponibles / (TotalRacks * BloquesMuestreo)
```

donde:
- `BloquesMuestreo` (C12) = cantidad de bloques de 15 min en el período seleccionado
- `BloquesRacksIndisponibles` (C14) = acumulado de `(racks indisponibles) × bloques`, opcionalmente ponderado por `FactorExcusable` y `FactorOperacional`
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
| `Sheet4`          | PlantActivity              | `FactorOperacional` (col C) y `FactorExcusable` (col D) por intervalo |
| `Sheet5`          | PCS-Fault                  | Catálogo de 68 códigos/descripciones de falla (TipoDetencion)|
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
- `mcoOrder` — ordena `ListOfFaults` por `HorasRackIndisponibles` (columna P) de mayor a menor.
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

Un PCS es indisponible cuando `ModulosDisponibles < 4`. Si `NUMBER_OF_MODULES` está **vacío**, el Excel lo ignora (trata el PCS como disponible); Python debe replicar ese comportamiento y marcar el registro con `ModulosDisponiblesNulo = True` para auditoría.

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
  → acquire-wait + scada_adapter (fechas mm-dd→dd-mm, mapping RawData-PCS)
  → copia de trabajo del .xlsm (backup)
  → macros COM: cmdCalcAvailability → mcoCreateList → mcoDailyAvailability → Graphupdate
  → extrae C12/C14/C16/C19 (referencia Excel)
  → pipeline Python → SQL Server (IdCorrida nuevo)
  → reconcile Python vs referencia
  → notificar a Misael: refresh Power BI
```

Orquestador: `scripts/run_lunes.py` con etapas `acquire-wait`, `prepare-workbook`, `run-macros`, `run-etl`, `reconcile`, `notify-bi`. Fases P0–P9 y detalle en `AGENTS.md` §14 Fase S. Runbook: `docs/runbook-lunes.md`.

**Reglas:** fechas del reporte las setea quien exporta (convención 01-01-2026 → último domingo, validada al recibir); transporte hoy solo TeamViewer (sin UNC/API); solo `RawData-PCS` desde SCADA (`PlantActivity` aparte); macros y ETL solo en PC local; Power BI modo notificación (owner Misael) hasta service principal.

---

## Arquitectura Python (ver AGENTS.md §18)

```
src/
  config/           <- ConfiguracionCalculo, CONFIG_POR_DEFECTO
  acquisition/      <- acquire_wait, scada_adapter (Fase SCADA)
  excel_macro/      <- runner COM de las 4 macros VBA (Fase M, solo PC local)
  ingestion/        <- ServicioIngesta, ReporteAnomalia
  staging/          <- RepositorioStaging
  normalization/    <- NormalizadorPCS
  enrichment/       <- ServicioEnriquecimiento
  availability/     <- MotorDisponibilidad, ResultadoDisponibilidad
  fault_events/     <- MotorEventosFalla, EventoFalla
  aggregation/      <- AgregacionDiaria, AgregacionAnual
  persistence/      <- ServicioPersistencia
  reconciliation/   <- ServicioReconciliacion, ReporteReconciliacion
  reporting/        <- reporte_calidad.py
data/
  inbox/            <- drop zone semanal del export SCADA
  processed/        <- originales inmutables + sha256
  work/             <- copia de trabajo del .xlsm
scripts/run_lunes.py
docs/runbook-lunes.md
tests/
sql/
```

Scripts principales: `scripts/run_lunes.py` (orquestador semanal) y `ejecutar_etl.py` (pipeline de corrida).

---

## Tablas SQL Server (ver AGENTS.md §13)

| Tabla | Columnas clave |
|---|---|
| `etl_run` | `IdCorrida`, `IdProyecto`, parámetros de corrida (período, flags, `TotalPCS`…), `Status`, `MensajeError`, `VersionAlgoritmo`. Los KPI C12/C14 viven en `availability_run_result`, no aquí |
| `proyecto` | `IdProyecto`, `NumPCS`, `NumBateriasPorPCS`, `NumRacksPorBAC`, `TotalRacks`, `MinutosMuestreo`, `ZonaHoraria` |
| `tipo_detencion` | `IdTipoDetencion`, `CodigoFalla`, `DescripcionFallaPE`, `CodigoDescripcion` |
| `raw_pcs_sample` | `IdCorrida`, `MarcaTiempoMuestra`, `NumeroPCS`, `ModulosDisponibles`, `ModulosDisponiblesNulo`, `SerialFechaExcelOrigen` |
| `plant_activity_sample` | `IdCorrida`, `MarcaTiempoMuestra`, `EsOperacional`, `EsEventoExcusable`, `PorcentajeSOC` |
| `availability_sample_result` | `IdCorrida`, `MarcaTiempoMuestra`, `NumeroPCS`, `BateriasIndisponibles`, `FactorExcusable`, `FactorOperacional`, `ImpactoRackPonderado` |
| `availability_run_result` | `IdCorrida`, `BloquesMuestreo`, `BloquesRacksIndisponibles`, `DisponibilidadPeriodo`, `DisponibilidadAnualAcumulada` |
| `fault_event` | `IdCorrida`, `NumeroPCS`, `MarcaTiempoInicio`, `MarcaTiempoFin`, `DuracionHoras`, `CodigoFalla`, `DescripcionFallaFallback`, `HorasRackIndisponibles` |
| `detencion` | `IdDetencion`, `IdProyecto`, `IdCorrida`, `FechaInicio`, `FechaTermino`, `DuracionSegundos`, `IdTipoDetencion`, `EstadoRevision`, `Observacion` |
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
8. **`VersionAlgoritmo`**: toda corrida debe registrar `"availability-v1-excel-parity"` durante la fase de paridad.
9. **No procesar en el server SCADA** — solo exportar/copy; macros y ETL solo en PC local.
10. **`data/processed` inmutable** — el archivo de origen con sha256 no se edita; el `.xlsm` de trabajo se copia/backup antes de macros o adapter.
11. Toda la lógica de negocio detallada, fórmulas exactas y plan por etapas están en [`AGENTS.md`](./AGENTS.md).

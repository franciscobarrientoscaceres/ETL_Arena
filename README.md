# ETL_Arena — Disponibilidad PCS & Baterías

Proyecto de reingeniería del cálculo de disponibilidad de unidades PCS (Power Conversion System) y racks de baterías, actualmente implementado en Excel/VBA, hacia un proceso reproducible y auditable en **Python + SQL Server**.

> **Estado:** Etapa 1 completada (reverse engineering). En preparación: data contract, esquema SQL y **cadena semanal SCADA** (adquisición + macros + ETL + SQL + handoff PBI).

---

## Contexto

El proyecto **Arena BESS**, operativo desde el **08/Abril/2026**, requiere calcular mensualmente la **disponibilidad** de 61 unidades PCS (Power Conversion System), sus 244 baterías BEC (61 PCS × 4 módulos BEC) y 2.928 racks (61 × 4 × 12). Actualmente el equipo TS-ESD (Trina Solar Chile) realiza ese cálculo con un libro Excel con macros VBA. El resultado es el KPI contractual que determina si el activo cumple su compromiso de disponibilidad.

El objetivo de este proyecto es reemplazar ese proceso de forma progresiva y controlada, garantizando que el resultado Python sea idéntico al Excel antes de retirar el Excel como fuente oficial.

---

## KPI principal

```
DisponibilidadPeriodo = 1 - BloquesRacksIndisponibles / (TotalRacks × BloquesMuestreo)
```

| Variable | Equivalente Excel | Significado |
|---|---|---|
| `BloquesRacksIndisponibles` | C14 | Acumulado de `(racks indisponibles) × bloques`, opcionalmente ponderado por FactorExcusable y FactorOperacional |
| `BloquesMuestreo` | C12 | Cantidad de bloques de 15 min en el período seleccionado |
| `TotalRacks` | C11 | `total_pcs × baterias_por_pcs × racks_por_pcs` = 61 × 4 × 12 = **2.928** |

Campo que gobierna la indisponibilidad: **`Arena - PCS XX - POWERELECTRONICS HEM-k NUMBER OF MODULES`**

Un PCS se considera indisponible en un bloque cuando `NUMBER_OF_MODULES < 4`. Si el campo está vacío, se trata como disponible (= 4) y se marca con `ModulosDisponiblesNulo = True`.

---

## Flujo semanal (lunes) — SCADA → Excel → Python → SQL → Power BI

La data cruda de `RawData-PCS` sale cada lunes del **server SCADA** (TeamViewer hoy). Solo extracción allí; macros y ETL en PC local.

```text
SCADA ~03:00 (extract only)
  → data/inbox/raw_pcs_<corte>.<ext>
  → acquire-wait + scada_adapter (fechas mm-dd → serial Excel, mapping, backup .xlsm)
  → macros COM + extrae C12/C14/C16/C19
  → run-etl → SQL (IdCorrida nuevo)
  → reconcile → notify-bi (Misael / PBI)
```

Runbook: [`docs/runbook-lunes.md`](./docs/runbook-lunes.md). Diseño: `AGENTS.md` §14 Fase S.

| Regla | Detalle |
|---|---|
| Fechas reporte | Export incremental: desde el dato siguiente al último cargado hasta el último dato del lunes; el adapter valida la continuidad (D-07) |
| Transporte | Solo TeamViewer (sin UNC/API) |
| Formato | Origen `mm-dd-aaaa` → serial/fecha Excel real en la hoja (formato visual `dd-mm-aaaa`; nunca texto) |
| PlantActivity | Fuente aparte (no desde SCADA) |
| PBI | Owner Misael; notificación hasta service principal |

---

## Arquitectura del flujo

```text
SCADA (server) --extract--> data/inbox/
  → Fase S/T: acquire-wait + scada_adapter
  → Fase M: macros Excel COM (referencia KPI)
  → ModuloStaging (copia inmutable)
  → ModuloNormalizacion (ancho → largo)
  → MotorDisponibilidad || MotorEventosFalla
  → AgregacionDiaria/Anual
  → SQL Server (KPI + intermedios + auditoría)
  → Reporting / BI (refresh PBI / notificación)
```

---

## Estructura del proyecto

```text
ETL_Arena/
├── data/
│   ├── AvailabilityCalculation_... .xlsm
│   ├── inbox/          # drop zone SCADA semanal
│   ├── processed/      # originales inmutables + sha256
│   └── work/           # copia de trabajo .xlsm
├── docs/runbook-lunes.md
├── scripts/run_lunes.py
├── src/etl_arena/
│   ├── config/
│   ├── excel_semantics/ # reglas VBA/Excel (celdas, texto, fechas, redondeo)
│   ├── model/
│   ├── acquisition/    # Fase S: acquire-wait, contrato/lector SCADA
│   ├── workbook/       # Fase M: sesión COM, preparar, macros, referencia
│   ├── ingestion/
│   ├── normalization/
│   ├── enrichment/
│   ├── availability/
│   ├── fault_events/
│   ├── aggregation/
│   ├── persistence/
│   ├── reconciliation/
│   └── reporting/
├── tests/
├── sql/
├── AGENTS.md
├── CLAUDE.md
└── README.md
```

**Estado:** `src/etl_arena/` (solo `__init__.py`), `sql/` y `scripts/` son estructura objetivo; el motor aún no está implementado (hay goldens, `pyproject.toml` y specs). Layout detallado en `design.md §Estructura del repositorio`.

---

## Desarrollo

```powershell
python -m venv .venv                       # Python >= 3.13 (probado con 3.14)
.venv\Scripts\python -m pip install -e ".[dev]"
.venv\Scripts\python -m pytest            # tests (markers: golden, sql, excel)
.venv\Scripts\python -m ruff check src tests
```

- SQL Server local para tests de integración: copiar `.env.example` a `.env` y `docker compose -f docker/mssql.compose.yml --env-file .env up -d`.
- Requiere **ODBC Driver 18 for SQL Server** (instalador de Microsoft, con permisos de administrador). El driver legacy `SQL Server` no sirve (no maneja `DATETIME2` ni `fast_executemany`).
- Golden references: `python tests/golden/extract_golden.py --excel <libro> --month N --year 2026 --label <mes>` (ver `tests/golden/data/golden_index.json`).
- Agentes de Claude Code en `.claude/agents/` (plan de asignación en `.kiro/specs/etl-arena-availability/tasks.md`).

## Tablas SQL Server

| Tabla | Contenido |
|---|---|
| `etl_run` | Metadatos de cada corrida: `IdCorrida`, período, parámetros, `VersionAlgoritmo` |
| `proyecto` | Tabla maestra de 4 proyectos BESS con parámetros de configuración |
| `tipo_detencion` | Catálogo de 167 códigos de falla (F0…F257; seed generado desde la hoja `PCS-Fault`) |
| `raw_pcs_sample` | Datos crudos normalizados — 1 fila por `NumeroPCS × MarcaTiempoMuestra` |
| `plant_activity_sample` | `FactorOperacional` (`EsOperacional`) y `FactorExcusable` (`EsEventoExcusable`) por timestamp |
| `availability_sample_result` | Resultado intermedio por `NumeroPCS × MarcaTiempoMuestra`: `BateriasIndisponibles`, factores, `ImpactoRackPonderado` |
| `availability_run_result` | KPI del período: `DisponibilidadPeriodo` (C16), `DisponibilidadAnualAcumulada` (C19), `BloquesMuestreo` (C12), `BloquesRacksIndisponibles` (C14) |
| `fault_event` | Eventos de falla consolidados con `DuracionHoras`, `CodigoFalla`, `PromedioBateriasInvolucradas`, `HorasRackIndisponibles` |
| `detencion` | Vista operacional de detenciones con `IdProyecto`, `DuracionSegundos`, `EstadoRevision` y `Observacion` |
| `daily_availability` | KPI diario (`Disponibilidad`) y `Variacion` respecto al día anterior |
| `annual_availability` | `DisponibilidadMensual` y `DisponibilidadAcumulada` anual |

SQL es **append-only por `IdCorrida`**. No se borran resultados históricos.

---

## Plan de implementación

| Etapa | Descripción | Estado |
|---|---|---|
| 1 | Reverse engineering de macros VBA y fórmulas Excel | Completado |
| **S** | **Adquisición SCADA (inbox + adapter de fechas) y orquestador del lunes** | **Diseñado — ver AGENTS §14 Fase S / tasks 15** |
| 2 | Data contract: columnas, tipos, timestamps, tratamiento DST | Pendiente (ampliado con data contract SCADA tras muestra P0) |
| 3 | DDL SQL Server: staging, normalized, KPI, auditoría | Pendiente |
| 4 | ETL: extract -> validate -> stage -> normalize -> enrich | Pendiente |
| 5 | MotorDisponibilidad (equivalente `cmdCalcAvailability`) | Pendiente |
| 6 | MotorEventosFalla (equivalente `mcoCreateList`) | Pendiente |
| 7 | Agregaciones: KPI diario, mensual, anual | Pendiente |
| 8 | ModuloReconciliacion automático Excel vs Python | Pendiente (integridad de goldens: `tests/golden/test_golden_integrity.py` listo) |
| 8b | Runner COM de macros (sustituye trabajo manual de Alex) | Pendiente (Fase M) |
| 9 | Shadow mode: Excel oficial, Python en paralelo | Pendiente |
| 10 | Producción: origen → orquestador lunes → Python/SQL → refresh PBI | Pendiente |

El detalle de cada etapa, las fórmulas exactas y las reglas de negocio están en [`AGENTS.md`](./AGENTS.md).

---

## Criterios de aceptación

La migración se considera completa cuando una corrida Python pueda reproducir exactamente una corrida Excel con el mismo input, verificando:

- `BloquesMuestreo` (C12) — bloques de 15 min procesados
- `BloquesRacksIndisponibles` (C14) — acumulado de racks indisponibles × bloques
- `DisponibilidadPeriodo` (C16) — disponibilidad del período
- `DisponibilidadAnualAcumulada` (C19) — disponibilidad acumulada anual
- Lista completa de eventos en `ListOfFaults`
- `Disponibilidad` diaria y acumulada mensual
- Períodos que atraviesan cambio de horario (DST Chile, septiembre 2026)

Tolerancias numéricas: tabla única en `.kiro/specs/etl-arena-availability/design.md §Tolerancias` (p. ej. C14 ≤ 1e-6, C16/C19 ≤ 1e-9).

---

## Riesgos principales

| Riesgo | Descripción | Acción |
|---|---|---|
| A | Comentario VBA dice `CURRENT FAULT` pero la lógica usa `NUMBER_OF_MODULES` | Priorizar código ejecutable sobre comentarios |
| B | Fórmula anual hardcodea `365 × 24 × 4` días/bloques | Conservar durante paridad; revisar con negocio después |
| C | Factor `12` está fijo en el motor | Parametrizar luego de obtener paridad |
| D | Fechas seriales Excel pueden diferir por DST | Conservar `SerialFechaExcelOrigen` + `MarcaTiempoLocalOrigen` en staging |
| E | `RawData-PCS` tiene 244 columnas en formato ancho | NormalizadorPCS convierte a modelo largo antes del cálculo |
| F | `mcoCleanTable` destruye resultados anteriores | SQL append-only por `IdCorrida`; nunca replicar ese comportamiento |
| G | `NUMBER_OF_MODULES` vacío ignorado silenciosamente por el Excel | Marcar con `ModulosDisponiblesNulo = True`; incluir en reporte de calidad |
| H | Descripción de falla puede provenir del intervalo anterior | Marcar con `DescripcionFallaFallback = True`; reportar frecuencia |
| I | Export SCADA con formato de fecha distinto al Excel | Fase T: transformador estricto `mm-dd-aaaa` → serial Excel (formato `dd-mm-aaaa`, nunca texto) + validación de rango |
| J | Sin API SCADA; transporte solo TeamViewer | `acquire-wait` + alerta si no hay archivo; pedir share/API a GPM como mejora |
| K | PlantActivity no se actualiza con SCADA | Fuente fuera de esta cadena; se une por fila: validar alineación y reportar filas sin timestamp en calidad de corrida |

---

## Reglas clave para desarrolladores

1. **No modificar el algoritmo** durante la fase de paridad — reproducirlo exactamente.
2. **No hardcodear** `61`, `4`, `12` o `15` — provienen de `ConfiguracionCalculo`.
3. **`NUMBER_OF_MODULES`** es el campo que gobierna la indisponibilidad, no el código de falla.
4. **SQL append-only** — usar `IdCorrida`; nunca sobreescribir resultados históricos.
5. **Timestamps**: conservar `SerialFechaExcelOrigen` + `MarcaTiempoLocalOrigen`; no convertir a UTC sin documentar.
6. **`ModulosDisponiblesNulo`**: intervalos con `NUMBER_OF_MODULES` vacío = disponible (= 4) + flag de auditoría.
7. **`DescripcionFallaFallback`**: eventos con descripción tomada del intervalo anterior se marcan en `fault_event`.
8. **Versionado**: cada resultado debe incluir `VersionAlgoritmo` (ej. `"availability-v1-excel-parity"`).

---

## Inspección del Excel desde CLI

Extraer código VBA:

```bash
pip install oletools
python -m oletools.olevba "data/AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm"
```

Inspeccionar estructura XML:

```bash
# listar hojas y nombres definidos
unzip -p "data/AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm" xl/workbook.xml

# buscar en hojas grandes sin cargarlas completas (~5 MB cada una)
unzip -p "data/..." xl/worksheets/sheet10.xml | grep "termino"
```

---

## Documentación de referencia

| Archivo | Contenido |
|---|---|
| [`AGENTS.md`](./AGENTS.md) | Plan completo de implementación, lógica exacta de las macros, fórmulas, diseño de datos SQL, ETL, estrategia de paridad, **Fase S SCADA** |
| [`CLAUDE.md`](./CLAUDE.md) | Contexto técnico para agentes de IA: arquitectura del Excel, módulos VBA, reglas para agentes |
| [`docs/runbook-lunes.md`](./docs/runbook-lunes.md) | Checklist operativa del lunes (SCADA → inbox → macros → ETL) |
| `.kiro/specs/etl-arena-availability/` | SDD revisión 2 (2026-09-24): `audit.md` (auditoría contra el VBA real), requirements, design, tasks por fases con agente asignado |

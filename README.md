# ETL_Arena — Reingeniería de Disponibilidad PCS & Baterías

Proyecto de reingeniería del cálculo de disponibilidad de unidades PCS (Power Conversion System) y racks de baterías, actualmente implementado en Excel/VBA, hacia un proceso reproducible y auditable en **Python + SQL Server**.

> **Estado:** Etapa 1 completada (reverse engineering). En preparación: data contract y esquema SQL.

---

## Contexto

El proyecto **Arena BESS**, operativo desde el **08/Abril/2026**, requiere calcular mensualmente la **disponibilidad** de 61 unidades PCS (Power Conversion System), sus 244 baterías BEC (61 PCS × 4 módulos) y 2.928 racks (61 × 4 × 12). Actualmente el equipo TS-ESD (Trina Solar Chile) realiza ese cálculo con un libro Excel con macros VBA. El resultado es el KPI contractual que determina si el activo cumple su compromiso de disponibilidad.

El objetivo de este proyecto es reemplazar ese proceso de forma progresiva y controlada, garantizando que el resultado Python sea idéntico al Excel antes de retirar el Excel como fuente oficial.

---

## KPI principal

```
Availability_Period = 1 - C14 / (Total_Racks × C12)
```

| Variable | Significado |
|---|---|
| `C14` | Acumulado de `(racks indisponibles) × bloques`, opcionalmente ponderado por factor excusable y operacional |
| `C12` | Cantidad de bloques de 15 min en el período seleccionado |
| `Total_Racks` | `total_pcs × batteries_per_pcs × racks_per_pcs` = 61 × 4 × 12 = **2.928** |

Campo que gobierna la indisponibilidad: **`Arena - PCS XX - POWERELECTRONICS HEM-k NUMBER OF MODULES`**

Un PCS se considera indisponible en un bloque cuando `NUMBER_OF_MODULES < 4`. Si el campo está vacío, se trata como disponible (= 4) y se marca con `modules_available_is_null = True`.

---

## Arquitectura del flujo

```
┌──────────────────┐
│  Datos origen    │  ← RawData-PCS + PlantActivity (Excel / exportación)
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│  ETL / Staging   │  ← copia inmutable, sin transformaciones
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│  Normalización   │  ← formato ancho (244 cols) → formato largo (1 fila por PCS×ts)
└────────┬─────────┘
         │
    ┌────┴────┐
    ▼         ▼
┌──────────┐ ┌──────────────┐
│Availability│ │ Fault Events │
│  Engine  │ │    Engine    │
└────┬─────┘ └──────┬───────┘
     │               │
     └───────┬───────┘
             ▼
    ┌──────────────────┐
    │    SQL Server    │  ← KPI + intermedios + auditoría
    └────────┬─────────┘
             │
             ▼
    ┌──────────────────┐
    │  Reporting / BI  │
    └──────────────────┘
```

---

## Estructura del proyecto

```
ETL_Arena/
├── data/
│   └── AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm
├── src/
│   ├── config/           # parámetros y versión del algoritmo
│   ├── ingestion/        # lectura del Excel / archivos origen
│   ├── staging/          # carga raw sin transformar
│   ├── normalization/    # ancho → largo por PCS
│   ├── enrichment/       # join con PlantActivity
│   ├── availability/     # equivalente a cmdCalcAvailability
│   ├── fault_events/     # equivalente a mcoCreateList
│   ├── aggregation/      # KPI diario, mensual, anual
│   ├── persistence/      # escritura en SQL Server
│   ├── reconciliation/   # comparación Excel vs Python
│   └── reporting/        # capa de presentación
├── tests/
├── sql/                  # DDL de tablas SQL Server
├── AGENTS.md             # plan de implementación detallado (fuente de verdad)
├── CLAUDE.md             # contexto técnico para agentes de IA
└── README.md
```

---

## Tablas SQL Server

| Tabla | Contenido |
|---|---|
| `etl_run` | Metadatos de cada corrida: `run_id`, período, parámetros, versión del algoritmo |
| `raw_pcs_sample` | Datos crudos normalizados — 1 fila por `PCS × timestamp` |
| `plant_activity_sample` | Factor operacional (`is_operational`) y factor excusable (`is_excused_event`) por timestamp |
| `availability_sample_result` | Resultado intermedio por `PCS × timestamp`: baterías indisponibles, factores, impacto ponderado |
| `availability_run_result` | KPI del período: disponibilidad y acumulados (`C12`, `C14`, `C16`, `C19`) |
| `fault_event` | Eventos de falla consolidados con duración, código, promedio de baterías e impacto |
| `daily_availability` | KPI diario y variación respecto al día anterior |
| `annual_availability` | KPI mensual y acumulado anual |

SQL es **append-only por `run_id`**. No se borran resultados históricos.

---

## Plan de implementación

| Etapa | Descripción | Estado |
|---|---|---|
| 1 | Reverse engineering de macros VBA y fórmulas Excel | ✅ Completado |
| 2 | Data contract: columnas, tipos, timestamps, tratamiento DST | 🔲 Pendiente |
| 3 | DDL SQL Server: staging, normalized, KPI, auditoría | 🔲 Pendiente |
| 4 | ETL: extract → validate → stage → normalize → enrich | 🔲 Pendiente |
| 5 | Motor de disponibilidad (equivalente `cmdCalcAvailability`) | 🔲 Pendiente |
| 6 | Motor de eventos de falla (equivalente `mcoCreateList`) | 🔲 Pendiente |
| 7 | Agregaciones: KPI diario, mensual, anual | 🔲 Pendiente |
| 8 | Reconciliación automática Excel vs Python | 🔲 Pendiente |
| 9 | Shadow mode: Excel oficial, Python en paralelo | 🔲 Pendiente |
| 10 | Producción: Python/SQL reemplaza al Excel | 🔲 Pendiente |

El detalle de cada etapa, las fórmulas exactas y las reglas de negocio están en [`AGENTS.md`](./AGENTS.md).

---

## Criterios de aceptación

La migración se considera completa cuando una corrida Python pueda reproducir exactamente una corrida Excel con el mismo input, verificando:

- `C12` (bloques de 15 min procesados)
- `C14` (acumulado de racks indisponibles × bloques)
- `C16` (disponibilidad del período)
- `C19` (disponibilidad acumulada anual)
- Lista completa de eventos en `ListOfFaults`
- Disponibilidad diaria y acumulada mensual
- Períodos que atraviesan cambio de horario (DST Chile, septiembre 2026)

Tolerancia numérica inicial: `|Δ| ≤ 1e-9` para cálculos intermedios.

---

## Riesgos principales

| Riesgo | Descripción | Acción |
|---|---|---|
| A | Comentario VBA dice `CURRENT FAULT` pero la lógica usa `NUMBER_OF_MODULES` | Priorizar código ejecutable sobre comentarios |
| B | Fórmula anual hardcodea `365 × 24 × 4` días/bloques | Conservar durante paridad; revisar con negocio después |
| C | Factor `12` está fijo en el motor | Parametrizar luego de obtener paridad |
| D | Fechas seriales Excel pueden diferir por DST | Conservar serial original + timestamp local en staging |
| E | `RawData-PCS` tiene 244 columnas en formato ancho | Normalizar a modelo largo antes del cálculo |
| F | `mcoCleanTable` destruye resultados anteriores | SQL append-only; nunca replicar ese comportamiento |
| G | `NUMBER_OF_MODULES` vacío ignorado silenciosamente por el Excel | Marcar con `modules_available_is_null = True`; incluir en reporte de calidad de datos |
| H | Descripción de falla puede provenir del intervalo anterior (`"NO FAULTS"` con anterior distinto) | Marcar con `fault_description_fallback = True`; reportar frecuencia |

---

## Reglas clave para desarrolladores

1. **No modificar el algoritmo** durante la fase de paridad — reproducirlo exactamente.
2. **No hardcodear** `61`, `4`, `12` o `15` — provienen de la configuración del cálculo.
3. **`NUMBER_OF_MODULES`** es el campo que gobierna la indisponibilidad, no el código de falla.
4. **SQL append-only** — usar `run_id`; nunca sobreescribir resultados históricos.
5. **Timestamps**: conservar serial Excel + timestamp local; no convertir a UTC sin documentar.
6. **`modules_available_is_null`**: intervalos con `NUMBER_OF_MODULES` vacío = disponible (= 4) + flag de auditoría.
7. **`fault_description_fallback`**: eventos con descripción tomada del intervalo anterior se marcan en `fault_event`.
8. **Versionado**: cada resultado debe incluir `algorithm_version` (ej. `availability-v1-excel-parity`).

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
unzip -p "data/..." xl/worksheets/sheet10.xml | grep "término"
```

---

## Documentación de referencia

| Archivo | Contenido |
|---|---|
| [`AGENTS.md`](./AGENTS.md) | Plan completo de implementación, lógica exacta de las macros, fórmulas, diseño de datos SQL, ETL, estrategia de paridad |
| [`CLAUDE.md`](./CLAUDE.md) | Contexto técnico para agentes de IA: arquitectura del Excel, módulos VBA, reglas para agentes |

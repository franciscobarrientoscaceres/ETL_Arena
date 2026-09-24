# Design Document — ETL Arena Availability

## Overview

El sistema ETL Arena Availability replica en Python la lógica de las macros VBA del libro Excel `AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm` para el activo **Arena BESS**. El objetivo de la primera versión es obtener paridad exacta con el Excel, permitiendo comparar corrida a corrida antes de retirarlo como fuente oficial del KPI contractual.

El sistema lee los datos crudos de `RawData-PCS` y `PlantActivity`, los normaliza, aplica la lógica de disponibilidad y eventos de falla, persiste todos los resultados intermedios y finales en SQL Server identificados por un `run_id`, y produce un reporte de reconciliación que compara los resultados Python contra los del Excel.

**Principio fundamental:** SQL Server almacena no solo el KPI final sino todos los valores intermedios necesarios para explicar exactamente cómo se obtuvo ese porcentaje.

---

## Architecture

```
Datos origen (Excel: RawData-PCS 244 cols + PlantActivity)
        |
        v
src/ingestion/       -- lectura openpyxl/pandas, seriales de fecha, anomalías
        |
        v
src/staging/ + SQL   -- copia inmutable (raw_pcs_sample, plant_activity_sample)
        |
        v
src/normalization/   -- ancho 244 cols -> largo 1 fila/PCS*timestamp
        |
        v
src/enrichment/      -- join con PlantActivity (operational_factor, excused_factor)
        |
      /   \
     v     v
src/availability/    src/fault_events/
cmdCalcAvailability  mcoCreateList
C12, C14, C16, C19   fault_event
     |                    |
     +--------+-----------+
              |
              v
src/aggregation/     -- mcoDailyAvailability + Annual_AVA
              |
              v
src/persistence/     -- SQLAlchemy -> SQL Server (append-only por run_id)
              |
              v
src/reconciliation/  -- comparación Excel vs Python en 5 niveles
```

---

## Components and Interfaces

### src/config/

**CalculationConfig** — dataclass con todos los parámetros de una corrida.

```python
@dataclass
class CalculationConfig:
    run_id: str                   # UUID generado al inicio de la corrida
    algorithm_version: str        # "availability-v1-excel-parity"
    project_name: str             # "Arena BESS"
    project_start_date: date      # date(2026, 4, 8)
    period_start: datetime
    period_end: datetime
    total_pcs: int                # 61
    batteries_per_pcs: int        # 4
    racks_per_pcs: int            # 12
    total_racks: int              # calculado: total_pcs * batteries_per_pcs * racks_per_pcs
    sampling_minutes: int         # 15
    only_operational_time: bool   # equivalente a celda C21
    apply_excused_event: bool     # equivalente a celda C31
    source_file: str
    timezone: str                 # "America/Santiago"
```

`total_racks` siempre se calcula como `total_pcs * batteries_per_pcs * racks_per_pcs`. Nunca se hardcodea.

Valores por defecto para la fase de paridad:
```python
DEFAULT_CONFIG = dict(
    algorithm_version="availability-v1-excel-parity",
    project_name="Arena BESS",
    project_start_date=date(2026, 4, 8),
    total_pcs=61,
    batteries_per_pcs=4,
    racks_per_pcs=12,
    sampling_minutes=15,
    timezone="America/Santiago",
)
```

---

### src/ingestion/

**ExcelIngestionService**

```python
class ExcelIngestionService:
    def read_raw_pcs(self, path: str) -> pd.DataFrame:
        """Lee RawData-PCS desde fila 2. Conserva todas las columnas.
        Añade columna 'source_excel_serial_datetime' con el serial numérico
        original de la celda de fecha."""

    def read_plant_activity(self, path: str) -> pd.DataFrame:
        """Lee PlantActivity desde fila 2. Retorna el DataFrame completo."""

    def detect_timestamp_anomalies(
        self, df: pd.DataFrame, sampling_minutes: int
    ) -> list[AnomalyReport]:
        """Detecta y reporta: huecos, duplicados, fuera de orden, frecuencia irregular."""
```

**AnomalyReport** — dataclass(anomaly_type, row_number, timestamp, detail).

Decisiones de diseño:
- `openpyxl` para acceder a los seriales numéricos de fecha antes de la conversión.
- `pandas` para manipulación posterior.
- Los seriales de fecha Excel se conservan como `float64` en `source_excel_serial_datetime`.
- La conversión a timestamp local usa `pytz` con zona `America/Santiago`, preservando la semántica DST de Chile.

---

### src/staging/

**StagingRepository**

```python
class StagingRepository:
    def save_raw_pcs_batch(self, run_id: str, df: pd.DataFrame) -> None:
        """Inserta en bulk en raw_pcs_sample sin ninguna transformación."""

    def save_plant_activity_batch(self, run_id: str, df: pd.DataFrame) -> None:
        """Inserta en bulk en plant_activity_sample sin ninguna transformación."""
```

---

### src/normalization/

**PCSNormalizer**

```python
class PCSNormalizer:
    COLUMN_PATTERN = re.compile(
        r"Arena - PCS (\d{2}) - POWERELECTRONICS (.*)"
    )
    FIELD_SUFFIXES = {
        "GEN3 HEx CURRENT FAULT": "fault_description_raw",
        "GEN3 HEx STATUS": "status_raw",
        "GEN3 HEx WARNING": "warning_raw",
        "HEM-k NUMBER OF MODULES": "modules_available",
    }

    def normalize(
        self, raw_df: pd.DataFrame, config: CalculationConfig
    ) -> pd.DataFrame:
        """Transforma formato ancho a largo. Aplica regla modules_available_is_null."""

    def detect_schema_anomalies(
        self, raw_df: pd.DataFrame, config: CalculationConfig
    ) -> list[AnomalyReport]:
        """Detecta PCS faltantes, columnas faltantes/desplazadas, nombres inesperados."""
```

Regla de normalización para módulos disponibles:
```python
if pd.isna(raw_value) or raw_value == "":
    modules_available = 4
    modules_available_is_null = True
else:
    modules_available = int(raw_value)
    modules_available_is_null = False
```

---

### src/enrichment/

**EnrichmentService**

```python
class EnrichmentService:
    def enrich(
        self,
        normalized_df: pd.DataFrame,
        plant_activity_df: pd.DataFrame,
    ) -> pd.DataFrame:
        """
        Join por sample_timestamp.
        Columna C de PlantActivity -> operational_factor (0/1).
        Columna D de PlantActivity -> excused_factor (0/1).
        Huecos en el join: operational_factor=1, excused_factor=1, con advertencia.
        """
```

---

### src/availability/

**AvailabilityEngine** — equivalente exacto de `cmdCalcAvailability`.

```python
@dataclass
class AvailabilityResult:
    c12: int                          # sample_blocks
    c14: float                        # unavailable_rack_blocks
    availability_period: float | None # 1 - C14 / (total_racks * C12)
    availability_annual: float | None # 1 - C14 / (total_racks * 365 * 24 * 4)
    sample_results: pd.DataFrame      # filas para availability_sample_result

class AvailabilityEngine:
    def calculate(
        self, enriched_df: pd.DataFrame, config: CalculationConfig
    ) -> AvailabilityResult: ...
```

**Lógica exacta del motor (transcripción del VBA a Python):**

```python
def calculate(self, enriched_df, config):
    c12 = 0
    c14 = 0.0
    rows = []

    # Ordenar por timestamp ASC, luego por pcs_number
    df = enriched_df.sort_values(["sample_timestamp", "pcs_number"])
    timestamps = df["sample_timestamp"].unique()
    timestamps.sort()

    for ts in timestamps:
        c12 += 1  # un incremento por intervalo temporal, no por PCS
        ts_rows = df[df["sample_timestamp"] == ts]

        for _, row in ts_rows.iterrows():
            is_null = row["modules_available_is_null"]
            modules = row["modules_available"]

            # Condición VBA: <> "" AND < 4
            # Python: not is_null AND modules < batteries_per_pcs
            if not is_null and modules < config.batteries_per_pcs:
                batteries_unavail = config.batteries_per_pcs - modules

                if config.apply_excused_event:
                    weighted = batteries_unavail * row["excused_factor"]
                else:
                    weighted = float(batteries_unavail)

                if config.only_operational_time:
                    rack_impact = config.racks_per_pcs * weighted * row["operational_factor"]
                else:
                    rack_impact = config.racks_per_pcs * weighted

                c14 += rack_impact
            else:
                batteries_unavail = 0
                weighted = 0.0
                rack_impact = 0.0

            rows.append({
                "run_id": config.run_id,
                "sample_timestamp": ts,
                "pcs_number": row["pcs_number"],
                "modules_available": modules,
                "modules_available_is_null": is_null,
                "batteries_unavailable": batteries_unavail,
                "excused_factor": row["excused_factor"],
                "operational_factor": row["operational_factor"],
                "weighted_batteries_unavailable": weighted,
                "weighted_rack_impact": rack_impact,
            })

    availability_period = (1 - c14 / (config.total_racks * c12)) if c12 > 0 else None
    # Fórmula anual: supuesto 365*24*4 del Excel (conservado en fase de paridad)
    denom_annual = config.total_racks * 365 * 24 * 4
    availability_annual = (1 - c14 / denom_annual) if denom_annual > 0 else None

    return AvailabilityResult(
        c12=c12, c14=c14,
        availability_period=availability_period,
        availability_annual=availability_annual,
        sample_results=pd.DataFrame(rows),
    )
```

---

### src/fault_events/

**FaultEventsEngine** — equivalente exacto de `mcoCreateList`.

```python
@dataclass
class FaultEvent:
    run_id: str
    pcs_number: int
    start_timestamp: datetime
    end_timestamp: datetime
    duration_hours: float
    fault_code: str
    fault_description: str
    fault_description_fallback: bool
    average_batteries_involved: float
    unavailable_rack_hours: float

class FaultEventsEngine:
    def detect_events(
        self, enriched_df: pd.DataFrame, config: CalculationConfig
    ) -> list[FaultEvent]: ...
```

**Lógica exacta (transcripción del VBA):**

```python
def detect_events(self, enriched_df, config):
    events = []
    sampling_fraction = config.sampling_minutes / (24 * 60)  # fracción de día

    for pcs in range(1, config.total_pcs + 1):
        pcs_df = enriched_df[enriched_df["pcs_number"] == pcs].sort_values("sample_timestamp")
        event_open = False
        prev_description = ""
        start_ts = None
        sumablocks = 0.0
        numBlock = 0
        event_description = ""
        fallback = False

        for _, row in pcs_df.iterrows():
            ts = row["sample_timestamp"]
            is_null = row["modules_available_is_null"]
            modules = row["modules_available"]
            description = row.get("fault_description_raw", "")

            is_fault = (not is_null) and (modules < config.batteries_per_pcs)

            if is_fault:
                if not event_open:
                    event_open = True
                    # start_timestamp: restar la frecuencia de muestreo (semántica VBA exacta)
                    start_ts = ts - timedelta(minutes=config.sampling_minutes)

                    # Lógica de descripción con fallback
                    if description == "NO FAULTS":
                        if prev_description not in ("NO FAULTS", ""):
                            event_description = prev_description
                            fallback = True
                        else:
                            event_description = "F13 NO MODULES"
                            fallback = False
                    elif description == "":
                        event_description = "F1 Watchdog"
                        fallback = False
                    else:
                        event_description = description
                        fallback = False

                    sumablocks = 0.0
                    numBlock = 0

                batt_diff = config.batteries_per_pcs - modules
                if config.apply_excused_event:
                    sumablocks += batt_diff * row["excused_factor"]
                else:
                    sumablocks += float(batt_diff)
                numBlock += 1

            else:
                if event_open:
                    end_ts = ts
                    duration_h = (end_ts - start_ts).total_seconds() / 3600
                    avg_bat = sumablocks / numBlock if numBlock > 0 else 0.0
                    rack_hours = config.racks_per_pcs * duration_h * avg_bat
                    fault_code = _extract_fault_code(event_description)

                    events.append(FaultEvent(
                        run_id=config.run_id,
                        pcs_number=pcs,
                        start_timestamp=start_ts,
                        end_timestamp=end_ts,
                        duration_hours=duration_h,
                        fault_code=fault_code,
                        fault_description=event_description,
                        fault_description_fallback=fallback,
                        average_batteries_involved=avg_bat,
                        unavailable_rack_hours=rack_hours,
                    ))
                    event_open = False

            prev_description = description

        # Cerrar evento abierto al final del período
        if event_open:
            end_ts = config.period_end + timedelta(days=1)
            duration_h = (end_ts - start_ts).total_seconds() / 3600
            avg_bat = sumablocks / numBlock if numBlock > 0 else 0.0
            rack_hours = config.racks_per_pcs * duration_h * avg_bat
            fault_code = _extract_fault_code(event_description)
            events.append(FaultEvent(
                run_id=config.run_id, pcs_number=pcs,
                start_timestamp=start_ts, end_timestamp=end_ts,
                duration_hours=duration_h, fault_code=fault_code,
                fault_description=event_description,
                fault_description_fallback=fallback,
                average_batteries_involved=avg_bat,
                unavailable_rack_hours=rack_hours,
            ))

    return events


def _extract_fault_code(description: str) -> str:
    idx = description.find(" ")
    if idx > 0:
        return description[:idx]
    return "F" + description
```

---

### src/aggregation/

**DailyAggregation**

```python
class DailyAggregation:
    def calculate(
        self, sample_results: pd.DataFrame, config: CalculationConfig
    ) -> pd.DataFrame:
        """
        Agrupa weighted_rack_impact por día.
        Calcula accumulated_unavailable_rack_blocks (suma acumulada).
        Fórmula diaria:
          availability = 1 - accum / (total_racks * 24 * 60 * day_N / sampling_minutes)
        variation = availability_N - availability_(N-1)
        """
```

**AnnualAggregation**

```python
class AnnualAggregation:
    def calculate(
        self, result: AvailabilityResult, config: CalculationConfig, month: int, year: int
    ) -> dict:
        """
        Fórmula anual (conservada del Excel para paridad):
          availability_annual = 1 - C14 / (total_racks * 365 * 24 * 4)
        Acumulación desde project_start_date (2026-04-08).
        """
```

---

### src/persistence/

**PersistenceService** — usa SQLAlchemy Core con pyodbc.

```python
class PersistenceService:
    def __init__(self, connection_string: str): ...
    def begin_run(self, config: CalculationConfig) -> None: ...
    def complete_run(self, run_id: str) -> None: ...
    def fail_run(self, run_id: str, error: str) -> None: ...
    def save_raw_pcs(self, run_id: str, df: pd.DataFrame) -> None: ...
    def save_plant_activity(self, run_id: str, df: pd.DataFrame) -> None: ...
    def save_sample_results(self, run_id: str, df: pd.DataFrame) -> None: ...
    def save_run_result(self, run_id: str, result: AvailabilityResult) -> None: ...
    def save_fault_events(self, run_id: str, events: list[FaultEvent]) -> None: ...
    def save_daily(self, run_id: str, df: pd.DataFrame) -> None: ...
    def save_annual(self, run_id: str, rows: list[dict]) -> None: ...
```

Reglas:
- Inserciones en bulk (`executemany` o `pd.to_sql` con `method="multi"`).
- Nunca `DELETE` ni `UPDATE` sobre datos de corridas anteriores.
- El `run_id` se genera en `begin_run` y se propaga a todas las llamadas `save_*`.

---

### src/reconciliation/

**ReconciliationService**

```python
@dataclass
class LevelResult:
    level: int
    name: str
    passed: bool
    discrepancies: list[dict]

@dataclass
class ReconciliationReport:
    python_run_id: str
    excel_source: str
    tolerance: float
    levels: list[LevelResult]
    overall_passed: bool

class ReconciliationService:
    def reconcile(
        self,
        python_run_id: str,
        excel_export_path: str,
        tolerance: float = 1e-9,
    ) -> ReconciliationReport: ...
```

**Niveles de comparación:**

| Nivel | Nombre | Qué compara |
|---|---|---|
| 1 | Input | filas totales, timestamps, PCS count, frecuencia, rango de fechas |
| 2 | Muestra individual | modules_available, batteries_unavailable, factores, weighted por cada (ts, pcs) |
| 3 | Acumulados | C12 y C14 |
| 4 | KPI | C16, C19, disponibilidad diaria y mensual |
| 5 | Eventos | PCS, start, end, duration, fault_code, avg_batteries, rack_hours por evento |

**Nota sobre el límite de 167 eventos:** el Excel solo ordena hasta 167 eventos en `mcoOrder`. Si el período supera ese umbral, el nivel 5 reporta advertencia y excluye la comparación de posición (ranking) para los eventos adicionales.

---

## Data Models

### SQL Server — DDL completo

```sql
-- Tabla de auditoría de corridas
CREATE TABLE etl_run (
    run_id                NVARCHAR(36)    NOT NULL PRIMARY KEY,
    source_file           NVARCHAR(500)   NOT NULL,
    source_system         NVARCHAR(50)    NOT NULL DEFAULT 'excel',
    started_at            DATETIME2(7)    NOT NULL,
    finished_at           DATETIME2(7)    NULL,
    period_start          DATETIME2(7)    NOT NULL,
    period_end            DATETIME2(7)    NOT NULL,
    sampling_minutes      INT             NOT NULL,
    total_pcs             INT             NOT NULL,
    batteries_per_pcs     INT             NOT NULL,
    racks_per_pcs         INT             NOT NULL,
    total_racks           INT             NOT NULL,
    only_operational_time BIT             NOT NULL,
    apply_excused_event   BIT             NOT NULL,
    status                NVARCHAR(20)    NOT NULL,
    error_message         NVARCHAR(MAX)   NULL,
    algorithm_version     NVARCHAR(100)   NOT NULL
);

-- Datos crudos normalizados (1 fila por PCS x timestamp)
CREATE TABLE raw_pcs_sample (
    id                           BIGINT          IDENTITY(1,1) PRIMARY KEY,
    run_id                       NVARCHAR(36)    NOT NULL REFERENCES etl_run(run_id),
    sample_timestamp             DATETIME2(7)    NOT NULL,
    source_excel_serial_datetime FLOAT           NULL,
    source_timestamp_local       DATETIME2(7)    NULL,
    pcs_number                   INT             NOT NULL,
    fault_code_raw               NVARCHAR(50)    NULL,
    fault_description_raw        NVARCHAR(255)   NULL,
    status_raw                   NVARCHAR(100)   NULL,
    warning_raw                  NVARCHAR(100)   NULL,
    modules_available            INT             NOT NULL,  -- 4 si era NULL en origen
    modules_available_is_null    BIT             NOT NULL,
    source_row_number            INT             NULL,
    source_columns               NVARCHAR(MAX)   NULL       -- JSON con nombres de columnas
);
CREATE INDEX IX_raw_pcs_sample_run_ts_pcs
    ON raw_pcs_sample(run_id, sample_timestamp, pcs_number);

-- Factores de actividad de planta
CREATE TABLE plant_activity_sample (
    id                              BIGINT          IDENTITY(1,1) PRIMARY KEY,
    run_id                          NVARCHAR(36)    NOT NULL REFERENCES etl_run(run_id),
    sample_timestamp                DATETIME2(7)    NOT NULL,
    is_operational                  BIT             NOT NULL,
    is_excused_event                BIT             NOT NULL,
    active_power_setpoint_kw        FLOAT           NULL,
    overfrequency_droop_enabled     BIT             NULL,
    underfrequency_droop_enabled    BIT             NULL,
    poi_active_power_kw             FLOAT           NULL,
    soc_percent                     FLOAT           NULL,
    source_row_number               INT             NULL
);
CREATE INDEX IX_plant_activity_run_ts
    ON plant_activity_sample(run_id, sample_timestamp);

-- Resultado intermedio por PCS x timestamp (auditoría completa)
CREATE TABLE availability_sample_result (
    id                              BIGINT          IDENTITY(1,1) PRIMARY KEY,
    run_id                          NVARCHAR(36)    NOT NULL REFERENCES etl_run(run_id),
    sample_timestamp                DATETIME2(7)    NOT NULL,
    pcs_number                      INT             NOT NULL,
    modules_available               INT             NOT NULL,
    modules_available_is_null       BIT             NOT NULL,
    batteries_unavailable           INT             NOT NULL,
    excused_factor                  FLOAT           NOT NULL,
    operational_factor              FLOAT           NOT NULL,
    weighted_batteries_unavailable  FLOAT           NOT NULL,
    weighted_rack_impact            FLOAT           NOT NULL
);
CREATE INDEX IX_avail_sample_run_ts_pcs
    ON availability_sample_result(run_id, sample_timestamp, pcs_number);

-- KPI del período (una fila por corrida)
CREATE TABLE availability_run_result (
    run_id                  NVARCHAR(36)    NOT NULL PRIMARY KEY REFERENCES etl_run(run_id),
    sample_blocks           INT             NOT NULL,        -- C12
    total_racks             INT             NOT NULL,
    unavailable_rack_blocks FLOAT           NOT NULL,        -- C14
    availability_period     FLOAT           NULL,            -- C16
    availability_annual     FLOAT           NULL             -- C19
);

-- Eventos de falla consolidados
CREATE TABLE fault_event (
    id                          BIGINT          IDENTITY(1,1) PRIMARY KEY,
    run_id                      NVARCHAR(36)    NOT NULL REFERENCES etl_run(run_id),
    pcs_number                  INT             NOT NULL,
    start_timestamp             DATETIME2(7)    NOT NULL,
    end_timestamp               DATETIME2(7)    NOT NULL,
    duration_hours              FLOAT           NOT NULL,
    fault_code                  NVARCHAR(50)    NULL,
    fault_description           NVARCHAR(255)   NULL,
    fault_description_fallback  BIT             NOT NULL DEFAULT 0,
    average_batteries_involved  FLOAT           NOT NULL,
    unavailable_rack_hours      FLOAT           NOT NULL
);
CREATE INDEX IX_fault_event_run_pcs
    ON fault_event(run_id, pcs_number);

-- KPI diario
CREATE TABLE daily_availability (
    id                                  BIGINT          IDENTITY(1,1) PRIMARY KEY,
    run_id                              NVARCHAR(36)    NOT NULL REFERENCES etl_run(run_id),
    day                                 DATE            NOT NULL,
    daily_unavailable_rack_blocks       FLOAT           NOT NULL,
    accumulated_unavailable_rack_blocks FLOAT           NOT NULL,
    availability                        FLOAT           NULL,
    variation                           FLOAT           NULL
);
CREATE INDEX IX_daily_avail_run_day
    ON daily_availability(run_id, day);

-- KPI mensual y acumulado anual
CREATE TABLE annual_availability (
    id                              BIGINT          IDENTITY(1,1) PRIMARY KEY,
    run_id                          NVARCHAR(36)    NOT NULL REFERENCES etl_run(run_id),
    year                            INT             NOT NULL,
    month                           INT             NOT NULL,
    days_in_month                   INT             NOT NULL,
    sampling_blocks                 INT             NOT NULL,
    unavailable_rack_blocks         FLOAT           NOT NULL,
    monthly_availability            FLOAT           NULL,
    accumulated_sampling_blocks     INT             NOT NULL,
    accumulated_unavailable_blocks  FLOAT           NOT NULL,
    accumulated_availability        FLOAT           NULL,
    contractual_availability        FLOAT           NULL
);
CREATE INDEX IX_annual_avail_run_ym
    ON annual_availability(run_id, year, month);
```

---

## Correctness Properties

*Una propiedad es una característica o comportamiento que debe mantenerse verdadero en todas las ejecuciones válidas del sistema — esencialmente, una declaración formal sobre lo que el sistema debe hacer. Las propiedades sirven como puente entre las especificaciones en lenguaje natural y las garantías de corrección verificables automáticamente.*

### Property 1: Normalización produce forma correcta y marca nulos consistentemente

*Para cualquier* DataFrame de entrada en formato ancho con N filas de timestamps y P PCS válidos, la salida normalizada debe tener exactamente N × P filas, una por cada combinación `(sample_timestamp, pcs_number)`. Además, para toda fila donde `NUMBER_OF_MODULES` estaba vacío en el origen, `modules_available` debe ser 4 y `modules_available_is_null` debe ser `True`; para toda fila con valor numérico, `modules_available_is_null` debe ser `False` y `modules_available` debe igualar el valor original.

**Validates: Requirements 2.1, 2.3, 2.4**

---

### Property 2: Detección de anomalías de timestamp es exhaustiva

*Para cualquier* secuencia de timestamps con anomalías inyectadas (huecos, duplicados, fuera de orden, intervalos irregulares), el módulo de detección debe reportar exactamente las posiciones y tipos de anomalías presentes, sin omisiones ni falsos positivos.

**Validates: Requirements 1.5, 1.6**

---

### Property 3: Cálculo de rack impact ponderado es correcto para toda combinación de parámetros

*Para cualquier* combinación válida de `(modules_available, excused_factor, operational_factor, apply_excused_event, only_operational_time)`, el `weighted_rack_impact` producido por el Availability_Engine debe igualar la evaluación directa de la fórmula:

- Si `modules_available >= batteries_per_pcs` o `modules_available_is_null`: `weighted_rack_impact = 0`
- Si `apply_excused_event = True` y `only_operational_time = True`: `weighted_rack_impact = racks_per_pcs * (batteries_per_pcs - modules_available) * excused_factor * operational_factor`
- Si `apply_excused_event = True` y `only_operational_time = False`: `weighted_rack_impact = racks_per_pcs * (batteries_per_pcs - modules_available) * excused_factor`
- Si `apply_excused_event = False` y `only_operational_time = True`: `weighted_rack_impact = racks_per_pcs * (batteries_per_pcs - modules_available) * operational_factor`
- Si `apply_excused_event = False` y `only_operational_time = False`: `weighted_rack_impact = racks_per_pcs * (batteries_per_pcs - modules_available)`

**Validates: Requirements 4.3, 4.4, 4.5, 4.6, 4.7**

---

### Property 4: C12 siempre cuenta intervalos temporales, no PCS

*Para cualquier* conjunto de N intervalos de tiempo con P PCS cada uno, el valor de C12 producido por el Availability_Engine debe ser igual a N, independientemente de P.

**Validates: Requirements 4.1**

---

### Property 5: C14 es igual a la suma de todos los weighted_rack_impact

*Para cualquier* ejecución del Availability_Engine, el valor de C14 debe ser igual a la suma de todos los `weighted_rack_impact` registrados en `availability_sample_result` para ese `run_id`.

**Validates: Requirements 4.1, 4.6, 4.7, 4.11**

---

### Property 6: Los eventos de falla cubren exactamente las sub-secuencias con modules_available < 4

*Para cualquier* secuencia de intervalos de un PCS, los eventos de falla detectados deben cubrir exactamente el conjunto de índices donde `modules_available < 4` y `modules_available_is_null = False`, sin solapamientos ni huecos respecto a las sub-secuencias consecutivas con esa condición.

**Validates: Requirements 5.1, 5.3**

---

### Property 7: La lógica de fallback de descripción es determinista y completa

*Para cualquier* par `(current_description, previous_description)` en el inicio de un evento de falla, el resultado de la lógica de descripción debe ser exactamente uno de los cuatro casos del árbol de decisión, y el flag `fault_description_fallback` debe ser `True` si y solo si se usó la descripción del intervalo anterior.

**Validates: Requirements 5.4, 5.5**

---

### Property 8: La disponibilidad diaria acumulada es consistente con la acumulación de rack blocks

*Para cualquier* día N con `accumulated_unavailable_rack_blocks > 0` y denominador no nulo, la disponibilidad diaria acumulada debe satisfacer `0 <= availability <= 1` y debe ser igual a `1 - accumulated_unavailable_rack_blocks / (total_racks * 24 * 60 * N / sampling_minutes)`.

**Validates: Requirements 6.3, 6.4**

---

## Error Handling

### Errores de ingesta
- Archivo no encontrado o no legible: abortar corrida, registrar en `etl_run.error_message`, `status = "failed"`.
- Hoja no encontrada en el Excel: abortar corrida con mensaje descriptivo.
- Celda vacía en columna A de RawData-PCS: registrar advertencia, continuar procesamiento.

### Errores de normalización
- Columnas faltantes o nombres inesperados: registrar advertencia con lista de columnas afectadas; los PCS sin columnas completas se procesan con los campos disponibles y se marca el resto como null.
- PCS faltantes: registrar advertencia indicando qué números de PCS no están en el archivo.

### Errores de persistencia
- Falla de conexión a SQL Server: `fail_run`, no dejar registros huérfanos sin `etl_run` de contexto.
- Error en inserción bulk: hacer rollback de toda la corrida; no persistir resultados parciales.

### Errores numéricos
- División por cero en C16 (C12 = 0): producir `None`, no lanzar excepción.
- División por cero en fórmula diaria (denominador = 0): producir `None`.
- No redondear ningún resultado intermedio; la tolerancia de comparación en reconciliación es `1e-9`.

### DST y zonas horarias
- Períodos que cruzan cambio de horario (ej. septiembre 2026 en Chile): preservar el serial Excel original; reportar advertencia indicando cuántos registros cruzan el cambio de horario.
- No convertir a UTC sin anotar el efecto en el reporte de la corrida.

---

## Testing Strategy

### Enfoque dual: tests de ejemplo + property-based tests

La suite de pruebas usa **pytest** para tests de ejemplo y de integración, y **Hypothesis** para property-based tests. Cada test de propiedad debe ejecutar un mínimo de 100 iteraciones con inputs generados aleatoriamente.

### Tests unitarios (ejemplo-based)

- `tests/unit/test_config.py`: verificar que `total_racks` se calcula correctamente y que los valores por defecto de paridad son los correctos.
- `tests/unit/test_ingestion.py`: verificar lectura de un Excel pequeño de muestra; verificar detección de un hueco conocido.
- `tests/unit/test_normalization.py`: verificar casos específicos de la regla `modules_available_is_null`; verificar detección de columna faltante con nombre concreto.
- `tests/unit/test_availability.py`: verificar un ejemplo completo con 3 intervalos y 2 PCS conocidos, comparando C12 = 3, C14 exacto.
- `tests/unit/test_fault_events.py`: verificar los cuatro casos del árbol de decisión de descripción con ejemplos concretos; verificar el caso de evento abierto al final del período.
- `tests/unit/test_aggregation.py`: verificar fórmula diaria con un ejemplo de 5 días conocidos.
- `tests/unit/test_reconciliation.py`: verificar que el reporte indica pass con datos idénticos y fail con una discrepancia conocida.

### Property-based tests (Hypothesis)

Tag de referencia para cada test: `# Feature: etl-arena-availability, Property {N}: {título}`

- **Property 1** — `tests/property/test_normalization_shape.py`:
  Genera DataFrames anchos con N timestamps y P PCS (1 ≤ P ≤ 61, 1 ≤ N ≤ 200). Verifica que la salida tiene N×P filas y que `modules_available_is_null` es coherente con los nulos del input.

- **Property 2** — `tests/property/test_anomaly_detection.py`:
  Genera secuencias de timestamps con anomalías inyectadas en posiciones aleatorias. Verifica que cada anomalía inyectada aparece en el reporte de detección.

- **Property 3** — `tests/property/test_rack_impact.py`:
  Genera combinaciones de parámetros `(modules_available ∈ [0,4], excused_factor ∈ {0,1}, operational_factor ∈ {0,1}, apply_excused_event ∈ {True,False}, only_operational_time ∈ {True,False})`. Verifica que `weighted_rack_impact` coincide con la evaluación directa de la fórmula.

- **Property 4** — `tests/property/test_c12_count.py`:
  Genera N intervalos con P PCS. Verifica que `result.c12 == N`.

- **Property 5** — `tests/property/test_c14_sum.py`:
  Genera conjuntos de registros con valores conocidos. Verifica que `result.c14 == sum(result.sample_results["weighted_rack_impact"])`.

- **Property 6** — `tests/property/test_event_coverage.py`:
  Genera secuencias de `modules_available` con patrones aleatorios de fallas consecutivas. Verifica que los eventos detectados cubren exactamente las sub-secuencias con `modules_available < 4`.

- **Property 7** — `tests/property/test_fault_description_fallback.py`:
  Genera pares `(current_description, previous_description)` con valores de `"NO FAULTS"`, `""` y cadenas arbitrarias. Verifica que el resultado del árbol de decisión y el flag `fault_description_fallback` son correctos.

- **Property 8** — `tests/property/test_daily_availability.py`:
  Genera secuencias de `daily_unavailable_rack_blocks` con denominador no nulo. Verifica que `availability` satisface `0 <= availability <= 1` y la fórmula exacta.

### Tests de integración

- `tests/integration/test_persistence.py`: verifica que después de dos corridas con `run_id` distintos, ambos conjuntos de registros coexisten en la base de datos (comportamiento append-only).
- `tests/integration/test_full_pipeline.py`: ejecuta el pipeline completo sobre un subconjunto de datos de muestra extraídos del Excel real y verifica que los KPI producidos coinciden con los valores conocidos del Excel (tolerancia `1e-4` para la comparación de integración).

---

## Golden Tests — Valores Reales del Excel

Esta sección contiene valores extraídos directamente del libro fuente con `openpyxl` (`data_only=True`). Son la referencia definitiva para los tests de paridad. Cualquier implementación Python debe reproducir estos valores dentro de la tolerancia definida (`1e-9` para acumulados internos, `1e-6` para KPI presentados).

> Extracción realizada el 2026-09-22 sobre el archivo:
> `AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm`

---

### GT-1: Corrección de nombres de columnas en RawData-PCS

La columna de estado real se llama **`CURRENT STATUS`**, no `STATUS`. Los nombres exactos de las 4 columnas por PCS son:

```text
Arena - PCS XX - POWERELECTRONICS GEN3 HEx CURRENT FAULT
Arena - PCS XX - POWERELECTRONICS GEN3 HEx CURRENT STATUS
Arena - PCS XX - POWERELECTRONICS GEN3 HEx CURRENT WARNING
Arena - PCS XX - POWERELECTRONICS HEM-k NUMBER OF MODULES
```

El regex de normalización debe usar `CURRENT STATUS` y `CURRENT WARNING`.

---

### GT-2: Primer timestamp del dataset

El primer registro de `RawData-PCS` es `2026-04-08 00:15:00`, lo que confirma que la fecha de inicio del activo Arena BESS (`2026-04-08`) es el primer día con datos operacionales.

El dataset comienza en `00:15`, no en `00:00` — el primer bloque del día es el intervalo `[00:00, 00:15)`.

---

### GT-3: KPI del período — Corrida de Junio 2026

Parámetros usados (`Calculation-Availability` con `data_only=True`):

| Parámetro | Celda | Valor |
|---|---|---|
| Período | C5–C7 | 2026-06-01 a 2026-06-30 |
| Total PCS | C2 | 61 |
| Total Batteries per PCS | C3 | 4 |
| Total Racks | C11 | 2928 |
| Only Operational Time | C21 | `"No"` |
| Excusable Event | C31 | `"No"` |
| Frecuencia de muestreo | C23 | 15 min |

Resultados calculados por la macro VBA:

| Variable | Celda | Valor exacto del Excel |
|---|---|---|
| `C12` (sample_blocks) | C12 | `2880` |
| `C14` (unavailable_rack_blocks) | C14 | `1211931.4879999976` |
| `C16` (availability_period) | C16 | `0.8562808932908321` |
| `C19` (availability_annual) | C19 | `0.9881874706814383` |

```python
# tests/golden/test_gt3_june_kpi.py
def test_june_2026_kpi_exact():
    """GT-3: KPI de Junio 2026 debe coincidir con la corrida VBA real."""
    EXPECTED_C12 = 2880
    EXPECTED_C14 = 1211931.4879999976
    EXPECTED_C16 = 0.8562808932908321
    EXPECTED_C19 = 0.9881874706814383
    TOLERANCE    = 1e-6

    result = run_availability_engine(
        period_start="2026-06-01",
        period_end="2026-06-30",
        only_operational_time=False,
        apply_excused_event=False,
    )

    assert result.c12 == EXPECTED_C12
    assert abs(result.c14 - EXPECTED_C14) < TOLERANCE
    assert abs(result.availability_period - EXPECTED_C16) < TOLERANCE
    assert abs(result.availability_annual - EXPECTED_C19) < TOLERANCE
```

---

### GT-4: KPI diario — Primeros 10 días de Junio 2026

Valores extraídos de `Daily!C9:G18` (`data_only=True`):

| Día | Fecha | daily_unavailable | accumulated_unavailable | availability | variation |
|---|---|---|---|---|---|
| 1 | 2026-06-01 | 4245.128 | 4245.128 | 0.984897512522 | 0 |
| 2 | 2026-06-02 | 4226.760 | 8471.888 | 0.984930185564 | +3.267e-3 |
| 3 | 2026-06-03 | 3213.880 | 11685.768 | 0.986142218806 | +1.212e-3 |
| 4 | 2026-06-04 | 4382.048 | 16067.816 | 0.985709265425 | -4.330e-4 |
| 5 | 2026-06-05 | 5130.228 | 21198.044 | 0.984917147654 | -7.921e-4 |
| 6 | 2026-06-06 | 4611.300 | 25809.344 | 0.984696759259 | -2.204e-4 |
| 7 | 2026-06-07 | 9465.964 | 35275.308 | 0.982072056742 | -2.625e-3 |
| 8 | 2026-06-08 | 863.648 | 36138.956 | 0.983928984873 | +1.857e-3 |
| 9 | 2026-06-09 | 10993.732 | 47132.688 | 0.981368947328 | -2.560e-3 |
| 10 | 2026-06-10 | 6108.176 | 53240.864 | 0.981059005009 | -3.099e-4 |

**Nota:** Los días 22–30 tienen `daily_unavailable = 0` y `availability = None`, lo que indica que la macro se corrió hasta el día 21 de junio (los bloques de esos días estaban vacíos en `Calculation-Availability`).

```python
# tests/golden/test_gt4_daily.py
DAILY_GT = [
    (1,  "2026-06-01", 4245.128,   4245.128,    0.984897512522,  0.0),
    (2,  "2026-06-02", 4226.760,   8471.888,    0.984930185564,  3.267e-3),
    (3,  "2026-06-03", 3213.880,   11685.768,   0.986142218806,  1.212e-3),
    (7,  "2026-06-07", 9465.964,   35275.308,   0.982072056742, -2.625e-3),
    (10, "2026-06-10", 6108.176,   53240.864,   0.981059005009, -3.099e-4),
    (21, "2026-06-21", 912.999,    104134.292,  0.982358635695,  7.197e-4),
]

def test_daily_availability_matches_excel():
    result_df = run_daily_aggregation(period="june_2026")
    for day_n, date_str, daily_unav, accum_unav, avail, variation in DAILY_GT:
        row = result_df[result_df["day"] == date_str].iloc[0]
        assert abs(row["daily_unavailable_rack_blocks"] - daily_unav) < 0.01
        assert abs(row["accumulated_unavailable_rack_blocks"] - accum_unav) < 0.01
        assert abs(row["availability"] - avail) < 1e-6
```

---

### GT-5: Annual_AVA — Acumulados históricos

Valores extraídos de la hoja `Annual_AVA` (`data_only=True`):

| Mes | sampling_blocks | unavailable_rack_blocks | monthly_availability | accum_blocks | accum_unavailable | nota |
|---|---|---|---|---|---|---|
| Julio 2026 | 2976 | 350972.30 | 0.959721912366 | 2976 | 350972.30 | primer mes completo |
| Agosto 2026 | 2976 | 305182.97 | 0.964976762185 | 5952 | 656155.27 | acumulado dos meses |
| Sep. 2026 (parcial) | 1977 | 1211931.49 | 0.790636809650 | 7929 | 1868086.76 | corrida hasta ~20.6 días |

**Observación:** septiembre aparece con `days_in_month = 20.59375`, lo que corresponde exactamente a `1977 bloques × 15 min / (60 × 24) = 20.59375 días`. La disponibilidad mensual de septiembre (0.7907) es significativamente inferior a julio y agosto.

**Observación:** los meses abril, mayo y junio no tienen valores en `Annual_AVA`, lo que confirma que los resultados de esos meses aún no han sido cargados en el acumulado anual al momento de la extracción.

```python
# tests/golden/test_gt5_annual.py
ANNUAL_GT = [
    (7, 2976, 350972.2999999995,  0.959721912366326, 2976, 350972.2999999995),
    (8, 2976, 305182.9679999994,  0.964976762184911, 5952, 656155.2679999989),
    (9, 1977, 1211931.4879999976, 0.7906368096497706, 7929, 1868086.7559999963),
]

def test_annual_accumulation_matches_excel():
    for month, blocks, unav, avail, acc_blocks, acc_unav in ANNUAL_GT:
        row = get_annual_row(year=2026, month=month)
        assert row["sampling_blocks"] == blocks
        assert abs(row["unavailable_rack_blocks"] - unav) < 0.01
        assert abs(row["monthly_availability"] - avail) < 1e-6
        assert row["accumulated_sampling_blocks"] == acc_blocks
        assert abs(row["accumulated_unavailable_blocks"] - acc_unav) < 0.01
```

---

### GT-6: ListOfFaults — Primeros 5 eventos del PCS 1 en septiembre 2026

Valores extraídos de `ListOfFaults!B6:I25` con los parámetros de la corrida de septiembre.

| # | PCS | start | end | duration_h | fault_code | fault_description | avg_batteries | rack_hours |
|---|---|---|---|---|---|---|---|---|
| 1 | 1 | 2026-09-01 00:15:00 | 2026-09-01 05:45:00 | 5.5000 | F55 | F55 EXTERNAL FAULT/OVGR | 3.3182 | 219.000 |
| 2 | 1 | 2026-09-01 08:15:00 | 2026-09-01 08:30:00 | 0.2500 | F13 | F13 NO MODULES | 0.3333 | 1.000 |
| 3 | 1 | 2026-09-01 15:45:00 | 2026-09-01 16:15:00 | 0.5000 | F55 | F55 EXTERNAL FAULT/OVGR | 2.3343 | 14.006 |
| 4 | 1 | 2026-09-01 23:45:00 | 2026-09-02 05:15:00 | 5.5000 | F13 | F13 NO MODULES | 2.6339 | 173.838 |
| 5 | 1 | 2026-09-02 09:00:00 | 2026-09-02 09:30:00 | 0.5000 | F13 | F13 NO MODULES | 0.2468 | 1.481 |

**Total (h unavail) × (racks unavail) de todos los eventos (ListOfFaults!L10):** `25549.875880984026`

```python
# tests/golden/test_gt6_fault_events.py
FAULT_GT_PCS1 = [
    (1, "2026-09-01 00:15:00", "2026-09-01 05:45:00", 5.5,   "F55", "F55 EXTERNAL FAULT/OVGR", 3.3182, 219.000),
    (1, "2026-09-01 08:15:00", "2026-09-01 08:30:00", 0.25,  "F13", "F13 NO MODULES",           0.3333,   1.000),
    (1, "2026-09-01 15:45:00", "2026-09-01 16:15:00", 0.5,   "F55", "F55 EXTERNAL FAULT/OVGR",  2.3343,  14.006),
    (1, "2026-09-01 23:45:00", "2026-09-02 05:15:00", 5.5,   "F13", "F13 NO MODULES",           2.6339, 173.838),
    (1, "2026-09-02 09:00:00", "2026-09-02 09:30:00", 0.5,   "F13", "F13 NO MODULES",           0.2468,   1.481),
]
TOTAL_RACK_HOURS = 25549.875880984026

def test_fault_events_pcs1_matches_excel():
    events = run_fault_events_engine(period="september_2026")
    pcs1_events = sorted([e for e in events if e.pcs_number == 1],
                         key=lambda e: e.start_timestamp)
    for i, (pcs, start, end, dur, code, desc, avg_bat, rack_h) in enumerate(FAULT_GT_PCS1):
        ev = pcs1_events[i]
        assert ev.fault_code == code
        assert abs(ev.duration_hours - dur) < 0.01
        assert abs(ev.average_batteries_involved - avg_bat) < 0.01
        assert abs(ev.unavailable_rack_hours - rack_h) < 0.1

def test_total_rack_hours_matches_excel():
    events = run_fault_events_engine(period="september_2026")
    total = sum(e.unavailable_rack_hours for e in events)
    assert abs(total - TOTAL_RACK_HOURS) < 1e-3
```

---

### GT-7: Muestra puntual — Fila 2 de RawData-PCS

El primer registro del dataset (fila 2, timestamp `2026-04-08 00:15:00`):

| Campo | PCS | Valor crudo | modules_available | modules_available_is_null |
|---|---|---|---|---|
| `NUMBER_OF_MODULES` | 01 | `4` | 4 | False |
| `NUMBER_OF_MODULES` | 02 | `4` | 4 | False |
| `NUMBER_OF_MODULES` | 03 | `4` | 4 | False |
| `NUMBER_OF_MODULES` | 04 | `3.23466666666667` | **3** (o float)* | False |

*El PCS 04 tiene `3.23466666666667` en el primer bloque — valor menor que 4 pero no entero. La normalización debe preservar el valor como float antes de aplicar la condición `< 4`, o convertirlo a int según la semántica VBA. Esto requiere validación explícita.

**Con `only_operational_time=No` y `apply_excused_event=No`, `PlantActivity.C=1` y `PlantActivity.D=1`:**

| PCS | batteries_unavailable | weighted | rack_impact |
|---|---|---|---|
| 01–03 | 0 | 0 | 0 |
| 04 | `4 - 3.234... = 0.765...` | `0.765...` | `12 × 0.765... = 9.181...` |

Este caso de valor decimal en `NUMBER_OF_MODULES` debe ser documentado y testeado explícitamente.

---

### GT-8: Corrección al nombre de columna en el spec

El campo `CURRENT STATUS` (no `STATUS`) debe usarse en el patrón de regex del `PCSNormalizer`:

```python
# Corrección: en design.md §Components la tabla FIELD_SUFFIXES debe ser:
FIELD_SUFFIXES = {
    "GEN3 HEx CURRENT FAULT":  "fault_description_raw",
    "GEN3 HEx CURRENT STATUS": "status_raw",          # ← CURRENT STATUS, no STATUS
    "GEN3 HEx CURRENT WARNING": "warning_raw",         # ← CURRENT WARNING, no WARNING
    "HEM-k NUMBER OF MODULES": "modules_available",
}
```

---

# Design Document — ETL Arena Availability

## Overview

El sistema ETL Arena Availability replica en Python la lógica de las macros VBA del libro Excel `AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm` para el activo **Arena BESS**. El objetivo de la primera versión es obtener paridad exacta con el Excel, permitiendo comparar corrida a corrida antes de retirarlo como fuente oficial del KPI contractual.

El sistema lee los datos crudos de `RawData-PCS` y `PlantActivity`, los normaliza, aplica la lógica de disponibilidad y eventos de falla, persiste todos los resultados intermedios y finales en SQL Server identificados por un `IdCorrida`, y produce un reporte de reconciliación que compara los resultados Python contra los del Excel.

**Principio fundamental:** SQL Server almacena no solo el KPI final sino todos los valores intermedios necesarios para explicar exactamente cómo se obtuvo ese porcentaje.

---

## Architecture

```
Datos origen (Excel: RawData-PCS 244 cols + PlantActivity)
        |
        v
src/ingestion/       -- ServicioIngesta: lectura openpyxl/pandas, seriales de fecha, anomalías
        |
        v
src/staging/ + SQL   -- RepositorioStaging: copia inmutable (raw_pcs_sample, plant_activity_sample)
        |
        v
src/normalization/   -- NormalizadorPCS: ancho 244 cols -> largo 1 fila/PCS*timestamp
        |
        v
src/enrichment/      -- ServicioEnriquecimiento: join con PlantActivity (FactorOperacional, FactorExcusable)
        |
      /   \
     v     v
src/availability/    src/fault_events/
MotorDisponibilidad  MotorEventosFalla
BloquesMuestreo,     EventoFalla
BloquesRacksIndisp.
     |                    |
     +--------+-----------+
              |
              v
src/aggregation/     -- AgregacionDiaria + AgregacionAnual
              |
              v
src/persistence/     -- ServicioPersistencia -> SQL Server (append-only por IdCorrida)
                     -- tablas: proyecto, tipo_detencion, etl_run,
                     --         raw_pcs_sample, plant_activity_sample,
                     --         availability_sample_result, availability_run_result,
                     --         detencion, daily_availability, annual_availability
              |
              v
src/reconciliation/  -- ServicioReconciliacion: comparación Excel vs Python en 5 niveles
```

---

## Components and Interfaces

### src/config/

**ConfiguracionCalculo** — dataclass con todos los parámetros de una corrida.

```python
@dataclass
class ConfiguracionCalculo:
    id_corrida: str                       # UUID generado al inicio de la corrida
    version_algoritmo: str                # "availability-v1-excel-parity"
    nombre_proyecto: str                  # "Arena BESS"
    fecha_inicio_proyecto: date           # date(2026, 4, 8)
    inicio_periodo: datetime
    fin_periodo: datetime
    total_pcs: int                        # 61
    baterias_por_pcs: int                 # 4
    racks_por_pcs: int                    # 12
    total_racks: int                      # calculado: total_pcs * baterias_por_pcs * racks_por_pcs
    minutos_muestreo: int                 # 15
    solo_tiempo_operacional: bool         # equivalente a celda C21
    aplicar_evento_excusable: bool        # equivalente a celda C31
    archivo_origen: str
    zona_horaria: str                     # "America/Santiago"
```

`total_racks` siempre se calcula como `total_pcs * baterias_por_pcs * racks_por_pcs`. Nunca se hardcodea.

Valores por defecto para la fase de paridad:
```python
CONFIG_POR_DEFECTO = dict(
    version_algoritmo="availability-v1-excel-parity",
    nombre_proyecto="Arena BESS",
    fecha_inicio_proyecto=date(2026, 4, 8),
    total_pcs=61,
    baterias_por_pcs=4,
    racks_por_pcs=12,
    minutos_muestreo=15,
    zona_horaria="America/Santiago",
)
```

---

### src/ingestion/

**ServicioIngesta**

```python
class ServicioIngesta:
    def leer_raw_pcs(self, ruta: str) -> pd.DataFrame:
        \"\"\"Lee RawData-PCS desde fila 2. Conserva todas las columnas.
        Añade columna SerialFechaExcelOrigen con el serial numérico
        original de la celda de fecha.\"\"\"

    def leer_actividad_planta(self, ruta: str) -> pd.DataFrame:
        \"\"\"Lee PlantActivity desde fila 2. Retorna el DataFrame completo.\"\"\"

    def detectar_anomalias_timestamp(
        self, df: pd.DataFrame, minutos_muestreo: int
    ) -> list[ReporteAnomalia]:
        \"\"\"Detecta y reporta: huecos, duplicados, fuera de orden, frecuencia irregular.\"\"\"
```

**ReporteAnomalia** — dataclass(tipo_anomalia, numero_fila, timestamp, detalle).

Decisiones de diseño:
- `openpyxl` para acceder a los seriales numéricos de fecha antes de la conversión.
- `pandas` para manipulación posterior.
- Los seriales de fecha Excel se conservan como `float64` en `SerialFechaExcelOrigen`.
- La conversión a timestamp local usa `pytz` con zona `America/Santiago`, preservando la semántica DST de Chile.

---

### src/staging/

**RepositorioStaging**

```python
class RepositorioStaging:
    def guardar_raw_pcs_lote(self, id_corrida: str, df: pd.DataFrame) -> None:
        \"\"\"Inserta en bulk en raw_pcs_sample sin ninguna transformación.\"\"\"

    def guardar_actividad_planta_lote(self, id_corrida: str, df: pd.DataFrame) -> None:
        \"\"\"Inserta en bulk en plant_activity_sample sin ninguna transformación.\"\"\"
```

---

### src/normalization/

**NormalizadorPCS**

```python
class NormalizadorPCS:
    PATRON_COLUMNA = re.compile(
        r"Arena - PCS (\d{2}) - POWERELECTRONICS (.*)"
    )
    SUFIJOS_CAMPO = {
        "GEN3 HEx CURRENT FAULT":   "descripcion_falla_raw",
        "GEN3 HEx CURRENT STATUS":  "estado_raw",
        "GEN3 HEx CURRENT WARNING": "advertencia_raw",
        "HEM-k NUMBER OF MODULES":  "modulos_disponibles",
    }

    def normalizar(
        self, df_crudo: pd.DataFrame, config: ConfiguracionCalculo
    ) -> pd.DataFrame:
        \"\"\"Transforma formato ancho a largo. Aplica regla modulos_disponibles_nulo.\"\"\"

    def detectar_anomalias_esquema(
        self, df_crudo: pd.DataFrame, config: ConfiguracionCalculo
    ) -> list[ReporteAnomalia]:
        \"\"\"Detecta PCS faltantes, columnas faltantes/desplazadas, nombres inesperados.\"\"\"
```

Regla de normalización para módulos disponibles:
```python
if pd.isna(valor_raw) or valor_raw == "":
    modulos_disponibles = 4
    modulos_disponibles_nulo = True
else:
    modulos_disponibles = int(valor_raw)
    modulos_disponibles_nulo = False
```

---

### src/enrichment/

**ServicioEnriquecimiento**

```python
class ServicioEnriquecimiento:
    def enriquecer(
        self,
        df_normalizado: pd.DataFrame,
        df_actividad_planta: pd.DataFrame,
    ) -> pd.DataFrame:
        \"\"\"
        Join por MarcaTiempoMuestra.
        Columna C de PlantActivity -> FactorOperacional (0/1).
        Columna D de PlantActivity -> FactorExcusable (0/1).
        Huecos en el join: FactorOperacional=1, FactorExcusable=1, con advertencia.
        \"\"\"
```

---

### src/availability/

**MotorDisponibilidad** — equivalente exacto de `cmdCalcAvailability`.

```python
@dataclass
class ResultadoDisponibilidad:
    bloques_muestreo: int                          # BloquesMuestreo (C12)
    bloques_racks_indisponibles: float             # BloquesRacksIndisponibles (C14)
    disponibilidad_periodo: float | None           # 1 - C14 / (total_racks * C12)
    disponibilidad_anual_acumulada: float | None   # 1 - C14 / (total_racks * 365 * 24 * 4)
    resultados_muestra: pd.DataFrame              # filas para availability_sample_result

class MotorDisponibilidad:
    def calcular(
        self, df_enriquecido: pd.DataFrame, config: ConfiguracionCalculo
    ) -> ResultadoDisponibilidad: ...
```

**Lógica exacta del motor (transcripción del VBA a Python):**

```python
def calcular(self, df_enriquecido, config):
    bloques_muestreo = 0
    bloques_racks_indisponibles = 0.0
    filas = []

    # Ordenar por MarcaTiempoMuestra ASC, luego por NumeroPCS
    df = df_enriquecido.sort_values(["MarcaTiempoMuestra", "NumeroPCS"])
    timestamps = df["MarcaTiempoMuestra"].unique()
    timestamps.sort()

    for ts in timestamps:
        bloques_muestreo += 1  # un incremento por intervalo temporal, no por PCS
        filas_ts = df[df["MarcaTiempoMuestra"] == ts]

        for _, fila in filas_ts.iterrows():
            es_nulo = fila["ModulosDisponiblesNulo"]
            modulos = fila["ModulosDisponibles"]

            # Condición VBA: <> "" AND < 4
            # Python: not es_nulo AND modulos < baterias_por_pcs
            if not es_nulo and modulos < config.baterias_por_pcs:
                baterias_indisp = config.baterias_por_pcs - modulos

                if config.aplicar_evento_excusable:
                    ponderado = baterias_indisp * fila["FactorExcusable"]
                else:
                    ponderado = float(baterias_indisp)

                if config.solo_tiempo_operacional:
                    impacto_rack = config.racks_por_pcs * ponderado * fila["FactorOperacional"]
                else:
                    impacto_rack = config.racks_por_pcs * ponderado

                bloques_racks_indisponibles += impacto_rack
            else:
                baterias_indisp = 0
                ponderado = 0.0
                impacto_rack = 0.0

            filas.append({
                "IdCorrida":                       config.id_corrida,
                "MarcaTiempoMuestra":              ts,
                "NumeroPCS":                       fila["NumeroPCS"],
                "ModulosDisponibles":              modulos,
                "ModulosDisponiblesNulo":          es_nulo,
                "BateriasIndisponibles":           baterias_indisp,
                "FactorExcusable":                 fila["FactorExcusable"],
                "FactorOperacional":               fila["FactorOperacional"],
                "BateriasIndisponiblesPonderadas": ponderado,
                "ImpactoRackPonderado":            impacto_rack,
            })

    disponibilidad_periodo = (
        1 - bloques_racks_indisponibles / (config.total_racks * bloques_muestreo)
    ) if bloques_muestreo > 0 else None

    # Fórmula anual: supuesto 365*24*4 del Excel (conservado en fase de paridad)
    denominador_anual = config.total_racks * 365 * 24 * 4
    disponibilidad_anual_acumulada = (
        1 - bloques_racks_indisponibles / denominador_anual
    ) if denominador_anual > 0 else None

    return ResultadoDisponibilidad(
        bloques_muestreo=bloques_muestreo,
        bloques_racks_indisponibles=bloques_racks_indisponibles,
        disponibilidad_periodo=disponibilidad_periodo,
        disponibilidad_anual_acumulada=disponibilidad_anual_acumulada,
        resultados_muestra=pd.DataFrame(filas),
    )
```

---

### src/fault_events/

**MotorEventosFalla** — equivalente exacto de `mcoCreateList`.

```python
@dataclass
class EventoFalla:
    id_corrida: str
    numero_pcs: int
    marca_tiempo_inicio: datetime
    marca_tiempo_fin: datetime
    duracion_horas: float
    codigo_falla: str
    descripcion_falla: str
    descripcion_falla_fallback: bool
    promedio_baterias_involucradas: float
    horas_rack_indisponibles: float

class MotorEventosFalla:
    def detectar_eventos(
        self, df_enriquecido: pd.DataFrame, config: ConfiguracionCalculo
    ) -> list[EventoFalla]: ...
```

**Lógica exacta (transcripción del VBA):**

```python
def detectar_eventos(self, df_enriquecido, config):
    eventos = []
    fraccion_muestreo = config.minutos_muestreo / (24 * 60)  # fracción de día

    for pcs in range(1, config.total_pcs + 1):
        df_pcs = df_enriquecido[
            df_enriquecido["NumeroPCS"] == pcs
        ].sort_values("MarcaTiempoMuestra")
        evento_abierto = False
        descripcion_anterior = ""
        marca_inicio = None
        sumablocks = 0.0
        numBlock = 0
        descripcion_evento = ""
        fallback = False

        for _, fila in df_pcs.iterrows():
            ts = fila["MarcaTiempoMuestra"]
            es_nulo = fila["ModulosDisponiblesNulo"]
            modulos = fila["ModulosDisponibles"]
            descripcion = fila.get("DescripcionFallaRaw", "")

            es_falla = (not es_nulo) and (modulos < config.baterias_por_pcs)

            if es_falla:
                if not evento_abierto:
                    evento_abierto = True
                    # MarcaTiempoInicio: restar la frecuencia de muestreo (semántica VBA exacta)
                    marca_inicio = ts - timedelta(minutes=config.minutos_muestreo)

                    # Lógica de descripción con fallback
                    if descripcion == "NO FAULTS":
                        if descripcion_anterior not in ("NO FAULTS", ""):
                            descripcion_evento = descripcion_anterior
                            fallback = True
                        else:
                            descripcion_evento = "F13 NO MODULES"
                            fallback = False
                    elif descripcion == "":
                        descripcion_evento = "F1 Watchdog"
                        fallback = False
                    else:
                        descripcion_evento = descripcion
                        fallback = False

                    sumablocks = 0.0
                    numBlock = 0

                diff_baterias = config.baterias_por_pcs - modulos
                if config.aplicar_evento_excusable:
                    sumablocks += diff_baterias * fila["FactorExcusable"]
                else:
                    sumablocks += float(diff_baterias)
                numBlock += 1

            else:
                if evento_abierto:
                    marca_fin = ts
                    duracion_h = (marca_fin - marca_inicio).total_seconds() / 3600
                    promedio_bat = sumablocks / numBlock if numBlock > 0 else 0.0
                    horas_rack = config.racks_por_pcs * duracion_h * promedio_bat
                    codigo_falla = _extraer_codigo_falla(descripcion_evento)

                    eventos.append(EventoFalla(
                        id_corrida=config.id_corrida,
                        numero_pcs=pcs,
                        marca_tiempo_inicio=marca_inicio,
                        marca_tiempo_fin=marca_fin,
                        duracion_horas=duracion_h,
                        codigo_falla=codigo_falla,
                        descripcion_falla=descripcion_evento,
                        descripcion_falla_fallback=fallback,
                        promedio_baterias_involucradas=promedio_bat,
                        horas_rack_indisponibles=horas_rack,
                    ))
                    evento_abierto = False

            descripcion_anterior = descripcion

        # Cerrar evento abierto al final del período
        if evento_abierto:
            marca_fin = config.fin_periodo + timedelta(days=1)
            duracion_h = (marca_fin - marca_inicio).total_seconds() / 3600
            promedio_bat = sumablocks / numBlock if numBlock > 0 else 0.0
            horas_rack = config.racks_por_pcs * duracion_h * promedio_bat
            codigo_falla = _extraer_codigo_falla(descripcion_evento)
            eventos.append(EventoFalla(
                id_corrida=config.id_corrida,
                numero_pcs=pcs,
                marca_tiempo_inicio=marca_inicio,
                marca_tiempo_fin=marca_fin,
                duracion_horas=duracion_h,
                codigo_falla=codigo_falla,
                descripcion_falla=descripcion_evento,
                descripcion_falla_fallback=fallback,
                promedio_baterias_involucradas=promedio_bat,
                horas_rack_indisponibles=horas_rack,
            ))

    return eventos


def _extraer_codigo_falla(descripcion: str) -> str:
    idx = descripcion.find(" ")
    if idx > 0:
        return descripcion[:idx]
    return "F" + descripcion
```

---

### src/aggregation/

**AgregacionDiaria**

```python
class AgregacionDiaria:
    def calcular(
        self, resultados_muestra: pd.DataFrame, config: ConfiguracionCalculo
    ) -> pd.DataFrame:
        \"\"\"
        Agrupa ImpactoRackPonderado por día.
        Calcula BloquesRacksIndisponiblesAcumulados (suma acumulada).
        Fórmula diaria:
          Disponibilidad = 1 - acumulado / (TotalRacks * 24 * 60 * dia_N / MinutosMuestreo)
        Variacion = Disponibilidad_N - Disponibilidad_(N-1)
        \"\"\"
```

**AgregacionAnual**

```python
class AgregacionAnual:
    def calcular(
        self, resultado: ResultadoDisponibilidad, config: ConfiguracionCalculo, mes: int, anio: int
    ) -> dict:
        \"\"\"
        Fórmula anual (conservada del Excel para paridad):
          DisponibilidadAnualAcumulada = 1 - BloquesRacksIndisponibles / (TotalRacks * 365 * 24 * 4)
        Acumulación desde fecha_inicio_proyecto (2026-04-08).
        \"\"\"
```

---

### src/persistence/

**ServicioPersistencia** — usa SQLAlchemy Core con pyodbc.

```python
class ServicioPersistencia:
    def __init__(self, cadena_conexion: str): ...
    def iniciar_corrida(self, config: ConfiguracionCalculo) -> None: ...
    def completar_corrida(self, id_corrida: str) -> None: ...
    def fallar_corrida(self, id_corrida: str, error: str) -> None: ...
    def guardar_raw_pcs(self, id_corrida: str, df: pd.DataFrame) -> None: ...
    def guardar_actividad_planta(self, id_corrida: str, df: pd.DataFrame) -> None: ...
    def guardar_resultados_muestra(self, id_corrida: str, df: pd.DataFrame) -> None: ...
    def guardar_resultado_corrida(self, id_corrida: str, resultado: ResultadoDisponibilidad) -> None: ...
    def guardar_eventos_falla(self, id_corrida: str, eventos: list[EventoFalla]) -> None: ...
    def guardar_diario(self, id_corrida: str, df: pd.DataFrame) -> None: ...
    def guardar_anual(self, id_corrida: str, filas: list[dict]) -> None: ...
```

Reglas:
- Inserciones en bulk (`executemany` o `pd.to_sql` con `method="multi"`).
- Nunca `DELETE` ni `UPDATE` sobre datos de corridas anteriores.
- El `IdCorrida` se genera en `iniciar_corrida` y se propaga a todas las llamadas `guardar_*`.

---

### src/reconciliation/

**ServicioReconciliacion**

```python
@dataclass
class ResultadoNivel:
    nivel: int
    nombre: str
    aprobado: bool
    discrepancias: list[dict]

@dataclass
class ReporteReconciliacion:
    id_corrida_python: str
    archivo_excel_origen: str
    tolerancia: float
    niveles: list[ResultadoNivel]
    aprobado_global: bool

class ServicioReconciliacion:
    def reconciliar(
        self,
        id_corrida_python: str,
        ruta_excel: str,
        tolerancia: float = 1e-9,
    ) -> ReporteReconciliacion: ...
```

**Niveles de comparación:**

| Nivel | Nombre | Qué compara |
|---|---|---|
| 1 | Input | filas totales, timestamps, PCS count, frecuencia, rango de fechas |
| 2 | Muestra individual | ModulosDisponibles, BateriasIndisponibles, factores, ponderado por cada (ts, NumeroPCS) |
| 3 | Acumulados | BloquesMuestreo (C12) y BloquesRacksIndisponibles (C14) |
| 4 | KPI | DisponibilidadPeriodo (C16), DisponibilidadAnualAcumulada (C19), disponibilidad diaria y mensual |
| 5 | Eventos | NumeroPCS, MarcaTiempoInicio, MarcaTiempoFin, DuracionHoras, CodigoFalla, PromedioBateriasInvolucradas, HorasRackIndisponibles |

**Nota sobre el límite de 167 eventos:** el Excel solo ordena hasta 167 eventos en `mcoOrder`. Si el período supera ese umbral, el nivel 5 reporta advertencia y excluye la comparación de posición (ranking) para los eventos adicionales.

---

## Data Models

### SQL Server — DDL completo

```sql
-- Catalogo de tipos de detencion (fuente: hoja PCS-Fault del Excel)
CREATE TABLE tipo_detencion (
    IdTipoDetencion   INT             NOT NULL PRIMARY KEY,
    CodigoFalla       NVARCHAR(10)    NOT NULL UNIQUE,   -- 'F55'
    DescripcionFallaPE NVARCHAR(100)  NOT NULL,          -- 'Fallo externo'
    CodigoDescripcion NVARCHAR(150)   NOT NULL           -- 'F55 Fallo externo'
);

-- Datos iniciales del catalogo PCS-Fault (68 codigos)
INSERT INTO tipo_detencion VALUES
(0,   'F0',   'NO FAULT',                      'F0 NO FAULT'),
(1,   'F1',   'Watchdog',                       'F1 Watchdog'),
(2,   'F2',   'HW Vbus',                        'F2 HW Vbus'),
(3,   'F3',   'Carga suave',                    'F3 Carga suave'),
(4,   'F4',   'Descarga',                       'F4 Descarga'),
(5,   'F5',   'Alta VAC',                       'F5 Alta VAC'),
(6,   'F6',   'Baja VAC',                       'F6 Baja VAC'),
(7,   'F7',   'Alta frecuencia',                'F7 Alta frecuencia'),
(8,   'F8',   'Baja frecuencia',                'F8 Baja frecuencia'),
(10,  'F10',  'Coms CIF FPGA-DSP',              'F10 Coms CIF FPGA-DSP'),
(11,  'F11',  'Anti-isla activo',               'F11 Anti-isla activo'),
(13,  'F13',  'No modulos',                     'F13 No modulos'),
(14,  'F14',  'Drive-Select DSP',               'F14 Drive-Select DSP'),
(15,  'F15',  'Sincronizacion',                 'F15 Sincronizacion'),
(23,  'F23',  'VAC desbalanceada',              'F23 VAC desbalanceada'),
(25,  'F25',  'Baja VDC',                       'F25 Baja VDC'),
(27,  'F27',  'Fallo arranque modulos',          'F27 Fallo arranque modulos'),
(28,  'F28',  'Anti-isla pasivo',               'F28 Anti-isla pasivo'),
(31,  'F31',  'Fallo autodiagnostico',           'F31 Fallo autodiagnostico'),
(32,  'F32',  'Error modulo autodiagnostico',    'F32 Error modulo autodiagnostico'),
(33,  'F33',  'Incapaz reconectar',              'F33 Incapaz reconectar'),
(35,  'F35',  'Premag MT',                       'F35 Premag MT'),
(40,  'F40',  'Sobretemperatura interna',        'F40 Sobretemperatura interna'),
(41,  'F41',  'GFDI',                            'F41 GFDI'),
(43,  'F43',  'Paro emergencia',                 'F43 Paro emergencia'),
(44,  'F44',  'Drive-Select MCU',                'F44 Drive-Select MCU'),
(45,  'F45',  'Aislamiento general',             'F45 Aislamiento general'),
(46,  'F46',  'Fallo de datos',                  'F46 Fallo de datos'),
(47,  'F47',  'Watchdog uP',                     'F47 Watchdog uP'),
(48,  'F48',  'Comunicaciones internas',         'F48 Comunicaciones internas'),
(49,  'F49',  'IMD autodiagnostico error',       'F49 IMD autodiagnostico error'),
(51,  'F51',  'Comunicaciones PPC',              'F51 Comunicaciones PPC'),
(54,  'F54',  'Sobrecorriente',                  'F54 Sobrecorriente'),
(55,  'F55',  'Fallo externo',                   'F55 Fallo externo'),
(56,  'F56',  'Paro emergencia remoto',          'F56 Paro emergencia remoto'),
(58,  'F58',  'SW control incompatible',         'F58 SW control incompatible'),
(59,  'F59',  'SW modulo incompatible',          'F59 SW modulo incompatible'),
(62,  'F62',  'Comunicaciones DU',               'F62 Comunicaciones DU'),
(63,  'F63',  'IMD conexion tierra',             'F63 IMD conexion tierra'),
(64,  'F64',  'MAC invalida',                    'F64 MAC invalida'),
(65,  'F65',  'Sobreintensidad AC',              'F65 Sobreintensidad AC'),
(66,  'F66',  'Desbalanceo VDC modulo',          'F66 Desbalanceo VDC modulo'),
(69,  'F69',  'Sobretension DC+',                'F69 Sobretension DC+'),
(70,  'F70',  'Sobretension DC-',                'F70 Sobretension DC-'),
(71,  'F71',  'Desaturacion R1(H)',              'F71 Desaturacion R1(H)'),
(72,  'F72',  'Desaturacion R2(H)',              'F72 Desaturacion R2(H)'),
(73,  'F73',  'Desaturacion R2(L)',              'F73 Desaturacion R2(L)'),
(74,  'F74',  'Desaturacion R3(L)',              'F74 Desaturacion R3(L)'),
(75,  'F75',  'Desaturacion S1(H)',              'F75 Desaturacion S1(H)'),
(76,  'F76',  'Desaturacion S2(H)',              'F76 Desaturacion S2(H)'),
(77,  'F77',  'Desaturacion S2(L)',              'F77 Desaturacion S2(L)'),
(78,  'F78',  'Desaturacion S3(L)',              'F78 Desaturacion S3(L)'),
(79,  'F79',  'Desaturacion T1(H)',              'F79 Desaturacion T1(H)'),
(80,  'F80',  'Desaturacion T2(H)',              'F80 Desaturacion T2(H)'),
(81,  'F81',  'Desaturacion T2(L)',              'F81 Desaturacion T2(L)'),
(82,  'F82',  'Desaturacion T3(L)',              'F82 Desaturacion T3(L)'),
(83,  'F83',  'Desaturaciones',                  'F83 Desaturaciones'),
(84,  'F84',  'Comunicaciones',                  'F84 Comunicaciones'),
(85,  'F85',  'Timeout carga suave',             'F85 Timeout carga suave'),
(86,  'F86',  'Retroaviso seccionador',          'F86 Retroaviso seccionador'),
(90,  'F90',  'Temperatura IGBT HF',             'F90 Temperatura IGBT HF'),
(95,  'F95',  'Fuente PCB',                      'F95 Fuente PCB'),
(98,  'F98',  'Derivacion Id',                   'F98 Derivacion Id'),
(99,  'F99',  'Mod. Derivacion Iac',             'F99 Mod. Derivacion Iac'),
(101, 'F101', 'MOD. MEDIDA DC',                  'F101 MOD. MEDIDA DC'),
(102, 'F102', 'Corriente desbalanceada',         'F102 Corriente desbalanceada'),
(103, 'F103', 'Temperatura IGBT',                'F103 Temperatura IGBT'),
(104, 'F104', 'Temperatura PCB',                 'F104 Temperatura PCB'),
(106, 'F106', 'Mod. desbalanceo VDC',            'F106 Mod. desbalanceo VDC'),
(107, 'F107', 'Iac no alcanzada',                'F107 Iac no alcanzada'),
(108, 'F108', 'Mod. alta VDC',                   'F108 Mod. alta VDC'),
(109, 'F109', 'Mod. baja VDC',                   'F109 Mod. baja VDC'),
(112, 'F112', 'Emergencia OCAC',                 'F112 Emergencia OCAC'),
(113, 'F113', 'CRC',                             'F113 CRC'),
(114, 'F114', 'Sensado Ir',                      'F114 Sensado Ir'),
(115, 'F115', 'Sensado Is',                      'F115 Sensado Is'),
(116, 'F116', 'Sensado It',                      'F116 Sensado It'),
(118, 'F118', 'Modulo deshabilitado',            'F118 Modulo deshabilitado');

-- Tabla maestra de proyectos BESS
CREATE TABLE proyecto (
    IdProyecto          INT             NOT NULL PRIMARY KEY,
    Nombre              NVARCHAR(100)   NOT NULL,
    Estado              NVARCHAR(20)    NOT NULL DEFAULT 'por_implementar',
                        -- valores: 'en_ejecucion', 'por_implementar'
    FechaInicio         DATE            NULL,
    NumPCS              INT             NULL,
    NumBateriasPorPCS   INT             NULL,
    NumRacksPorBAC      INT             NULL,
    TotalRacks          AS (NumPCS * NumBateriasPorPCS * NumRacksPorBAC) PERSISTED,
    MinutosMuestreo     INT             NULL,
    ZonaHoraria         NVARCHAR(50)    NULL,
    Descripcion         NVARCHAR(500)   NULL
);

-- Datos iniciales de proyectos
INSERT INTO proyecto (IdProyecto, Nombre, Estado, FechaInicio, NumPCS, NumBateriasPorPCS, NumRacksPorBAC, MinutosMuestreo, ZonaHoraria) VALUES
(1, 'Arena',        'en_ejecucion',    '2026-04-08', 61, 4, 12, 15, 'America/Santiago'),
(2, 'Copiapo A',    'por_implementar', NULL,         NULL, NULL, NULL, NULL, NULL),
(3, 'Luz del Norte','por_implementar', NULL,         NULL, NULL, NULL, NULL, NULL),
(4, 'Maria Elena',  'por_implementar', NULL,         NULL, NULL, NULL, NULL, NULL);

-- Tabla de auditoria de corridas
CREATE TABLE etl_run (
    IdCorrida               NVARCHAR(36)    NOT NULL PRIMARY KEY,
    IdProyecto              INT             NULL REFERENCES proyecto(IdProyecto),
    ArchivoOrigen           NVARCHAR(500)   NOT NULL,
    SistemaOrigen           NVARCHAR(50)    NOT NULL DEFAULT 'excel',
    IniciadoEn              DATETIME2(7)    NOT NULL,
    FinalizadoEn            DATETIME2(7)    NULL,
    InicioPeriodoAnalisis   DATETIME2(7)    NOT NULL,
    FinPeriodoAnalisis      DATETIME2(7)    NOT NULL,
    MinutosMuestreo         INT             NOT NULL,
    TotalPCS                INT             NOT NULL,
    BateriasPorPCS          INT             NOT NULL,
    RacksPorPCS             INT             NOT NULL,
    TotalRacks              INT             NOT NULL,
    SoloTiempoOperacional   BIT             NOT NULL,
    AplicarEventoExcusable  BIT             NOT NULL,
    Estado                  NVARCHAR(20)    NOT NULL,
    MensajeError            NVARCHAR(MAX)   NULL,
    VersionAlgoritmo        NVARCHAR(100)   NOT NULL
);

-- Datos crudos normalizados (1 fila por PCS x timestamp)
CREATE TABLE raw_pcs_sample (
    Id                           BIGINT          IDENTITY(1,1) PRIMARY KEY,
    IdCorrida                    NVARCHAR(36)    NOT NULL REFERENCES etl_run(IdCorrida),
    MarcaTiempoMuestra           DATETIME2(7)    NOT NULL,
    SerialFechaExcelOrigen       FLOAT           NULL,
    MarcaTiempoLocalOrigen       DATETIME2(7)    NULL,
    NumeroPCS                    INT             NOT NULL,
    CodigoFallaRaw               NVARCHAR(50)    NULL,
    DescripcionFallaRaw          NVARCHAR(255)   NULL,
    EstadoRaw                    NVARCHAR(100)   NULL,
    AdvertenciaRaw               NVARCHAR(100)   NULL,
    ModulosDisponibles           INT             NOT NULL,  -- 4 si era NULL en origen
    ModulosDisponiblesNulo       BIT             NOT NULL,
    NumeroFilaOrigen             INT             NULL,
    ColumnasOrigen               NVARCHAR(MAX)   NULL       -- JSON con nombres de columnas
);
CREATE INDEX IX_raw_pcs_sample_corrida_ts_pcs
    ON raw_pcs_sample(IdCorrida, MarcaTiempoMuestra, NumeroPCS);

-- Factores de actividad de planta
CREATE TABLE plant_activity_sample (
    Id                              BIGINT          IDENTITY(1,1) PRIMARY KEY,
    IdCorrida                       NVARCHAR(36)    NOT NULL REFERENCES etl_run(IdCorrida),
    MarcaTiempoMuestra              DATETIME2(7)    NOT NULL,
    EsOperacional                   BIT             NOT NULL,
    EsEventoExcusable               BIT             NOT NULL,
    SetpointPotenciaActivaKW        FLOAT           NULL,
    DroopSobrefrecuenciaHabilitado  BIT             NULL,
    DroopBajafrecuenciaHabilitado   BIT             NULL,
    PotenciaActivaPOIKW             FLOAT           NULL,
    PorcentajeSOC                   FLOAT           NULL,
    NumeroFilaOrigen                INT             NULL
);
CREATE INDEX IX_plant_activity_corrida_ts
    ON plant_activity_sample(IdCorrida, MarcaTiempoMuestra);

-- Resultado intermedio por PCS x timestamp (auditoría completa)
CREATE TABLE availability_sample_result (
    Id                                BIGINT          IDENTITY(1,1) PRIMARY KEY,
    IdCorrida                         NVARCHAR(36)    NOT NULL REFERENCES etl_run(IdCorrida),
    MarcaTiempoMuestra                DATETIME2(7)    NOT NULL,
    NumeroPCS                         INT             NOT NULL,
    ModulosDisponibles                INT             NOT NULL,
    ModulosDisponiblesNulo            BIT             NOT NULL,
    BateriasIndisponibles             INT             NOT NULL,
    FactorExcusable                   FLOAT           NOT NULL,
    FactorOperacional                 FLOAT           NOT NULL,
    BateriasIndisponiblesPonderadas   FLOAT           NOT NULL,
    ImpactoRackPonderado              FLOAT           NOT NULL
);
CREATE INDEX IX_avail_sample_corrida_ts_pcs
    ON availability_sample_result(IdCorrida, MarcaTiempoMuestra, NumeroPCS);

-- KPI del período (una fila por corrida)
CREATE TABLE availability_run_result (
    IdCorrida                   NVARCHAR(36)    NOT NULL PRIMARY KEY REFERENCES etl_run(IdCorrida),
    BloquesMuestreo             INT             NOT NULL,        -- C12
    TotalRacks                  INT             NOT NULL,
    BloquesRacksIndisponibles   FLOAT           NOT NULL,        -- C14
    DisponibilidadPeriodo       FLOAT           NULL,            -- C16
    DisponibilidadAnualAcumulada FLOAT          NULL             -- C19
);

-- Eventos de falla consolidados
CREATE TABLE fault_event (
    Id                              BIGINT          IDENTITY(1,1) PRIMARY KEY,
    IdCorrida                       NVARCHAR(36)    NOT NULL REFERENCES etl_run(IdCorrida),
    NumeroPCS                       INT             NOT NULL,
    MarcaTiempoInicio               DATETIME2(7)    NOT NULL,
    MarcaTiempoFin                  DATETIME2(7)    NOT NULL,
    DuracionHoras                   FLOAT           NOT NULL,
    CodigoFalla                     NVARCHAR(50)    NULL,
    DescripcionFalla                NVARCHAR(255)   NULL,
    DescripcionFallaFallback        BIT             NOT NULL DEFAULT 0,
    PromedioBateriasInvolucradas    FLOAT           NOT NULL,
    HorasRackIndisponibles          FLOAT           NOT NULL
);
CREATE INDEX IX_fault_event_corrida_pcs
    ON fault_event(IdCorrida, NumeroPCS);

-- Vista operacional de detenciones
CREATE TABLE detencion (
    IdDetencion                 BIGINT          IDENTITY(1,1) PRIMARY KEY,
    IdProyecto                  INT             NOT NULL REFERENCES proyecto(IdProyecto),
    IdCorrida                   NVARCHAR(36)    NOT NULL REFERENCES etl_run(IdCorrida),
    NumeroPCS                   INT             NOT NULL,
    FechaInicio                 DATETIME2(7)    NOT NULL,
    FechaTermino                DATETIME2(7)    NOT NULL,
    DuracionSegundos            INT             NOT NULL,
    IdTipoDetencion             INT             NULL REFERENCES tipo_detencion(IdTipoDetencion),
    CodigoFalla                 NVARCHAR(10)    NULL,
    DescripcionFalla            NVARCHAR(255)   NULL,
    DescripcionFallaFallback    BIT             NOT NULL DEFAULT 0,
    PromedioBateriasInvolucradas FLOAT          NOT NULL,
    HorasRackIndisponibles      FLOAT           NOT NULL,
    ModulosDisponiblesNulo      BIT             NOT NULL DEFAULT 0,
    EsExcusable                 BIT             NOT NULL DEFAULT 0,
    Observacion                 NVARCHAR(500)   NULL,
    EstadoRevision              NVARCHAR(20)    NOT NULL DEFAULT 'pendiente'
    -- valores: 'pendiente', 'revisado', 'excluido'
);
CREATE INDEX IX_detencion_proyecto_pcs
    ON detencion(IdProyecto, NumeroPCS);
CREATE INDEX IX_detencion_corrida
    ON detencion(IdCorrida);
CREATE INDEX IX_detencion_tipo
    ON detencion(IdTipoDetencion);
CREATE INDEX IX_detencion_fechas
    ON detencion(IdProyecto, FechaInicio, FechaTermino);

-- KPI diario
CREATE TABLE daily_availability (
    Id                                  BIGINT          IDENTITY(1,1) PRIMARY KEY,
    IdCorrida                           NVARCHAR(36)    NOT NULL REFERENCES etl_run(IdCorrida),
    Dia                                 DATE            NOT NULL,
    BloquesRacksIndisponiblesDiarios    FLOAT           NOT NULL,
    BloquesRacksIndisponiblesAcumulados FLOAT           NOT NULL,
    Disponibilidad                      FLOAT           NULL,
    Variacion                           FLOAT           NULL
);
CREATE INDEX IX_daily_avail_corrida_dia
    ON daily_availability(IdCorrida, Dia);

-- KPI mensual y acumulado anual
CREATE TABLE annual_availability (
    Id                              BIGINT          IDENTITY(1,1) PRIMARY KEY,
    IdCorrida                       NVARCHAR(36)    NOT NULL REFERENCES etl_run(IdCorrida),
    Anio                            INT             NOT NULL,
    Mes                             INT             NOT NULL,
    DiasMes                         INT             NOT NULL,
    BloquesMuestreo                 INT             NOT NULL,
    BloquesRacksIndisponibles       FLOAT           NOT NULL,
    DisponibilidadMensual           FLOAT           NULL,
    BloquesMuestreoAcumulados       INT             NOT NULL,
    BloquesIndisponiblesAcumulados  FLOAT           NOT NULL,
    DisponibilidadAcumulada         FLOAT           NULL,
    DisponibilidadContractual       FLOAT           NULL
);
CREATE INDEX IX_annual_avail_corrida_anio_mes
    ON annual_availability(IdCorrida, Anio, Mes);
```

---

## Correctness Properties

*Una propiedad es una característica o comportamiento que debe mantenerse verdadero en todas las ejecuciones válidas del sistema.*

### Property 1: Normalización produce forma correcta y marca nulos consistentemente

Para cualquier DataFrame de entrada en formato ancho con N filas de timestamps y P PCS válidos, la salida normalizada debe tener exactamente N × P filas. Para toda fila donde `NUMBER_OF_MODULES` estaba vacío en el origen, `ModulosDisponibles` debe ser 4 y `ModulosDisponiblesNulo` debe ser `True`; para toda fila con valor numérico, `ModulosDisponiblesNulo` debe ser `False`.

**Validates: Requirements 2.1, 2.3, 2.4**

---

### Property 2: Detección de anomalías de timestamp es exhaustiva

Para cualquier secuencia de timestamps con anomalías inyectadas (huecos, duplicados, fuera de orden, intervalos irregulares), el módulo de detección debe reportar exactamente las posiciones y tipos de anomalías presentes, sin omisiones ni falsos positivos.

**Validates: Requirements 1.5, 1.6**

---

### Property 3: Cálculo de ImpactoRackPonderado es correcto para toda combinación de parámetros

Para cualquier combinación válida de `(ModulosDisponibles, FactorExcusable, FactorOperacional, aplicar_evento_excusable, solo_tiempo_operacional)`, el `ImpactoRackPonderado` producido por el MotorDisponibilidad debe igualar la evaluación directa de la fórmula:

- Si `ModulosDisponibles >= baterias_por_pcs` o `ModulosDisponiblesNulo`: `ImpactoRackPonderado = 0`
- Si `aplicar_evento_excusable = True` y `solo_tiempo_operacional = True`: `ImpactoRackPonderado = racks_por_pcs * (baterias_por_pcs - ModulosDisponibles) * FactorExcusable * FactorOperacional`
- Si `aplicar_evento_excusable = True` y `solo_tiempo_operacional = False`: `ImpactoRackPonderado = racks_por_pcs * (baterias_por_pcs - ModulosDisponibles) * FactorExcusable`
- Si `aplicar_evento_excusable = False` y `solo_tiempo_operacional = True`: `ImpactoRackPonderado = racks_por_pcs * (baterias_por_pcs - ModulosDisponibles) * FactorOperacional`
- Si `aplicar_evento_excusable = False` y `solo_tiempo_operacional = False`: `ImpactoRackPonderado = racks_por_pcs * (baterias_por_pcs - ModulosDisponibles)`

**Validates: Requirements 4.3, 4.4, 4.5, 4.6, 4.7**

---

### Property 4: BloquesMuestreo siempre cuenta intervalos temporales, no PCS

Para cualquier conjunto de N intervalos de tiempo con P PCS cada uno, el valor de `BloquesMuestreo` producido por el MotorDisponibilidad debe ser igual a N, independientemente de P.

**Validates: Requirements 4.1**

---

### Property 5: BloquesRacksIndisponibles es igual a la suma de todos los ImpactoRackPonderado

Para cualquier ejecución del MotorDisponibilidad, el valor de `BloquesRacksIndisponibles` (C14) debe ser igual a la suma de todos los `ImpactoRackPonderado` registrados en `availability_sample_result` para ese `IdCorrida`.

**Validates: Requirements 4.1, 4.6, 4.7, 4.11**

---

### Property 6: Los eventos de falla cubren exactamente las sub-secuencias con ModulosDisponibles < 4

Para cualquier secuencia de intervalos de un PCS, los eventos de falla detectados deben cubrir exactamente el conjunto de índices donde `ModulosDisponibles < 4` y `ModulosDisponiblesNulo = False`, sin solapamientos ni huecos respecto a las sub-secuencias consecutivas con esa condición.

**Validates: Requirements 5.1, 5.3**

---

### Property 7: La lógica de fallback de descripción es determinista y completa

Para cualquier par `(descripcion_actual, descripcion_anterior)` en el inicio de un evento de falla, el resultado de la lógica de descripción debe ser exactamente uno de los cuatro casos del árbol de decisión, y el flag `DescripcionFallaFallback` debe ser `True` si y solo si se usó la descripción del intervalo anterior.

**Validates: Requirements 5.4, 5.5**

---

### Property 8: La disponibilidad diaria acumulada es consistente con la acumulación de rack blocks

Para cualquier día N con `BloquesRacksIndisponiblesAcumulados > 0` y denominador no nulo, la disponibilidad diaria acumulada debe satisfacer `0 <= Disponibilidad <= 1` y debe ser igual a `1 - BloquesRacksIndisponiblesAcumulados / (TotalRacks * 24 * 60 * N / MinutosMuestreo)`.

**Validates: Requirements 6.3, 6.4**

---

## Error Handling

### Errores de ingesta
- Archivo no encontrado o no legible: abortar corrida, registrar en `etl_run.MensajeError`, `Estado = "failed"`.
- Hoja no encontrada en el Excel: abortar corrida con mensaje descriptivo.
- Celda vacía en columna A de RawData-PCS: registrar advertencia, continuar procesamiento.

### Errores de normalización
- Columnas faltantes o nombres inesperados: registrar advertencia con lista de columnas afectadas; los PCS sin columnas completas se procesan con los campos disponibles y se marca el resto como null.
- PCS faltantes: registrar advertencia indicando qué números de PCS no están en el archivo.

### Errores de persistencia
- Falla de conexión a SQL Server: llamar `fallar_corrida`, no dejar registros huérfanos sin `etl_run` de contexto.
- Error en inserción bulk: hacer rollback de toda la corrida; no persistir resultados parciales.

### Errores numéricos
- División por cero en DisponibilidadPeriodo (BloquesMuestreo = 0): producir `None`, no lanzar excepción.
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
- `tests/unit/test_normalization.py`: verificar casos específicos de la regla `ModulosDisponiblesNulo`; verificar detección de columna faltante con nombre concreto.
- `tests/unit/test_availability.py`: verificar un ejemplo completo con 3 intervalos y 2 PCS conocidos, comparando `BloquesMuestreo = 3`, `BloquesRacksIndisponibles` exacto.
- `tests/unit/test_fault_events.py`: verificar los cuatro casos del árbol de decisión de descripción con ejemplos concretos; verificar el caso de evento abierto al final del período.
- `tests/unit/test_aggregation.py`: verificar fórmula diaria con un ejemplo de 5 días conocidos.
- `tests/unit/test_reconciliation.py`: verificar que el reporte indica pass con datos idénticos y fail con una discrepancia conocida.

### Property-based tests (Hypothesis)

Tag de referencia para cada test: `# Feature: etl-arena-availability, Property {N}: {título}`

- **Property 1** — `tests/property/test_normalization_shape.py`:
  Genera DataFrames anchos con N timestamps y P PCS (1 <= P <= 61, 1 <= N <= 200). Verifica que la salida tiene N×P filas y que `ModulosDisponiblesNulo` es coherente con los nulos del input.

- **Property 2** — `tests/property/test_anomaly_detection.py`:
  Genera secuencias de timestamps con anomalías inyectadas en posiciones aleatorias. Verifica que cada anomalía inyectada aparece en el reporte de detección.

- **Property 3** — `tests/property/test_rack_impact.py`:
  Genera combinaciones de parámetros `(ModulosDisponibles en [0,4], FactorExcusable en {0,1}, FactorOperacional en {0,1}, aplicar_evento_excusable en {True,False}, solo_tiempo_operacional en {True,False})`. Verifica que `ImpactoRackPonderado` coincide con la evaluación directa de la fórmula.

- **Property 4** — `tests/property/test_bloques_muestreo_count.py`:
  Genera N intervalos con P PCS. Verifica que `resultado.bloques_muestreo == N`.

- **Property 5** — `tests/property/test_bloques_racks_sum.py`:
  Genera conjuntos de registros con valores conocidos. Verifica que `resultado.bloques_racks_indisponibles == sum(resultado.resultados_muestra["ImpactoRackPonderado"])`.

- **Property 6** — `tests/property/test_event_coverage.py`:
  Genera secuencias de `ModulosDisponibles` con patrones aleatorios de fallas consecutivas. Verifica que los eventos detectados cubren exactamente las sub-secuencias con `ModulosDisponibles < 4`.

- **Property 7** — `tests/property/test_fault_description_fallback.py`:
  Genera pares `(descripcion_actual, descripcion_anterior)` con valores de `"NO FAULTS"`, `""` y cadenas arbitrarias. Verifica que el resultado del árbol de decisión y el flag `DescripcionFallaFallback` son correctos.

- **Property 8** — `tests/property/test_daily_availability.py`:
  Genera secuencias de `BloquesRacksIndisponiblesDiarios` con denominador no nulo. Verifica que `Disponibilidad` satisface `0 <= Disponibilidad <= 1` y la fórmula exacta.

### Tests de integración

- `tests/integration/test_persistence.py`: verifica que después de dos corridas con `IdCorrida` distintos, ambos conjuntos de registros coexisten en la base de datos (comportamiento append-only).
- `tests/integration/test_full_pipeline.py`: ejecuta el pipeline completo sobre un subconjunto de datos de muestra extraídos del Excel real y verifica que los KPI producidos coinciden con los valores conocidos del Excel (tolerancia `1e-4` para la comparación de integración).

---

## Golden Tests — Valores Reales del Excel

Esta sección contiene valores extraídos directamente del libro fuente con `openpyxl` (`data_only=True`). Son la referencia definitiva para los tests de paridad.

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

El dict `SUFIJOS_CAMPO` del NormalizadorPCS debe usar `CURRENT STATUS` y `CURRENT WARNING`.

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
| Only Operational Time | C21 | "No" |
| Excusable Event | C31 | "No" |
| Frecuencia de muestreo | C23 | 15 min |

Resultados calculados por la macro VBA:

| Variable | Celda Excel | Nuevo nombre | Valor exacto del Excel |
|---|---|---|---|
| C12 | C12 | BloquesMuestreo | 2880 |
| C14 | C14 | BloquesRacksIndisponibles | 1211931.4879999976 |
| C16 | C16 | DisponibilidadPeriodo | 0.8562808932908321 |
| C19 | C19 | DisponibilidadAnualAcumulada | 0.9881874706814383 |

```python
# tests/golden/test_gt3_june_kpi.py
def test_june_2026_kpi_exacto():
    ESPERADO_BLOQUES_MUESTREO         = 2880
    ESPERADO_BLOQUES_RACKS_INDISP     = 1211931.4879999976
    ESPERADO_DISPONIBILIDAD_PERIODO   = 0.8562808932908321
    ESPERADO_DISPONIBILIDAD_ANUAL     = 0.9881874706814383
    TOLERANCIA                        = 1e-6

    resultado = ejecutar_motor_disponibilidad(
        inicio_periodo="2026-06-01",
        fin_periodo="2026-06-30",
        solo_tiempo_operacional=False,
        aplicar_evento_excusable=False,
    )

    assert resultado.bloques_muestreo == ESPERADO_BLOQUES_MUESTREO
    assert abs(resultado.bloques_racks_indisponibles - ESPERADO_BLOQUES_RACKS_INDISP) < TOLERANCIA
    assert abs(resultado.disponibilidad_periodo - ESPERADO_DISPONIBILIDAD_PERIODO) < TOLERANCIA
    assert abs(resultado.disponibilidad_anual_acumulada - ESPERADO_DISPONIBILIDAD_ANUAL) < TOLERANCIA
```

---

### GT-4: KPI diario — Primeros 10 días de Junio 2026

Valores extraídos de `Daily!C9:G18` (`data_only=True`):

| Día | Fecha | BloquesRacksIndisp.Diarios | BloquesRacksIndisp.Acumulados | Disponibilidad | Variacion |
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

```python
# tests/golden/test_gt4_daily.py
GOLDEN_DIARIO = [
    (1,  "2026-06-01", 4245.128,   4245.128,    0.984897512522,  0.0),
    (2,  "2026-06-02", 4226.760,   8471.888,    0.984930185564,  3.267e-3),
    (3,  "2026-06-03", 3213.880,   11685.768,   0.986142218806,  1.212e-3),
    (7,  "2026-06-07", 9465.964,   35275.308,   0.982072056742, -2.625e-3),
    (10, "2026-06-10", 6108.176,   53240.864,   0.981059005009, -3.099e-4),
    (21, "2026-06-21", 912.999,    104134.292,  0.982358635695,  7.197e-4),
]

def test_disponibilidad_diaria_coincide_con_excel():
    df_resultado = ejecutar_agregacion_diaria(periodo="june_2026")
    for dia_n, fecha_str, indisp_diario, indisp_acum, dispon, variacion in GOLDEN_DIARIO:
        fila = df_resultado[df_resultado["Dia"] == fecha_str].iloc[0]
        assert abs(fila["BloquesRacksIndisponiblesDiarios"] - indisp_diario) < 0.01
        assert abs(fila["BloquesRacksIndisponiblesAcumulados"] - indisp_acum) < 0.01
        assert abs(fila["Disponibilidad"] - dispon) < 1e-6
```

---

### GT-5: Annual_AVA — Acumulados históricos

Valores extraídos de la hoja `Annual_AVA` (`data_only=True`):

| Mes | BloquesMuestreo | BloquesRacksIndisp. | DisponibilidadMensual | BloquesMuestreoAcum. | BloquesIndisp.Acum. |
|---|---|---|---|---|---|
| Julio 2026 | 2976 | 350972.30 | 0.959721912366 | 2976 | 350972.30 |
| Agosto 2026 | 2976 | 305182.97 | 0.964976762185 | 5952 | 656155.27 |
| Sep. 2026 (parcial 1-21) | 1977 | 104134.29 | 0.982010626992 | 7929 | 760289.56 |

```python
# tests/golden/test_gt5_annual.py
GOLDEN_ANUAL = [
    (7, 2976, 350972.2999999995,  0.959721912366326, 2976, 350972.2999999995),
    (8, 2976, 305182.9679999994,  0.964976762184911, 5952, 656155.2679999989),
    (9, 1977, 104134.2920000001,  0.9820106269918267, 7929, 760289.559999999),    # Daily corregido 2026-09-24
]

def test_acumulacion_anual_coincide_con_excel():
    for mes, bloques, indisp, dispon, acc_bloques, acc_indisp in GOLDEN_ANUAL:
        fila = obtener_fila_anual(anio=2026, mes=mes)
        assert fila["BloquesMuestreo"] == bloques
        assert abs(fila["BloquesRacksIndisponibles"] - indisp) < 0.01
        assert abs(fila["DisponibilidadMensual"] - dispon) < 1e-6
        assert fila["BloquesMuestreoAcumulados"] == acc_bloques
        assert abs(fila["BloquesIndisponiblesAcumulados"] - acc_indisp) < 0.01
```

---

### GT-6: ListOfFaults — Primeros 5 eventos del PCS 1 en septiembre 2026

| # | NumeroPCS | MarcaTiempoInicio | MarcaTiempoFin | DuracionHoras | CodigoFalla | DescripcionFalla | PromedioBaterias | HorasRackIndisp. |
|---|---|---|---|---|---|---|---|---|
| 1 | 1 | 2026-09-01 00:15:00 | 2026-09-01 05:45:00 | 5.5000 | F55 | F55 EXTERNAL FAULT/OVGR | 3.3182 | 219.000 |
| 2 | 1 | 2026-09-01 08:15:00 | 2026-09-01 08:30:00 | 0.2500 | F13 | F13 NO MODULES | 0.3333 | 1.000 |
| 3 | 1 | 2026-09-01 15:45:00 | 2026-09-01 16:15:00 | 0.5000 | F55 | F55 EXTERNAL FAULT/OVGR | 2.3343 | 14.006 |
| 4 | 1 | 2026-09-01 23:45:00 | 2026-09-02 05:15:00 | 5.5000 | F13 | F13 NO MODULES | 2.6339 | 173.838 |
| 5 | 1 | 2026-09-02 09:00:00 | 2026-09-02 09:30:00 | 0.5000 | F13 | F13 NO MODULES | 0.2468 | 1.481 |

**Total HorasRackIndisponibles de todos los eventos (ListOfFaults!L10):** `26033.573000032695` (Daily corregido 2026-09-24; total de 334 eventos)

```python
# tests/golden/test_gt6_fault_events.py
# CodigoFalla = primer token (ej "F55"), DescripcionFalla = descripcion completa (ej "F55 EXTERNAL FAULT/OVGR")
# Nota: el golden actualizado tiene 334 eventos y total_horas_rack=26033.57 (correccion Daily sep-2026)
GOLDEN_EVENTOS_PCS1 = [
    # (NumeroPCS, MarcaTiempoInicio, MarcaTiempoFin, DuracionHoras, CodigoFalla, DescripcionFalla, PromedioBaterias, HorasRack)
    (1, "2026-09-01 00:15:00", "2026-09-01 05:45:00", 5.5,   "F55", "F55 EXTERNAL FAULT/OVGR", 3.3182, 219.000),
    (1, "2026-09-01 08:15:00", "2026-09-01 08:30:00", 0.25,  "F13", "F13 NO MODULES",          0.3333,   1.000),
    (1, "2026-09-01 15:45:00", "2026-09-01 16:15:00", 0.5,   "F55", "F55 EXTERNAL FAULT/OVGR", 2.3343,  14.006),
    (1, "2026-09-01 23:45:00", "2026-09-02 05:15:00", 5.5,   "F13", "F13 NO MODULES",          2.6339, 173.838),
    (1, "2026-09-02 09:00:00", "2026-09-02 09:30:00", 0.5,   "F13", "F13 NO MODULES",          0.2468,   1.481),
]
TOTAL_HORAS_RACK_INDISPONIBLES = 26033.573000032695  # actualizado con Daily corregido (2026-09-24)

def test_eventos_falla_pcs1_coincide_con_excel():
    eventos = ejecutar_motor_eventos_falla(periodo="september_2026")
    eventos_pcs1 = sorted([e for e in eventos if e.numero_pcs == 1],
                          key=lambda e: e.marca_tiempo_inicio)
    for i, (pcs, inicio, fin, dur, cod, desc, avg_bat, rack_h) in enumerate(GOLDEN_EVENTOS_PCS1):
        ev = eventos_pcs1[i]
        assert ev.codigo_falla == cod
        assert abs(ev.duracion_horas - dur) < 0.01
        assert abs(ev.promedio_baterias_involucradas - avg_bat) < 0.01
        assert abs(ev.horas_rack_indisponibles - rack_h) < 0.1

def test_total_horas_rack_coincide_con_excel():
    eventos = ejecutar_motor_eventos_falla(periodo="september_2026")
    total = sum(e.horas_rack_indisponibles for e in eventos)
    assert abs(total - TOTAL_HORAS_RACK_INDISPONIBLES) < 1e-3
```

---

### GT-7: Muestra puntual — Fila 2 de RawData-PCS

El primer registro del dataset (fila 2, timestamp `2026-04-08 00:15:00`):

| Campo | PCS | Valor crudo | ModulosDisponibles | ModulosDisponiblesNulo |
|---|---|---|---|---|
| NUMBER_OF_MODULES | 01 | 4 | 4 | False |
| NUMBER_OF_MODULES | 02 | 4 | 4 | False |
| NUMBER_OF_MODULES | 03 | 4 | 4 | False |
| NUMBER_OF_MODULES | 04 | 3.23466666666667 | 3 (o float)* | False |

*El PCS 04 tiene `3.23466666666667` en el primer bloque — valor decimal menor que 4. La normalización debe preservar el valor antes de aplicar la condición `< 4`, o convertirlo a int según la semántica VBA. Este caso requiere validación explícita.

Con `solo_tiempo_operacional=No` y `aplicar_evento_excusable=No`, `PlantActivity.C=1` y `PlantActivity.D=1`:

| PCS | BateriasIndisponibles | BateriasIndisponiblesPonderadas | ImpactoRackPonderado |
|---|---|---|---|
| 01-03 | 0 | 0 | 0 |
| 04 | 4 - 3.234... = 0.765... | 0.765... | 12 * 0.765... = 9.181... |

---

### GT-8: Corrección al dict SUFIJOS_CAMPO en el NormalizadorPCS

```python
# Corrección: en NormalizadorPCS.SUFIJOS_CAMPO usar:
SUFIJOS_CAMPO = {
    "GEN3 HEx CURRENT FAULT":   "descripcion_falla_raw",
    "GEN3 HEx CURRENT STATUS":  "estado_raw",          # <- CURRENT STATUS, no STATUS
    "GEN3 HEx CURRENT WARNING": "advertencia_raw",     # <- CURRENT WARNING, no WARNING
    "HEM-k NUMBER OF MODULES":  "modulos_disponibles",
}
```

---

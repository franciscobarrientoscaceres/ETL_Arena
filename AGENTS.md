# AGENTS.md — Reingeniería de Availability Calculation

## 1. Objetivo

Este proyecto reemplaza progresivamente el cálculo de disponibilidad realizado actualmente en Excel/VBA por un proceso reproducible en Python + SQL Server.

**Objetivo funcional:**

1. Tomar los datos crudos que actualmente alimentan `RawData-PCS` y `PlantActivity`.
2. Reproducir en Python la lógica efectiva de las macros VBA y las fórmulas Excel que participan en el KPI.
3. Obtener los mismos valores de:
   - racks indisponibles ponderados por bloque;
   - disponibilidad del período;
   - disponibilidad acumulada anual;
   - disponibilidad diaria;
   - eventos de falla y su duración.
4. Registrar los resultados y los datos intermedios relevantes en SQL Server.
5. Poder comparar una corrida Python contra una corrida Excel antes de retirar el Excel como fuente de cálculo.

**Importante:** este documento define el plan de implementación. No se debe implementar aquí el código Python definitivo.

---

## 2. Fuente de verdad y alcance

El libro fuente es:

`AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm`

El `CLAUDE.md` existente documenta que el libro tiene 11 hojas visibles, que el código VBA vive en `xl/vbaProject.bin` y que el flujo típico es:

`parámetros -> cmdCalcAvailability -> mcoCreateList -> mcoDailyAvailability -> Graphupdate`.

El libro utiliza bloques de 4 columnas por PCS en `RawData-PCS` y actualmente tiene 61 PCS, 4 baterías por PCS y 12 racks por PCS.

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

El cálculo de disponibilidad se aplica **desde el 08/Abril/2026** en adelante.
El acumulado anual y las tablas históricas deben tener esto en cuenta al inicializar la acumulación.

### Parámetros del libro Excel

| Parámetro | Celda/nombre Excel | Valor observado |
|---|---|---:|
| Total PCS | `C2` / `Total_PCS` | 61 |
| Baterías por PCS | `C3` / `Total_Batteries_per_PCS` | 4 |
| Total racks | `C11` / `Total_Racks` | 2928 |
| Frecuencia de muestreo | `C23` / `Frecuencia_de_muestreo__min` | 15 min |
| Fecha inicio | `C5` | fecha seleccionada |
| Fecha fin | `C7` | fecha seleccionada |
| Solo tiempo operacional | `C21` | Yes/No |
| Aplicar evento excusable | `C31` | Yes/No |

---

## 3. Arquitectura actual del Excel

### 3.1 Hojas relevantes

| Hoja | Función en la reingeniería |
|---|---|
| `Calculation-Availability` | Parámetros y resultado principal |
| `RawData-PCS` | Fuente cruda de estados/fallas por PCS |
| `PlantActivity` | Factor de actividad y evento excusable |
| `ListOfFaults` | Eventos de falla consolidados |
| `Daily` | KPI acumulado diario |
| `Annual_AVA` | KPI mensual/acumulado anual |
| `Graph` | Agregaciones para gráficos |
| `PCS-Fault` | Catálogo de códigos/descripciones de falla |
| `PCS-Status` | Catálogo de estados |
| `DateFormat_Correction` | Normalización auxiliar de fechas |
| `Verificación` | Controles |

Los módulos VBA de las hojas no contienen manejadores de eventos; la lógica productiva está en `Module1`, `Module2` y `Module3`.

---

# 4. Lógica exacta de las macros

## 4.1 `cmdCalcAvailability` — cálculo principal

Esta es la macro principal y debe ser tratada como la referencia primaria del KPI de disponibilidad.

### Paso 1 — limpieza

Ejecuta:

`mcoCleanTable`

La macro limpia:

`Calculation-Availability!E4:BO10000`

y:

`Calculation-Availability!C25:C29`

Luego inicializa:

- `C12 = 0` → cantidad de bloques de muestreo seleccionados.
- `C14 = 0` → acumulado de `(racks indisponibles) x bloques`.

### Paso 2 — parámetros

Lee:

- `intPCSNumber = C2`
- fecha inicio = `C5`
- fecha fin = `C7`

El recorrido de `RawData-PCS` comienza en la fila 2.

### Paso 3 — selección temporal

Una fila de `RawData-PCS` se procesa si:

```text
fecha >= fecha_inicio
AND
fecha < fecha_fin + 1 día
```

Por tanto, `C7` es inclusiva a nivel de día.

La macro termina el recorrido al encontrar una celda vacía en `RawData-PCS!A`.

### Paso 4 — salida por bloque

Para cada fila seleccionada:

- copia `RawData-PCS!A[row]` a `Calculation-Availability!E[result_row]`;
- copia `PlantActivity!D[row]` a `Calculation-Availability!BO[result_row]`.

Luego procesa cada PCS.

### Paso 5 — estructura de columnas por PCS

Cada PCS ocupa 4 columnas:

```text
PCS 1: B:E
PCS 2: F:I
PCS 3: J:M
...
```

La columna que realmente utiliza la macro para calcular disponibilidad es:

```text
4 * pcs_number + 1
```

Ejemplos:

```text
PCS 1 -> columna E
PCS 2 -> columna I
PCS 3 -> columna M
```

Los nombres de columna en `RawData-PCS` siguen el patron:

```text
Arena - PCS XX - POWERELECTRONICS GEN3 HEx CURRENT FAULT
Arena - PCS XX - POWERELECTRONICS GEN3 HEx STATUS
Arena - PCS XX - POWERELECTRONICS GEN3 HEx WARNING
Arena - PCS XX - POWERELECTRONICS HEM-k NUMBER OF MODULES
```

Donde `XX` va de `01` a `61`.

Esto corresponde al campo clave: **`Arena - PCS XX - POWERELECTRONICS HEM-k NUMBER OF MODULES`**

**Advertencia crítica:** aunque el comentario VBA menciona `CURRENT FAULT`, el algoritmo efectivo compara el valor de `NUMBER OF MODULES` contra 4. No se debe sustituir esta lógica por una interpretación basada únicamente en el código de falla.

### Paso 6 — condición de indisponibilidad

El codigo VBA ejecutable es:

```vba
If Sheet2.Cells(dblRec, 4 * intRecPCS + 1) <> "" And Sheet2.Cells(dblRec, 4 * intRecPCS + 1) < 4 Then
```

Por tanto, para cada PCS en cada bloque de 15 minutos:

| Valor de `NUMBER_OF_MODULES` | Tratamiento |
|---|---|
| Numerico < 4 | **Indisponible.** `baterias_indisponibles = 4 - NUMBER_OF_MODULES` |
| Igual a 4 | Disponible. No se acumula impacto. |
| **Vacio / nulo** | **El Excel lo ignora. Python: tratar como disponible (= 4) y marcar `modules_available_is_null = True`.** |

Decisión de negocio: asumir disponible (= 4) pero registrar el flag para identificar todos estos intervalos en la data completa.

La indisponibilidad base es:

```text
baterías_indisponibles =
    Total_Batteries_per_PCS - NUMBER_OF_MODULES
```

Con el libro actual:

```text
baterías_indisponibles = 4 - NUMBER_OF_MODULES
```

### Paso 7 — evento excusable

Si:

`Calculation-Availability!C31 = "Yes"`

entonces:

```text
baterías_indisponibles_ponderadas =
    (4 - NUMBER_OF_MODULES) * PlantActivity.D
```

Si:

`C31 = "No"`

entonces:

```text
baterías_indisponibles_ponderadas =
    4 - NUMBER_OF_MODULES
```

El factor `PlantActivity.D` es el campo `Excused Event`.

Por lo tanto, con `C31 = Yes`, un factor 0 elimina el impacto del bloque y un factor 1 lo mantiene.

### Paso 8 — solo tiempo operacional

Si:

`C21 = "No"`

se acumula directamente:

```text
C14 += 12 * baterías_indisponibles_ponderadas
```

Si:

`C21 = "Yes"`

se aplica adicionalmente:

```text
C14 +=
    12
    * baterías_indisponibles_ponderadas
    * PlantActivity.C
```

`PlantActivity.C` es el indicador:

```text
Activo = 1
Inactivo = 0
```

Por tanto, `C21 = Yes` hace que la indisponibilidad solo impacte durante períodos operacionales.

### Paso 9 — contador de bloques

Por cada fila temporal válida:

```text
C12 += 1
```

No se incrementa por PCS; se incrementa una vez por intervalo temporal.

### Paso 10 — frecuencia

La macro utiliza:

```text
12
```

como factor fijo de conversión de cada bloque de 15 minutos a horas/rack:

```text
15 minutos = 0,25 h
4 baterías/racks por PCS -> 12 como factor de referencia
```

En una futura implementación Python **no asumir ciegamente que 12 siempre será válido**. Debe derivarse/validarse contra la frecuencia y la configuración del activo, pero durante la fase de paridad se debe reproducir exactamente el comportamiento del Excel.

---

# 5. Fórmula del KPI principal

El Excel calcula:

```text
Total_Racks = 12 * Total_PCS * Total_Batteries_per_PCS
```

Con los parámetros actuales:

```text
Total_Racks = 12 * 61 * 4 = 2928
```

La disponibilidad preliminar del período es:

```text
Availability_Period =
    1 - C14 / (Total_Racks * C12)
```

con manejo de error equivalente a:

```text
IFERROR(..., "N/A")
```

En términos conceptuales:

```text
Disponibilidad =
1 -
Racks-indisponibles-acumulados-ponderados
/
Racks-totales-disponibles-en-los-bloques
```

**Esta fórmula debe conservarse sin reinterpretaciones durante la etapa de paridad.**

---

# 6. Disponibilidad acumulada anual

El Excel calcula:

```text
Accumulated Annual Availability =
1 -
Accumulated_Unavailability
/
(Total_Racks * Accumulated_15min_Blocks)
```

La hoja `Annual_AVA` acumula mensualmente:

- bloques de 15 minutos;
- `(racks unavailable) x blocks`;
- disponibilidad mensual;
- disponibilidad anual acumulada.

El valor de septiembre de `Annual_AVA!F15` está enlazado directamente a:

`Calculation-Availability!C14`.

Por lo tanto, SQL deberá poder conservar tanto el KPI mensual como el acumulado histórico.

La acumulación anual parte desde el **08/Abril/2026** (fecha de inicio del proyecto Arena BESS).

---

# 7. `mcoCleanTable`

Responsabilidad:

1. Seleccionar `Calculation-Availability`.
2. Limpiar `E4:BO10000`.
3. Limpiar `C25:C29`.

No contiene lógica de negocio.

En Python/SQL será reemplazada por una estrategia de corrida:

- `run_id`;
- carga de resultados de esa corrida;
- no borrar físicamente resultados históricos.

**No reproducir el comportamiento destructivo del Excel en SQL.**

---

# 8. `mcoCreateList` — consolidación de eventos de falla

Esta macro genera `ListOfFaults`.

### Objetivo

Transformar intervalos consecutivos donde:

```text
NUMBER_OF_MODULES < 4
```

en eventos discretos:

- PCS;
- inicio;
- fin;
- duración;
- código de falla;
- descripción;
- promedio de baterías involucradas;
- `(h unavailable) x (racks unavailable)`.

### Recorrido

Para cada PCS:

```text
intColRec = 2
```

y luego:

```text
intColRec += 4
```

Por tanto:

```text
2, 6, 10, 14, ...
```

corresponde a la columna de descripción de falla de cada PCS.

El algoritmo reinicia `dblRec = 2` para cada PCS.

### Condición para considerar un intervalo

Dentro del rango de fechas de `ListOfFaults!L2:L4`:

```text
RawData[row, intColRec + 3] <> ""
AND
RawData[row, intColRec + 3] < 4
```

`intColRec + 3` vuelve a ser la columna `NUMBER_OF_MODULES`.

### Acumulación

Para cada bloque:

```text
4 - NUMBER_OF_MODULES
```

Si `ListOfFaults!L14 = "Yes"`:

```text
sumablocks +=
    (4 - NUMBER_OF_MODULES) * PlantActivity.D
```

Si no:

```text
sumablocks +=
    4 - NUMBER_OF_MODULES
```

Luego:

```text
numBlock += 1
```

### Inicio del evento

Un nuevo evento comienza si el registro anterior:

- tiene `NUMBER_OF_MODULES = 4`, o
- está vacío, o
- está antes de la fecha inicial.

Entonces se registra:

```text
PCS =
    0.5 + intColRec / 4
```

Con `intColRec = 2`, esto produce PCS 1.

Inicio:

```text
fecha_actual -
Frecuencia_de_muestreo_min / (24 * 60)
```

La resta de una frecuencia de muestreo es intencional y debe conservarse para la paridad.

### Descripción de falla — lógica completa con fallback

El código VBA real implementa la siguiente lógica de prioridad:

```vba
' 1. Se toma la descripcion del intervalo actual
Sheet3.Cells(dblResult, 7) = Sheet2.Cells(dblRec, intColRec)

If Sheet3.Cells(dblResult, 7) = "NO FAULTS" Then
    If Sheet2.Cells(dblRec - 1, intColRec) <> "NO FAULTS" Then
        ' 2a. El anterior NO era NO FAULTS: usar descripcion del intervalo anterior
        Sheet3.Cells(dblResult, 7) = Sheet2.Cells(dblRec - 1, intColRec)
    Else
        ' 2b. El anterior tambien era NO FAULTS: asignar F13
        Sheet3.Cells(dblResult, 7) = "F13 NO MODULES"
    End If
End If

If Sheet3.Cells(dblResult, 7) = "" Then
    Sheet3.Cells(dblResult, 7) = "F1 Watchdog"
End If
```

Arbol de decision:

| Descripcion actual | Descripcion intervalo anterior | Resultado |
|---|---|---|
| Distinta de `"NO FAULTS"` y no vacia | — | Se usa el valor actual |
| `"NO FAULTS"` | distinto de `"NO FAULTS"` | **Se usa la descripcion del intervalo anterior** |
| `"NO FAULTS"` | `"NO FAULTS"` o sin anterior | `"F13 NO MODULES"` |
| `""` (vacio) | — | `"F1 Watchdog"` |

**Esta logica de fallback al intervalo anterior es un comportamiento confirmado.** Se han identificado casos de este tipo en la data de septiembre 2026 y pueden ocurrir en cualquier periodo. Todos los eventos donde se aplico el fallback deben marcarse con `fault_description_fallback = True` en `fault_event`.

Luego se calcula el código de falla a partir de la descripción final:

```text
IFERROR(
    MID(descripción, 1, FIND(" ", descripción, 1)-1),
    CONCATENATE("F", descripción)
)
```

### Fin del evento

Un evento termina si el siguiente registro:

- tiene `NUMBER_OF_MODULES = 4`, o
- está vacío, o
- supera la fecha final + 1 día.

Se calcula:

```text
duration_hours =
    24 * (fecha_fin - fecha_inicio)
```

Promedio de baterías involucradas:

```text
average_batteries =
    sumablocks / numBlock
```

Impacto:

```text
unavailability_rack_hours =
    12
    * duration_hours
    * average_batteries
```

Acumulado:

```text
ListOfFaults!L10 += unavailability_rack_hours
```

### Límite físico de eventos en el Excel

`mcoOrder` ordena el rango `P6:P172`, lo que implica un **límite de 167 eventos visibles** en `ListOfFaults`. Los eventos que superen ese límite no se incluirán en la ordenación en Excel.

Este límite es exclusivo del Excel. En la implementación Python/SQL no existe tal restricción. Se debe documentar al comparar resultados si el período analizado supera ese umbral.

### Punto importante de interpretación

`mcoCreateList` no reconstruye el evento a partir del código de falla. El criterio temporal de inicio/fin está basado en `NUMBER_OF_MODULES` y el código/descripción de falla se utiliza como información asociada al evento.

---

# 9. `mcoOrder`

Ordena `ListOfFaults` por:

```text
P6:P172
```

de mayor a menor.

La columna P contiene el impacto:

```text
(h unav.) x (racks unav.)
```

**Límite implícito:** el rango `P6:P172` cubre un máximo de **167 eventos**. Si existen más de 167 eventos en el período, los sobrantes no quedan ordenados en el Excel.

Para SQL no se debe depender de una ordenación física. La consulta/reporte debe aplicar explícitamente:

```sql
ORDER BY unavailable_rack_hours DESC
```

cuando se requiera el ranking visual equivalente.

---

# 10. `mcoDailyAvailability`

Limpia:

`Daily!D9:D39`

y recorre los resultados de `Calculation-Availability`.

Para cada día:

```text
Suma += 12 * Calculation-Availability[PCS]
```

para cada PCS.

El resultado diario se almacena en `Daily!D`.

Después:

```text
AccumulatedMonth =
    DailyValue + AccumulatedPreviousDay
```

La disponibilidad acumulada es:

```text
DailyAvailability =
1 -
AccumulatedMonth
/
(
    Total_Racks
    * 24
    * 60
    * DayNumber
    / Frecuencia_de_muestreo
)
```

La variación diaria es:

```text
Variation =
CurrentAvailability - PreviousAvailability
```

con manejo de error equivalente a `IFERROR`.

---

# 11. `Graphupdate`

`Graphupdate` no calcula el KPI primario.

Su función es de presentación:

1. ordenar el área de `Graph`;
2. copiar PCS y porcentajes hacia áreas fijas;
3. copiar tiempos para gráficos;
4. volver a ordenar.

En Python/SQL debe considerarse **capa de reporting**, no parte del motor de cálculo.

---

# 12. Fórmulas Excel que también deben reproducirse

Aunque el objetivo es migrar las macros, el resultado final depende también de fórmulas Excel.

Las fórmulas críticas identificadas son:

### Total racks

```text
12 * Total_PCS * Total_Batteries_per_PCS
```

### Frecuencia

```text
ROUND((E5-E4)*24*60,2)
```

### Disponibilidad período

```text
1 - C14 / (Total_Racks * C12)
```

### Disponibilidad anual

```text
1 - C14 / (C11 * (365*24*4))
```

### Disponibilidad diaria acumulada

```text
1 -
AccumulatedMonth
/
(
 Total_Racks
 * 24
 * 60
 * DayNumber
 / Frecuencia_de_muestreo
)
```

La fórmula anual contiene un supuesto de `365` días y `4` bloques/hora. En la reingeniería debe existir una decisión explícita sobre si se conserva exactamente por compatibilidad o si se corrige posteriormente como una versión de negocio distinta.

**Durante la fase de paridad, conservar el comportamiento actual.**

---

# 13. Diseño de datos propuesto para SQL Server

La base de datos debe separar:

1. datos crudos;
2. datos normalizados;
3. eventos;
4. resultados;
5. auditoría de ejecución.

## 13.1 `etl_run`

Campos sugeridos:

```text
run_id
source_file
source_system
started_at
finished_at
period_start
period_end
sampling_minutes
total_pcs
batteries_per_pcs
racks_per_pcs
total_racks
only_operational_time
apply_excused_event
status
error_message
algorithm_version
```

## 13.2 `raw_pcs_sample`

Modelo normalizado recomendado:

```text
run_id
sample_timestamp
pcs_number
fault_code_raw
fault_description_raw
status_raw
warning_raw
modules_available
modules_available_is_null     -- True si NUMBER_OF_MODULES estaba vacío en el origen
source_row_number
source_columns
```

El campo `modules_available_is_null` permite identificar todos los intervalos donde el dato estaba ausente. El valor de `modules_available` en esos registros se almacena como `4` para que el motor replique el comportamiento del Excel.

No mantener como diseño principal las 244 columnas repetidas.

El Excel actual tiene:

```text
61 PCS x 4 columnas = 244 columnas de PCS
```

La ETL debe convertir ese formato ancho a formato largo.

## 13.3 `plant_activity_sample`

```text
run_id
sample_timestamp
is_operational
is_excused_event
active_power_setpoint_kw
overfrequency_droop_enabled
underfrequency_droop_enabled
poi_active_power_kw
soc_percent
source_row_number
```

## 13.4 `availability_sample_result`

Una fila por:

```text
run_id + sample_timestamp + pcs_number
```

Campos:

```text
run_id
sample_timestamp
pcs_number
modules_available
modules_available_is_null     -- propagado desde raw_pcs_sample
batteries_unavailable
excused_factor
operational_factor
weighted_batteries_unavailable
weighted_rack_impact
```

Esta tabla es fundamental para auditoría y reconciliación.

## 13.5 `availability_run_result`

Una fila por corrida/período:

```text
run_id
sample_blocks
total_racks
unavailable_rack_blocks
availability_period
availability_annual
```

## 13.6 `fault_event`

```text
run_id
pcs_number
start_timestamp
end_timestamp
duration_hours
fault_code
fault_description
fault_description_fallback    -- True si la descripción fue tomada del intervalo anterior
average_batteries_involved
unavailable_rack_hours
```

## 13.7 `daily_availability`

```text
run_id
day
daily_unavailable_rack_blocks
accumulated_unavailable_rack_blocks
availability
variation
```

## 13.8 `annual_availability`

```text
run_id
year
month
days_in_month
sampling_blocks
unavailable_rack_blocks
monthly_availability
accumulated_sampling_blocks
accumulated_unavailable_rack_blocks
accumulated_availability
contractual_availability
```

---

# 14. ETL propuesta

## Fase A — Extract

Origen:

- archivos exportados de datos PCS;
- datos de actividad de planta;
- eventualmente catálogos de fallas/estado.

La extracción debe conservar una copia inmutable del dato original.

## Fase B — Raw/Staging

Guardar los registros exactamente como llegan.

Objetivo:

- trazabilidad;
- auditoría;
- posibilidad de reprocesar;
- no perder columnas nuevas del sistema origen.

## Fase C — Normalize

Transformar:

```text
PCS01 Fault / Status / Warning / Modules
PCS02 Fault / Status / Warning / Modules
...
PCS61 ...
```

a filas:

```text
timestamp | pcs_number | fault | status | warning | modules
```

No hacer todavía el cálculo de disponibilidad.

## Fase D — Enrichment

Unir por timestamp:

```text
raw_pcs_sample
    +
plant_activity_sample
```

Aplicando la misma semántica de los factores `C` y `D` de `PlantActivity`.

## Fase E — Availability Engine

Reproducir exactamente:

```text
batteries_unavailable = 4 - modules_available
```

cuando:

```text
modules_available < 4
```

Aplicar:

```text
excused_factor
```

si corresponde.

Aplicar:

```text
operational_factor
```

si corresponde.

Acumular:

```text
unavailable_rack_blocks
```

## Fase F — Event Engine

Construir eventos consecutivos según `modules_available < 4`.

Reproducir:

- inicio;
- fin;
- duración;
- promedio;
- descripción;
- código;
- impacto.

## Fase G — KPI

Calcular:

- disponibilidad período;
- disponibilidad diaria;
- acumulada;
- mensual;
- anual.

## Fase H — Persist

Insertar resultados en SQL Server con el mismo `run_id`.

## Fase I — Reconciliation

Comparar automáticamente:

```text
Excel vs Python/SQL
```

antes de declarar una corrida válida.

---

# 15. Estrategia de paridad Excel vs Python

La migración no debe validarse únicamente comparando el porcentaje final.

Para cada corrida deben compararse, como mínimo:

### Nivel 1 — input

- cantidad de filas;
- timestamps;
- cantidad de PCS;
- cantidad de registros por PCS;
- frecuencia;
- fecha inicial/final.

### Nivel 2 — muestra individual

Para cada:

```text
timestamp + PCS
```

comparar:

- modules_available;
- batteries_unavailable;
- factor excusable;
- factor operacional;
- valor ponderado.

### Nivel 3 — acumulados

Comparar:

```text
C12
C14
```

y sus equivalentes SQL.

### Nivel 4 — KPI

Comparar:

```text
C16
C19
```

y los valores diarios/mensuales.

### Nivel 5 — eventos

Comparar `ListOfFaults`:

- PCS;
- start;
- end;
- duration;
- fault code;
- average batteries;
- unavailable rack hours.

### Tolerancia numérica

Para cálculos de punto flotante se debe definir una tolerancia explícita.

Recomendación inicial:

```text
absolute tolerance <= 1e-9
```

para cálculos internos cuando sea posible, y una tolerancia operacional documentada para valores presentados.

No redondear prematuramente los cálculos intermedios.

---

# 16. Tratamiento de fechas y horario de Chile

El libro contiene una nota explícita relacionada con el cambio de horario de Chile durante septiembre de 2026.

Además, Excel utiliza números seriales de fecha.

Por ello:

1. conservar el valor temporal original;
2. conservar el timestamp interpretado;
3. definir explícitamente la zona horaria de negocio;
4. no convertir timestamps a UTC sin documentar el efecto;
5. validar especialmente períodos que atraviesen cambios de DST;
6. comparar Excel y Python en fechas afectadas por cambio de horario.

Para paridad, se recomienda conservar en staging:

```text
source_excel_serial_datetime
source_timestamp_local
```

El tipo SQL definitivo debe ser decidido antes de producción.

---

# 17. Reglas de datos que deben preservarse

## Regla 1 — orden temporal

El Excel depende de recorrer filas secuencialmente.

Python debe ordenar explícitamente:

```text
sample_timestamp ASC
```

antes de detectar eventos.

## Regla 2 — huecos

El VBA termina al encontrar una celda vacía en la fecha de `RawData-PCS`.

Python no debe copiar silenciosamente ese comportamiento como una condición de finalización.

Debe detectar y reportar:

- huecos;
- timestamps duplicados;
- timestamps fuera de orden;
- intervalos con frecuencia distinta.

## Regla 3 — bloques de 4 columnas

El formato de origen del Excel es frágil.

La ETL debe convertirlo a modelo normalizado y detectar automáticamente:

- PCS faltantes;
- columnas faltantes;
- columnas desplazadas;
- nombres de columna inesperados.

## Regla 4 — `NUMBER_OF_MODULES`

Es el campo que realmente gobierna el cálculo de indisponibilidad actual.

No sustituirlo por `FAULT_CODE` sin una nueva definición de negocio.

## Regla 5 — `NUMBER_OF_MODULES` vacío

Un intervalo con `NUMBER_OF_MODULES` vacío en `RawData-PCS` **no se acumula como indisponibilidad** (el Excel lo ignora). Python debe:

1. Tratar el valor como 4 (disponible completo) en el motor de cálculo.
2. Marcar el registro con `modules_available_is_null = True` en staging y en `availability_sample_result`.
3. **Nunca suprimir estos registros** — deben ser consultables para identificar todos los intervalos con esta condición a lo largo de toda la historia del activo.

## Regla 6 — descripción de falla con fallback

Cuando un evento de falla comienza en un intervalo cuya descripción es `"NO FAULTS"`:

- Si el intervalo inmediatamente anterior tiene una descripción distinta de `"NO FAULTS"`, **se usa la descripción del intervalo anterior** como descripción del evento.
- Si el intervalo anterior también era `"NO FAULTS"` (o no existe), se asigna `"F13 NO MODULES"`.
- Si la descripción resultante está vacía, se asigna `"F1 Watchdog"`.

Todos los eventos donde se aplicó el fallback deben marcarse con `fault_description_fallback = True`.

## Regla 7 — valores especiales de descripción

Las únicas normalizaciones automáticas de descripción son:

| Condición | Valor asignado |
|---|---|
| Descripción = `"NO FAULTS"` y anterior ≠ `"NO FAULTS"` | Descripción del intervalo anterior |
| Descripción = `"NO FAULTS"` y anterior = `"NO FAULTS"` | `"F13 NO MODULES"` |
| Descripción = `""` (vacío) | `"F1 Watchdog"` |

Estas normalizaciones aplican **solo al motor de paridad de eventos** (`mcoCreateList`).

---

# 18. Arquitectura Python propuesta

No implementar todavía, pero diseñar el proyecto alrededor de módulos conceptuales:

```text
src/
  config/
  ingestion/
  staging/
  normalization/
  enrichment/
  availability/
  fault_events/
  aggregation/
  persistence/
  reconciliation/
  reporting/
tests/
sql/
```

Separar explícitamente:

```text
ingestion != business logic != persistence
```

El motor de disponibilidad debe poder recibir DataFrames/estructuras normalizadas y producir resultados sin depender de Excel.

---

# 19. Configuración

No hardcodear:

```text
61
4
12
15
```

como constantes de negocio.

Deben provenir de una configuración/versionado de cálculo.

Pero para la primera versión de paridad los valores por defecto deben ser:

```text
project_name         = "Arena BESS"
project_start_date   = "2026-04-08"
total_pcs            = 61
batteries_per_pcs    = 4
racks_per_pcs        = 12
sampling_minutes     = 15
```

La fórmula equivalente a `Total_Racks` debe ser:

```text
total_pcs * batteries_per_pcs * racks_per_pcs
```

---

# 20. Versionamiento del algoritmo

Cada resultado SQL debe guardar:

```text
algorithm_version
```

Ejemplo:

```text
availability-v1-excel-parity
```

Cuando posteriormente se modifique una regla de negocio, crear otra versión.

Nunca sobrescribir resultados históricos con una nueva lógica.

---

# 21. Plan de implementación por etapas

## Etapa 1 — Reverse engineering

Entregables:

- inventario de macros;
- mapa de hojas;
- mapa de celdas;
- reglas de cálculo;
- catálogo de datos;
- fórmulas equivalentes;
- riesgos.

**Estado:** completado para la primera versión de análisis.

## Etapa 2 — Data contract

Definir:

- formato de archivos origen;
- columnas;
- tipos;
- timestamps;
- identificador PCS;
- reglas de duplicados;
- frecuencia;
- tratamiento DST.

## Etapa 3 — SQL Server

Crear:

- staging;
- raw normalized;
- plant activity;
- sample results;
- fault events;
- KPI results;
- audit/run tables.

## Etapa 4 — ETL

Implementar:

```text
extract -> validate -> stage -> normalize -> enrich
```

sin calcular KPI todavía.

## Etapa 5 — Motor de disponibilidad

Implementar el equivalente de `cmdCalcAvailability`.

## Etapa 6 — Motor de eventos

Implementar el equivalente de `mcoCreateList`.

## Etapa 7 — Agregaciones

Implementar equivalentes de:

- `mcoDailyAvailability`;
- `Annual_AVA`;
- agregaciones de `Graph`.

## Etapa 8 — Reconciliation

Crear pruebas automatizadas Excel vs Python.

## Etapa 9 — Shadow mode

Durante un período definido:

```text
Excel continúa siendo oficial
Python/SQL calcula en paralelo
```

Comparar resultados.

## Etapa 10 — Producción

Cuando la paridad sea aceptable:

```text
origen -> ETL -> Python -> SQL Server -> reporting
```

Excel pasa a ser herramienta de consulta/legacy, no motor oficial.

---

# 22. Criterios de aceptación

La migración no se considera terminada hasta que:

1. Una corrida Python pueda reproducir la corrida Excel usando exactamente el mismo input.
2. `C12` sea reproducible.
3. `C14` sea reproducible.
4. `C16` sea reproducible.
5. `C19` sea reproducible.
6. Los eventos de `ListOfFaults` sean reconciliables.
7. La disponibilidad diaria sea reconciliable.
8. La disponibilidad acumulada anual sea reconciliable.
9. Los períodos que atraviesan cambio de horario sean validados.
10. Una corrida SQL pueda ser auditada hasta:
    `KPI -> acumulado -> muestra -> input raw`.
11. Todos los intervalos con `modules_available_is_null = True` estén identificados y accesibles en la base de datos.

---

# 23. Riesgos detectados

### Riesgo A — comentarios VBA no coinciden completamente con el comportamiento

El comentario identifica una columna como `CURRENT FAULT`, pero la lógica efectiva utiliza la columna de `NUMBER_OF_MODULES`.

**Acción:** priorizar el código ejecutable sobre comentarios.

### Riesgo B — fórmula anual fija

El Excel utiliza:

```text
365 * 24 * 4
```

para el KPI anual.

**Acción:** mantener durante paridad; posteriormente definir con negocio si se requiere calendario real.

### Riesgo C — factor fijo 12

El motor utiliza `12` repetidamente.

**Acción:** reproducir inicialmente; luego parametrizar con una definición explícita.

### Riesgo D — fechas seriales Excel

Puede haber diferencias de precisión y DST.

**Acción:** conservar serial original y timestamp interpretado.

### Riesgo E — modelo ancho

`RawData-PCS` usa bloques repetidos de 4 columnas.

**Acción:** normalizar a modelo largo antes del cálculo.

### Riesgo F — Excel destruye resultados anteriores

`mcoCleanTable` y `mcoCleanList` limpian rangos.

**Acción:** SQL debe ser append-only por `run_id`.

### Riesgo G — `NUMBER_OF_MODULES` vacío silencioso

El Excel ignora silenciosamente los intervalos con `NUMBER_OF_MODULES` vacío. La frecuencia de este evento en el histórico completo es desconocida.

**Acción:** marcar con `modules_available_is_null = True`; incluir en el reporte de calidad de datos de cada corrida.

### Riesgo H — descripción de falla con fallback al intervalo anterior

La macro `mcoCreateList` puede asignar a un evento la descripción del intervalo anterior cuando el intervalo de inicio reporta `"NO FAULTS"`. Esto se ha confirmado en datos de septiembre 2026 y puede ocurrir en cualquier período.

**Acción:** marcar con `fault_description_fallback = True`; reportar la frecuencia de este caso en cada corrida.

---

# 24. No hacer

Durante la primera implementación NO:

- modificar el algoritmo de disponibilidad;
- corregir supuestos de negocio;
- reinterpretar códigos de falla;
- cambiar la fórmula anual;
- cambiar el tratamiento de DST;
- redondear resultados intermedios;
- eliminar tablas de staging;
- reemplazar el Excel sin período de shadow mode;
- implementar primero el dashboard;
- comenzar por optimización de performance antes de obtener paridad.

Primero:

```text
entender -> reproducir -> reconciliar -> versionar -> optimizar
```

---

# 25. Resultado esperado de la reingeniería

La arquitectura final debe permitir:

```text
                    ┌──────────────────┐
                    │ Datos origen     │
                    └────────┬─────────┘
                             │
                             ▼
                    ┌──────────────────┐
                    │ ETL / Staging    │
                    └────────┬─────────┘
                             │
                             ▼
                    ┌──────────────────┐
                    │ Normalización    │
                    └────────┬─────────┘
                             │
                 ┌───────────┴───────────┐
                 ▼                       ▼
        ┌─────────────────┐     ┌─────────────────┐
        │ Availability    │     │ Fault Events    │
        │ Engine          │     │ Engine          │
        └────────┬────────┘     └────────┬────────┘
                 │                       │
                 └───────────┬───────────┘
                             ▼
                    ┌──────────────────┐
                    │ SQL Server       │
                    │ KPI + Audit      │
                    └────────┬─────────┘
                             │
                             ▼
                    ┌──────────────────┐
                    │ Reporting / BI   │
                    └──────────────────┘
```

El principio fundamental es que SQL Server almacene no solo el porcentaje final, sino también los resultados intermedios necesarios para explicar exactamente por qué se obtuvo ese porcentaje.

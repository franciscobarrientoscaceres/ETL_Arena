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

> **Auditoría 2026-09-24:** varias reglas de este documento fueron corregidas contra el VBA real y las fórmulas del libro. El detalle y la evidencia están en [`.kiro/specs/etl-arena-availability/audit.md`](./.kiro/specs/etl-arena-availability/audit.md) (hallazgos `F-xx`). Ante cualquier diferencia, mandan `requirements.md` / `design.md` revisión 2.

---

## 2. Fuente de verdad y alcance

El libro fuente es:

`AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm`

El `CLAUDE.md` existente documenta que el libro tiene 11 hojas visibles, que el código VBA vive en `xl/vbaProject.bin` y que el flujo típico es:

`parámetros -> cmdCalcAvailability -> mcoCreateList -> mcoDailyAvailability -> Graphupdate`.

El libro utiliza bloques de 4 columnas por PCS en `RawData-PCS` y actualmente tiene 61 PCS, 4 baterías (BAC) por PCS y 12 racks por BAC (48 racks por PCS).

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

- `BloquesMuestreo (C12) = 0` → cantidad de bloques de muestreo seleccionados.
- `BloquesRacksIndisponibles (C14) = 0` → acumulado de `(racks indisponibles) x bloques`.

### Paso 2 — parámetros

Lee:

- `intPCSNumber = C2`
- fecha inicio = `C5`
- fecha fin = `C7`

El recorrido de `RawData-PCS` comienza en la fila 2.

### Paso 3 — selección temporal

Una fila de `RawData-PCS` se procesa si:

```text
fecha >= FechaInicio
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
4 * numero_pcs + 1
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
Arena - PCS XX - POWERELECTRONICS GEN3 HEx CURRENT STATUS
Arena - PCS XX - POWERELECTRONICS GEN3 HEx CURRENT WARNING
Arena - PCS XX - POWERELECTRONICS HEM-k NUMBER OF MODULES
```

Donde `XX` va de `01` a `61`.

**Unión con `PlantActivity` (F-01):** el VBA lee `PlantActivity` con el **mismo número de fila** (`Sheet4.Cells(dblRec, 3|4)`), no por timestamp. Una celda vacía vale 0. En el libro actual las filas 15333–15991 de `PlantActivity` no tienen timestamp.

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
| Numerico < 4 (puede ser **fraccionario**, p. ej. 3,2347 — F-02) | **Indisponible.** `baterias_indisponibles = 4 - NUMBER_OF_MODULES` (float, sin truncar) |
| Igual a 4 | Disponible. No se acumula impacto. |
| **Vacio / nulo** | **El Excel lo ignora. Python: tratar como disponible (= 4) y marcar `ModulosDisponiblesNulo = True`.** |

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

No se incrementa por PCS; se incrementa una vez por **fila** de `RawData-PCS` en rango (si hay timestamps repetidos, cuenta cada fila — F-07).

**Semántica de flags (F-12):** el factor operacional se aplica si `C21 <> "No"`; el excusable solo si `C31 = "Yes"` (comparación exacta).

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

**Nota (F-20):** `Annual_AVA` se mantiene a mano: jul/ago son valores tipeados, `D15` (días de sep) está tipeado como `20+14.25/24`, y la acumulación empieza en **julio 2026** (no en abril). Su acumulado `J = 1 - I/(Racks·H)` es una métrica distinta de `C19`.

La hoja `Annual_AVA` acumula mensualmente:

- bloques de 15 minutos;
- `(racks unavailable) x blocks`;
- disponibilidad mensual;
- disponibilidad anual acumulada.

El valor de septiembre de `Annual_AVA!F15` está enlazado directamente a:

`Calculation-Availability!BloquesRacksIndisponibles (C14)`.

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

- `IdCorrida`;
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

El algoritmo reinicia `dblRec = 2` para cada PCS, pero **no** reinicia `sumablocks`, `numBlock` ni `dblResult` (ver "Defecto de arrastre" más abajo).

**Parámetros propios (F-05):** `mcoCreateList` no lee `C5`/`C7`/`C31`. Usa `ListOfFaults!L2` (inicio), `L4` (fin) y `L14` (evento excusable), que son **valores independientes**. En el libro actual `L14 = "Yes"` mientras `C31 = "No"`.

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
| `"NO FAULTS"` | distinto de `"NO FAULTS"` y no vacía | **Se usa la descripcion del intervalo anterior** |
| `"NO FAULTS"` | vacía (`""`) | Se copia `""` → **`"F1 Watchdog"`** (F-04) |
| `"NO FAULTS"` | `"NO FAULTS"` | `"F13 NO MODULES"` |
| `""` (vacio) | — | `"F1 Watchdog"` |

La fila "anterior" es la fila anterior de la **hoja completa** (puede estar fuera del período). Si el evento empieza en la fila 2, la anterior es el encabezado y la macro falla con *Type mismatch* (F-11).

**Esta logica de fallback al intervalo anterior es un comportamiento confirmado.** Se han identificado casos de este tipo en la data de septiembre 2026 y pueden ocurrir en cualquier periodo. Todos los eventos donde se aplico el fallback deben marcarse con `descripcion_falla_fallback = True` en `fault_event`.

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
- supera **estrictamente** la fecha final + 1 día (`>`).

La fecha fin registrada es la de la **fila actual** (la última en falla), no la de la fila siguiente (F-03).

**Defecto de arrastre (F-06):** si el período termina a las 23:45 y la fila siguiente es exactamente `L4 + 1` 00:00 con el PCS aún en falla, el evento no se cierra; los acumuladores pasan al siguiente PCS, que sobrescribe esa fila de `ListOfFaults`. Python lo replica en v1 y lo marca `EventoArrastradoExcel`.

Se calcula:

```text
duracion_horas =
    24 * (fecha_fin - FechaInicio)
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
    * duracion_horas
    * average_batteries
```

Acumulado:

```text
ListOfFaults!L10 += unavailability_rack_hours
```

### Sin límite de eventos (corregido — F-18)

La lista de eventos (`ListOfFaults!B:I`) no tiene tope (se limpia hasta la fila 150000) y **no se ordena**: queda en orden de escritura (PCS por PCS). El rango `N5:Q172` que ordena `mcoOrder` es la tabla resumen por **código de falla** (167 códigos), no la lista de eventos.

### Punto importante de interpretación

`mcoCreateList` no reconstruye el evento a partir del código de falla. El criterio temporal de inicio/fin está basado en `NUMBER_OF_MODULES` y el código/descripción de falla se utiliza como información asociada al evento.

---

# 9. `mcoOrder`

Ordena la tabla resumen `ListOfFaults!N5:Q172` por la columna `P` de mayor a menor.

Esa tabla tiene una fila por código del catálogo `PCS-Fault` (167 códigos): `N` = código, `O` = descripción PE, `P = SUMIF(F:F, N, I:I)` (rack-hours por código), `Q = P / L10`. La lista de eventos `B:I` no se ordena (F-18).

Para SQL no se debe depender de una ordenación física. La consulta/reporte debe aplicar explícitamente:

```sql
ORDER BY horas_rack_indisponibles DESC
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

**Importante (F-08, F-09):** los valores de `Calculation-Availability[PCS]` son baterías ponderadas por el factor excusable pero **sin** el factor operacional, así que con `C21 = "Yes"` la suma diaria no coincide con C14. `DayNumber` es el índice del día desde `C5` (`Daily!B`), y el rango de días llega hasta `Daily!D5` (valor propio, máx. 31 días).

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

## 13.0 `proyecto`

Tabla maestra de proyectos BESS. Se crea una sola vez y no cambia entre corridas.

```text
IdProyecto
nombre
estado                -- 'en_ejecucion' | 'por_implementar'
FechaInicio
NumPCS
NumBateriasPorPCS
NumRacksPorBAC
TotalRacks           -- columna calculada: NumPCS * NumBateriasPorPCS * NumRacksPorBAC
minutos_muestreo
ZonaHoraria
descripcion
```

Proyectos registrados:

| id | nombre | estado |
|---|---|---|
| 1 | Arena | en_ejecucion |
| 2 | Copiapó A | por_implementar |
| 3 | Luz del Norte | por_implementar |
| 4 | María Elena | por_implementar |

## 13.0b `tipo_detencion`

Catálogo de códigos de falla, extraído de la hoja `PCS-Fault` del Excel: **167 registros** (F0…F257), con columnas `Meaning` (criticidad) y `Operative` (F-19). El seed se genera desde la hoja con un script; no se hardcodea.

```text
IdTipoDetencion     -- código numérico (ej: 55)
codigo_falla            -- código texto (ej: 'F55')
DescripcionFallaPE  -- descripción PowerElectronics (ej: 'Fallo externo')
CodigoDescripcion      -- código + descripción (ej: 'F55 Fallo externo')
```

## 13.1 `etl_run`

Campos sugeridos:

```text
IdCorrida
archivo_origen
SistemaOrigen
IniciadoEn
FinalizadoEn
inicio_periodo
fin_periodo
minutos_muestreo
TotalPCS
baterias_por_pcs
racks_por_pcs
TotalRacks
solo_tiempo_operacional
aplicar_evento_excusable
status
MensajeError
version_algoritmo
```

## 13.2 `raw_pcs_sample`

Modelo normalizado recomendado:

```text
IdCorrida
MarcaTiempoMuestra
numero_pcs
codigo_falla_raw
descripcion_falla_raw
estado_raw
advertencia_raw
ModulosDisponibles
ModulosDisponiblesNulo     -- True si NUMBER_OF_MODULES estaba vacío en el origen
NumeroFilaOrigen
ColumnasOrigen
```

El campo `ModulosDisponiblesNulo` permite identificar todos los intervalos donde el dato estaba ausente. El valor de `ModulosDisponibles` en esos registros se almacena como `4` para que el motor replique el comportamiento del Excel.

No mantener como diseño principal las 244 columnas repetidas.

El Excel actual tiene:

```text
61 PCS x 4 columnas = 244 columnas de PCS
```

La ETL debe convertir ese formato ancho a formato largo.

## 13.3 `plant_activity_sample`

```text
IdCorrida
MarcaTiempoMuestra
EsOperacional
EsEventoExcusable
SetpointPotenciaActivaKW
overfrequency_droop_enabled
underfrequency_droop_enabled
PotenciaActivaPOIKW
PorcentajeSOC
NumeroFilaOrigen
```

## 13.4 `availability_sample_result`

Una fila por:

```text
IdCorrida + MarcaTiempoMuestra + numero_pcs
```

Campos:

```text
IdCorrida
MarcaTiempoMuestra
numero_pcs
ModulosDisponibles
ModulosDisponiblesNulo     -- propagado desde raw_pcs_sample
baterias_indisponibles
factor_excusable
factor_operacional
baterias_indisponibles_ponderadas
impacto_rack_ponderado
```

Esta tabla es fundamental para auditoría y reconciliación.

## 13.5 `availability_run_result`

Una fila por corrida/período:

```text
IdCorrida
BloquesMuestreo
TotalRacks
BloquesRacksIndisponibles
disponibilidad_periodo
disponibilidad_anual_acumulada
```

## 13.6 `fault_event`

```text
IdCorrida
numero_pcs
marca_tiempo_inicio
marca_tiempo_fin
duracion_horas
codigo_falla
descripcion_falla
descripcion_falla_fallback    -- True si la descripción fue tomada del intervalo anterior
promedio_baterias_involucradas
horas_rack_indisponibles
```

## 13.7 `daily_availability`

```text
IdCorrida
day
BloquesRacksIndisponiblesDiarios
BloquesRacksIndisponiblesAcumulados
availability
variation
```

## 13.8 `annual_availability`

```text
IdCorrida
year
month
DiasMes
BloquesMuestreo
BloquesRacksIndisponibles
DisponibilidadMensual
BloquesMuestreoAcumulados
BloquesRacksIndisponiblesAcumulados
DisponibilidadAcumulada
DisponibilidadContractual
```

## 13.9 `detencion`

Vista operacional de detenciones. Poblada desde `fault_event` pero orientada a consultas de negocio y análisis de fallas. Coexiste con `fault_event` (tabla técnica de auditoría).

```text
IdDetencion
IdProyecto                 -- FK -> proyecto
IdCorrida                      -- FK -> etl_run
numero_pcs
FechaInicio
FechaTermino
DuracionSegundos           -- DATEDIFF(seconds, FechaInicio, FechaTermino)
IdTipoDetencion           -- FK -> tipo_detencion (NULL si código no existe en catálogo)
codigo_falla
descripcion_falla
descripcion_falla_fallback
promedio_baterias_involucradas
horas_rack_indisponibles
ModulosDisponiblesNulo   -- flag de calidad de dato
EsExcusable                -- flag de evento excusable
Observacion                 -- campo libre para anotaciones operacionales
EstadoRevision             -- 'pendiente' | 'revisado' | 'excluido'
```

La diferencia entre `fault_event` y `detencion`:

| `fault_event` | `detencion` |
|---|---|
| Tabla técnica de auditoría | Tabla operacional de negocio |
| Append-only por IdCorrida | Permite UPDATE en Observacion y EstadoRevision |
| Sin FK a proyecto | Con FK a proyecto |
| Sin DuracionSegundos | Con DuracionSegundos calculado |
| Sin campos de workflow | Con EstadoRevision y Observacion |

---

# 14. ETL propuesta

## Fase S — SCADA acquisition (Fase del lunes)

La fuente cruda de `RawData-PCS` es el **server SCADA**. Cada lunes se exporta un reporte desde allí y se lleva a esta PC. Esta fase es **previa a Extract** y no depende del motor de disponibilidad.

### Flujo

1. **~03:00 AM lunes** — ventana libre en SCADA. Exportar solo; **no procesar en el server**.
2. **Obtener el archivo** por TeamViewer (file transfer) a `data/inbox/raw_pcs_<corte>.<ext>`.
3. **`acquire-wait`** (etapa 1 de `scripts/run_lunes.py`): espera/valida el archivo en inbox (nombre, rango de fechas, sha256, no vacío).
4. **`scada_adapter` / `prepare-workbook`** (etapa 2): parsea fechas origen `mm-dd-aaaa hh:mm:ss` y las escribe como **fecha/serial Excel real** (formato visual `dd-mm-aaaa hh:mm:ss`; escribirlas como texto rompe el filtro de la macro — F-21), aplica mapping de columnas = `RawData-PCS`, escribe **vía COM** en copia de trabajo del `.xlsm` (backup del original) y valida la alineación por fila con `PlantActivity`.
5. **`run-macros`** (etapa 3): macros en PC local vía COM Excel: `cmdCalcAvailability` → `mcoCreateList` → `mcoDailyAvailability` → `Graphupdate`, seteando `C5`/`C7`/`C21`/`C31`, `ListOfFaults!L2`/`L4`/`L14` y `Daily!D5`; extrae la referencia completa (C12/C14/C16/C19/C23, tabla de resultados, `ListOfFaults` completa, `Daily`).
6. **`run-etl`** (etapa 4): pipeline Python completo → SQL Server (`IdCorrida` nuevo por lunes).
7. **`reconcile`** (etapa 5): Python vs referencia del paso 5.
8. **`notify-bi`** (etapa 6): notificar a Misael / webhook → refresh Power BI (sin service principal aún).

### Reglas

| Regla | Detalle |
|---|---|
| Fechas del reporte | **Las setea quien exporta** (manual). Export **incremental** (D-07): `DESDE` = dato siguiente al último cargado (hoy 2026-09-21 14:15), `HASTA` = último dato disponible ese lunes. El adapter **valida la continuidad** con el libro base (hueco → confirmar; solape distinto → rechazar). |
| Transporte hoy | Solo TeamViewer. **Sin UNC / share / API** al SCADA. Pedir share o API GPM a GPM como mejora (riesgo J). |
| Formato origen | Probable CSV; columnas = `RawData-PCS`. Fechas origen `mm-dd-aaaa hh:mm:ss` → **fecha/serial Excel real** con formato visual `dd-mm-aaaa hh:mm:ss` (transformación que hoy hace Alex a mano; escribirla como texto rompe el filtro de la macro — F-21). |
| PlantActivity | **No** se actualiza desde SCADA. Fuente aparte; se une por fila (F-01): desalineaciones y filas sin timestamp se reportan en calidad de corrida (F-31). |
| Power BI | Owner: **Misael**. Modo **notificación** hasta que haya service principal/API; no integración de API en P0–P9. |
| Server SCADA | Solo extracción. **Prohibido** correr macros o ETL allí. |
| IdCorrida | Cada lunes genera un `IdCorrida` nuevo; append-only en SQL. |

### Estructura de carpetas

```text
data/
  inbox/         # drop zone: archivo crudo recién bajado
  processed/     # originales inmutables + sha256 (Fase A/B)
  work/          # copia de trabajo del .xlsm para macros/ETL
```

### Componentes (ver design.md / tasks 15)

- `scripts/run_lunes.py` — orquestador único con etapas: `acquire-wait`, `prepare-workbook`, `run-macros`, `run-etl`, `reconcile`, `notify-bi`.
- `src/etl_arena/acquisition/` — `acquire_wait`, `contrato_scada`, `lector_scada` (validación de rango, transformador de fechas, mapping de columnas).
- `src/etl_arena/workbook/` — sesión COM, `preparar` (escritura en la copia de trabajo), `macros` (las 4 macros) y `referencia` (extracción); solo PC local con Windows + Excel.

### Fases por etapa (P0–P9) y etiquetas S/T/M/E/R/B

| Fase | Nombre | Contenido |
|---|---|---|
| **P0** | Muestra lunes | Muestra real del CSV/headers/formato de fecha/delimiter desde SCADA (bloqueante). |
| **P1** | Data contract SCADA | Formalizar contract tras P0: columnas, tipos, fecha, delimiter, encoding. |
| **P2** | inbox + acquire-wait | `data/inbox`, validación de archivo, sha256, logging. |
| **P3** | scada_adapter | Transformación fechas + mapping columnas + backup `.xlsm`. |
| **P4** | run-macros (M) | Runner COM de `cmdCalcAvailability` → `mcoCreateList` → `mcoDailyAvailability` → `Graphupdate`; extracción C12/C14/C16/C19. |
| **P5** | run-etl (E) | Pipeline Python → SQL con nuevo `IdCorrida`. |
| **P6** | reconcile (R) | Python vs referencia Excel del P4. |
| **P7** | notify-bi (B) | Notificación de refresh PBI (modo notificación, Misael). |
| **P8** | fallback / RPA | Solo si TeamViewer falla: RPA TeamViewer como fallback, **no** camino crítico. |
| **P9** | docs | Runbook + handoff. |

Etapas del orquestador: **S** (acquire) → **T** (transform/prepare) → **M** (macros) → **E** (ETL) → **R** (reconcile) → **B** (BI notify).

---

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
timestamp | numero_pcs | fault | status | warning | modules
```

No hacer todavía el cálculo de disponibilidad.

## Fase D — Enrichment

Unir por **número de fila origen** (`NumeroFilaOrigen`), no por timestamp (F-01):

```text
raw_pcs_sample
    +
plant_activity_sample
```

Aplicando la misma semántica de los factores `C` y `D` de `PlantActivity`: celda vacía = 0. Si `PlantActivity` trae timestamp en esa fila y no coincide con el de `RawData-PCS`, se reporta como desalineación (no se corrige en modo paridad).

## Fase E — Availability Engine

Reproducir exactamente:

```text
baterias_indisponibles = 4 - ModulosDisponibles
```

cuando:

```text
ModulosDisponibles < 4
```

Aplicar:

```text
factor_excusable
```

si corresponde.

Aplicar:

```text
factor_operacional
```

si corresponde.

Acumular:

```text
BloquesRacksIndisponibles
```

## Fase F — Event Engine

Construir eventos consecutivos según `ModulosDisponibles < 4`.

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

Insertar resultados en SQL Server con el mismo `IdCorrida`.

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

- ModulosDisponibles;
- baterias_indisponibles;
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

Las tolerancias por magnitud (C14, C16/C19, Daily, eventos, L10) están en una **única tabla**: `design.md §Tolerancias`; `golden_index._meta` se sincroniza con ella (F-26). No definir tolerancias en otro lugar.

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
SerialFechaExcelOrigen
MarcaTiempoLocalOrigen
```

El tipo SQL definitivo debe ser decidido antes de producción.

---

# 17. Reglas de datos que deben preservarse

## Regla 1 — orden temporal

El Excel depende de recorrer filas secuencialmente.

En modo paridad Python debe recorrer las filas **en el orden de fila origen** (como el VBA), sin reordenar ni deduplicar, y **reportar** cualquier fila fuera de orden `MarcaTiempoMuestra ASC` como anomalía (F-07).

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
2. Marcar el registro con `ModulosDisponiblesNulo = True` en staging y en `availability_sample_result`.
3. **Nunca suprimir estos registros** — deben ser consultables para identificar todos los intervalos con esta condición a lo largo de toda la historia del activo.

## Regla 6 — descripción de falla con fallback

Cuando un evento de falla comienza en un intervalo cuya descripción es `"NO FAULTS"`:

- Si el intervalo inmediatamente anterior tiene una descripción distinta de `"NO FAULTS"`, **se usa la descripción del intervalo anterior** como descripción del evento (si esa descripción es vacía, el resultado es `"F1 Watchdog"`).
- Si el intervalo anterior también era `"NO FAULTS"`, se asigna `"F13 NO MODULES"`.
- Si la descripción resultante está vacía, se asigna `"F1 Watchdog"`.

Todos los eventos donde se aplicó el fallback deben marcarse con `descripcion_falla_fallback = True`.

## Regla 7 — valores especiales de descripción

Las únicas normalizaciones automáticas de descripción son:

| Condición | Valor asignado |
|---|---|
| Descripción = `"NO FAULTS"` y anterior ≠ `"NO FAULTS"` | Descripción del intervalo anterior (`"F1 Watchdog"` si es vacía) |
| Descripción = `"NO FAULTS"` y anterior = `"NO FAULTS"` | `"F13 NO MODULES"` |
| Descripción = `""` (vacío) | `"F1 Watchdog"` |

Estas normalizaciones aplican **solo al motor de paridad de eventos** (`mcoCreateList`).

---

# 18. Arquitectura Python propuesta

No implementar todavía, pero diseñar el proyecto alrededor de módulos conceptuales:

```text
src/etl_arena/       -- layout src (ADR-01); detalle en design.md
  config/
  excel_semantics/  -- reglas de celda/texto/fechas/redondeo VBA-Excel
  model/
  acquisition/      -- Fase S: acquire-wait, contrato/lector SCADA
  workbook/         -- Fase M: sesión COM, preparar, macros, referencia (PC local)
  ingestion/
  normalization/
  enrichment/
  availability/
  fault_events/
  aggregation/
  persistence/
  reconciliation/
  reporting/
scripts/
  run_lunes.py      -- orquestador semanal (etapas S→T→M→E→R→B)
data/
  inbox/ processed/ work/
tests/
sql/
```

Separar explícitamente:

```text
acquisition != ingestion != business logic != persistence
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
nombre_proyecto         = "Arena BESS"
fecha_inicio_proyecto   = "2026-04-08"
TotalPCS            = 61
baterias_por_pcs    = 4
racks_por_pcs        = 12
minutos_muestreo     = 15
```

La fórmula equivalente a `Total_Racks` debe ser:

```text
TotalPCS * baterias_por_pcs * racks_por_pcs
```

---

# 20. Versionamiento del algoritmo

Cada resultado SQL debe guardar:

```text
version_algoritmo
```

Ejemplo:

```text
availability-v1-excel-parity
```

Cuando posteriormente se modifique una regla de negocio, crear otra versión.

Nunca sobrescribir resultados históricos con una nueva lógica.

---

# 21. Plan de implementación por etapas

## Etapa 0 — Adquisición SCADA semanal (Fase S)

**Estado:** diseñado; pendiente P0 (muestra real lunes).

Entregables:

- `data/inbox` + `scripts/run_lunes.py` (etapas `acquire-wait` … `notify-bi`);
- `src/etl_arena/acquisition` (fechas origen → serial Excel, validación de continuidad del export incremental, mapping columnas);
- `src/etl_arena/workbook` (runner COM de las 4 macros; solo PC local);
- runbook [`docs/runbook-lunes.md`](./docs/runbook-lunes.md);
- handoff de notificación a Power BI (Misael).

**Bloqueante:** P0 — muestra real del export SCADA (nombre archivo, header, delimiter, formato de fecha, encoding).

---

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
SCADA (solo extract) -> inbox -> adapter -> macros COM -> ETL Python -> SQL Server -> notificación/refresh PBI
```

El flujo operativo semanal queda descrito en §14 Fase S y `docs/runbook-lunes.md`. Excel de cálculo pasa a ser herramienta de consulta/legacy durante shadow mode; el `.xlsm` de trabajo sigue usándose para macros de referencia hasta retirarlo.

---

# 22. Criterios de aceptación

La migración no se considera terminada hasta que:

1. Una corrida Python pueda reproducir la corrida Excel usando exactamente el mismo input.
2. `BloquesMuestreo (C12)` sea reproducible.
3. `BloquesRacksIndisponibles (C14)` sea reproducible.
4. `DisponibilidadPeriodo (C16)` sea reproducible.
5. `DisponibilidadAnualAcumulada (C19)` sea reproducible.
6. Los eventos de `ListOfFaults` sean reconciliables.
7. La disponibilidad diaria sea reconciliable.
8. La disponibilidad acumulada anual sea reconciliable.
9. Los períodos que atraviesan cambio de horario sean validados.
10. Una corrida SQL pueda ser auditada hasta:
    `KPI -> acumulado -> muestra -> input raw`.
11. Todos los intervalos con `ModulosDisponiblesNulo = True` estén identificados y accesibles en la base de datos.

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

**Acción:** SQL debe ser append-only por `IdCorrida`.

### Riesgo G — `NUMBER_OF_MODULES` vacío silencioso

El Excel ignora silenciosamente los intervalos con `NUMBER_OF_MODULES` vacío. La frecuencia de este evento en el histórico completo es desconocida.

**Acción:** marcar con `ModulosDisponiblesNulo = True`; incluir en el reporte de calidad de datos de cada corrida.

### Riesgo H — descripción de falla con fallback al intervalo anterior

La macro `mcoCreateList` puede asignar a un evento la descripción del intervalo anterior cuando el intervalo de inicio reporta `"NO FAULTS"`. Esto se ha confirmado en datos de septiembre 2026 y puede ocurrir en cualquier período.

**Acción:** marcar con `descripcion_falla_fallback = True`; reportar la frecuencia de este caso en cada corrida.

### Riesgo K — Sin API SCADA; transporte manual (TeamViewer)

Hoy no hay API GPM ni share UNC al server. El archivo se copia a mano a `data/inbox`.

**Acción:** `acquire-wait` con alerta si no hay archivo; pedir share o API a GPM; RPA TeamViewer solo como fallback (P8), no camino crítico.

### Riesgo L — Formato de fechas del export SCADA

Origen probable: `mm-dd-aaaa hh:mm:ss`; destino hoja: `dd-mm-aaaa hh:mm:ss`.

**Acción:** transformador estricto en `scada_adapter` (Fase T) + validación de continuidad del export incremental con el libro base (D-07).

### Riesgo M — PlantActivity fuera de la cadena SCADA

SCADA solo actualiza `RawData-PCS`.

**Acción:** mantener fuente aparte; `prepare-workbook` valida la alineación por fila con `RawData-PCS` (F-01, F-21) y el reporte de calidad informa desalineaciones y filas sin timestamp (F-31).

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

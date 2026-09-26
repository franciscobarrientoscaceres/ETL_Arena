# Data contract — libro de cálculo de disponibilidad

**Libro:** `data/AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm` (versión del panel: `Calculation-Availability!A33 = "Version: 01-09-2026"`).
**Perfilado:** 2026-09-24, lectura del XML crudo (valores cacheados, sin conversión de fechas; mismo lector que `tests/golden/extract_golden.py`).
**Alcance:** lo que el pipeline Python (`etl_arena.ingestion`) puede asumir del libro y qué debe rechazar o reportar. Las reglas de cálculo están en `requirements.md` y `design.md`; los hallazgos `F-xx`, en `.kiro/specs/etl-arena-availability/audit.md`.

Convenciones: fila y columna en notación Excel (fila 1 = encabezado). "Vacío" significa que la celda no existe en el XML o no tiene `<v>`. Los seriales son días desde 1899-12-30 (`float64`).

---

## 1. Hojas

| Hoja | Parte XML | Uso en el pipeline |
|---|---|---|
| `RawData-PCS` | `sheet1.xml` (~126 MB) | **Entrada principal** (§2) |
| `Calculation-Availability` | `sheet2.xml` | Parámetros KPI (§4) y tabla de resultados (referencia, nivel 2) |
| `ListOfFaults` | `sheet3.xml` | Parámetros de eventos (§4) y lista de eventos (referencia, nivel 5) |
| `Graph` | `sheet4.xml` | Presentación; se ignora |
| `Daily` | `sheet5.xml` | Parámetro `D5` (§4) y serie diaria (referencia, nivel 4) |
| `Annual_AVA` | `sheet6.xml` | KPI mensual mantenido a mano (§6) |
| `PCS-Fault` | `sheet7.xml` | Catálogo de fallas → seed `tipo_detencion` (§5) |
| `PCS-Status` | `sheet8.xml` | Catálogo de estados (§5) |
| `DateFormat_Correction` | `sheet9.xml` | Área de trabajo manual del pegado SCADA (§7); se ignora |
| `PlantActivity` | `sheet10.xml` | **Entrada**: factor operacional (col C, §3); D solo espejo de BO (F-37) |
| `Exclusion_Matrix` | (desde el libro de agosto) | **Entrada**: eventos de exclusión por fila y PCS (§3b, F-37) |
| `Verificación` | `sheet11.xml` | Rota (`#VALUE!`); se ignora (F-24) |

El pipeline localiza las hojas **por nombre** vía `workbook.xml` + rels, nunca por el número de la parte XML.

---

## 2. `RawData-PCS`

### 2.1 Estructura

| Columna | Encabezado (fila 1) | Contenido |
|---|---|---|
| A | `Date/time` | Serial Excel de la muestra |
| 4·(p−1)+2 | `Arena - PCS pp - POWERELECTRONICS GEN3 HEx CURRENT FAULT` | Falla |
| 4·(p−1)+3 | `Arena - PCS pp - POWERELECTRONICS GEN3 HEx CURRENT STATUS` | Estado |
| 4·(p−1)+4 | `Arena - PCS pp - POWERELECTRONICS GEN3 HEx CURRENT WARNING` | Advertencia |
| 4·(p−1)+5 | `Arena - PCS pp - POWERELECTRONICS HEM-k NUMBER OF MODULES` | `ModulosDisponibles` |

- `p` = 1…`total_pcs`, `pp` con dos dígitos. PCS 1 → B:E (módulos en **E**, índice 5 = 4·p+1).
- Observado: 245 columnas = 1 + 4 × 61; los 244 encabezados coinciden exactamente con el patrón.
- **Validación:** el número de PCS deducido de los encabezados debe ser igual a `C2` (F-17); un encabezado que no calce con el patrón rechaza el archivo.

### 2.2 Columna A (tiempo)

| Propiedad | Observado | Contrato |
|---|---|---|
| Tipo | 15.990 seriales numéricos (166 enteros = medianoche) | Numérico. Texto en A **rechaza** el archivo (F-21) |
| Rango | 2026-04-08 00:15 (fila 2) → 2026-09-21 14:15 (fila 15991) | — |
| Paso | 15 min; un único salto de 60 min: 2026-09-06 00:00 → 01:00 (DST) | Pasos ≠ `minutos_muestreo` se **reportan**, no se corrigen (F-32, ADR-07) |
| Orden | Estrictamente creciente, sin duplicados | Filas fuera de orden o duplicadas se reportan y **no** se reordenan ni deduplican (F-07) |
| Fin | Primera A vacía en la fila 15992 | El VBA termina en la **primera A vacía**; filas posteriores se truncan y reportan (`modo_huecos = "excel"`, D-01, F-10) |

**Convención de medianoche:** la fila `dd 00:00` lleva la fecha del día que empieza. El filtro KPI es `C5 ≤ A < C7 + 1`, así que el período 01-sep…20-sep incluye 01-sep 00:00 y excluye 21-sep 00:00. El primer dato del proyecto es 08-abr 00:15 (no hay fila 00:00 ese día).

**Seriales, no fechas:** filtros y restas usan el `float64` del `<v>` crudo (ADR-03). Ejemplo: una duración de 5,5 h en Excel vale `5.5000000001164153`.

### 2.3 Tipos por campo (61 PCS × 15.990 filas = 975.390 celdas por campo)

| Campo | Texto | Número entero | Número fraccionario | Vacío |
|---|---:|---:|---:|---:|
| FAULT | 973.119 | 1.705 | 0 | 566 |
| STATUS | 974.824 | 0 | 0 | 566 |
| WARNING | 875.550 | 99.274 | 0 | 566 |
| NUMBER OF MODULES | 0 | 959.169 | 15.655 | 566 |

**NUMBER OF MODULES** (gobierna la indisponibilidad):
- Rango observado [0, 4]; fraccionarios como 2,3333 / 2,6667 / 3,2347 / 3,6667 (F-02).
- Se conserva como `float` sin truncar (Python `float`, SQL `FLOAT`).
- Vacío → el PCS cuenta como disponible y se marca `ModulosDisponiblesNulo = True`.
- Texto → *Type mismatch* en el VBA: el archivo se **rechaza** en modo paridad (F-16). Valores < 0 o > 4 se rechazan (no observados).

**FAULT** (descripción de falla para `mcoCreateList`):
- Texto `"F<n> <DESCRIPCIÓN>"` en mayúsculas (`"F55 EXTERNAL FAULT/OVGR"`, `"F1 WATCHDOG"`) o `"NO FAULTS"` (936.739 celdas).
- 1.705 celdas son el **número** de la falla sin texto (`169`, `228`, `170`…): Excel produce el código `"F" & n` (F-14).
- Un texto con espacio inicial da código `""` (`FIND` devuelve 1). No hay casos hoy, pero el contrato lo cubre.
- El fallback del VBA escribe `"F1 Watchdog"` y `"F13 NO MODULES"` con capitalización propia, distinta de la del SCADA (F-04).

**STATUS / WARNING:** no participan en el cálculo; se persisten como texto (un número se guarda con `texto_excel`).


---

## 3. `PlantActivity`

### 3.1 Estructura

| Columna | Encabezado (fila 1) | Uso |
|---|---|---|
| A | (vacía) | — |
| B | `Date/Time` | Serial; **solo para validar la alineación** |
| C | `Activo 1\n Inactivo 0` | `FactorOperacional` (se aplica si `C21 <> "No"`) |
| D | `Excused Event` | **Ya no pondera** (F-37): se conserva como espejo de `Calc!BO`. La exclusión viene de `Exclusion_Matrix` (§3b) |
| E:I | Señales PPC/POI/SOC (`… SETPOINT (kW)`, droop, `ACTIVE POWER (kW)`, `Total Banks State of Charge (%)`) | Informativas; `PorcentajeSOC` desde la columna I |

### 3.2 Unión con `RawData-PCS` (F-01, D-02)

- La fila `r` de `PlantActivity` se asocia a la fila `r` de `RawData-PCS` (`Sheet4.Cells(dblRec, 3|4)`), **no** por timestamp.
- Celda C o D vacía → **0** (valor Excel de una celda vacía).
- Si B tiene timestamp y difiere del de `RawData-PCS!A` en la misma fila (> 1e-6 días) → se reporta desalineación; no se corrige.

### 3.3 Estado observado

| Tramo de filas | B (timestamp) | C | D |
|---|---|---|---|
| 2 … 15332 | Presente; **0 desalineaciones** con `RawData-PCS` | 1 / 0 | 1 / 0 |
| 15333 … 15991 (desde 2026-09-14 17:45) | **Vacío** (659 filas) | 0 en todas | 1 en todas |
| 15992 … 26617 | Vacío | 0 (prellenado) | 1 (prellenado) |

En el rango de datos, C tiene 10.504 unos y 5.486 ceros; D tiene 15.437 unos y 553 ceros; ninguna C/D está vacía.

**Consecuencia (F-31):** `PlantActivity` está atrasada respecto a `RawData-PCS`. Con `C21 = "Yes"`, las 659 filas sin timestamp cuentan como **inactivas** (factor 0) y su indisponibilidad desaparece del KPI. El reporte de calidad informa: filas sin timestamp en el rango de cálculo, primera fila afectada y si `C21 = "Yes"`.

---

## 3b. `Exclusion_Matrix` (F-37, desde el libro de agosto 2026)

| Columna | Encabezado (fila 1) | Contenido |
|---|---|---|
| A | `Date/time` | Serial; solo para validar la alineación con `RawData-PCS!A` |
| k+1 (B…BJ) | `PCS01` … `PCS61` | Valor de exclusión del PCS k: `0`/vacío, `1` o `2` |
| p+2 (BK) | `Excused Event` | Resumen por fila (1/2/0); informativo, el cálculo no lo usa |
| p+3 (BL) | `Comments` | Causa del evento (`CPF`, `External`, …) |

- **Unión por fila** con `RawData-PCS` (igual que PlantActivity). Observado en agosto: 15.990 filas, 0 desalineadas.
- **Valores:** `0`/vacío = sin evento de exclusión; `1` = se consideran todos los módulos (0 baterías indisponibles); `2` = se consideran los módulos previos al evento (baterías previas del tramo contiguo). Cualquier otro valor o un encabezado distinto **rechaza** la corrida.
- **Solo pondera** con `C31 = "Yes"` (KPI) / `L14 = "Yes"` (eventos) y cuando la celda de módulos está en falla (`M < 4`).
- **Libro de agosto:** 4.982 celdas con 1 y 1.498 con 2, todas en agosto salvo PCS61 = 1 en la fila 2 (F-39); los 6 PCS con valor 2 tenían 3 módulos antes del evento.
- **Sin la hoja** (libros hasta septiembre): todo 0 y anomalía informativa `exclusion_matrix_ausente`.

---

## 4. Celdas de parámetros

Todas son **valores**, salvo `C23`. El runner COM las escribe explícitamente antes de las macros (F-23); el pipeline las lee para `ConfiguracionCalculo`.

| Celda | Etiqueta en el libro | Observado | Semántica |
|---|---|---|---|
| `Calculation-Availability!C2` | Total PCS | 61 | `total_pcs` |
| `C3` | Total Batteries per PCS | 4 | `baterias_por_pcs` |
| `C5` | Fecha inicio | 46266 (2026-09-01) | Inicio KPI (inclusive) |
| `C7` | Fecha fin | 46286 (2026-09-21) | Fin KPI (día inclusive: `A < C7 + 1`) |
| `C11` | Total Racks | 2928 | `total_racks` |
| `C21` | Only Operational Time? | `"No"` | Aplica factor operacional si `<> "No"` (exacto, sensible a mayúsculas; F-12) |
| `C23` | Frecuencia de muestreo [min] | 15 | **Fórmula** `ROUND((E5-E4)*24*60, 2)`: se deriva de las dos primeras filas de resultado; vale 60 si el período empieza en el salto DST (F-13) |
| `C31` | Excusable event? | `"No"` | Aplica `Exclusion_Matrix` solo si `= "Yes"` (F-37). Oficial: `"Yes"` (D-03) |
| `ListOfFaults!L2` | Fecha inicio | 46266 | Inicio eventos (independiente de C5; F-05) |
| `ListOfFaults!L4` | Fecha fin | 46286 | Fin eventos (independiente de C7) |
| `ListOfFaults!L14` | Excusable event? | `"Yes"` | Aplica `Exclusion_Matrix` en eventos (independiente de C31; D-16). Oficial: `"Yes"`, igual que C31 (D-03) |
| `Daily!D3` | Fecha inicio | 46266 | `= C5` |
| `Daily!D5` | Fecha fin | 46286 | Informativo: la macro corta en la primera `Daily!C` vacía (C9 = D3, C10:C29 = anterior + 1, extendidas a mano; máx. 31 días) — F-09, F-33 |

Salidas de referencia (las escriben las macros): `C12` (1975), `C14`, `C16`, `C19`, `ListOfFaults!L10`, tabla `Calculation-Availability!E4:BO` (E = serial; F…BN = `PCS01…PCS61`; BO = factor excusable con encabezado `"1=No, 0=Yes"`), `ListOfFaults!B6:I` y `Daily!B9:G39`. Las celdas `C25`/`C27`/`C29` (tiempos de la macro) no son parámetros.

Invariante de conteo observado: `C12 = 1975` = filas en [01-sep 00:00, 21-sep 14:15] = 20 × 96 + 58 − 3 (salto DST).

---

## 5. Catálogos

### `PCS-Fault` → `tipo_detencion`
- Encabezado en la **fila 2**, columnas B:G: `Code Number`, `Code Fault`, `Description PE`, `Code + Description`, `Meaning`, `Operative`.
- 167 filas de datos (3…169): de `0 / F0 / NO FAULT / "F0 NO FAULT" / No fault / Yes` a `257 / F257 / Timeout carga suave AC`.
- **163 códigos distintos:** F228, F230, F231 y F232 están dos veces, idénticos salvo `Meaning` (`Crítico` / `Parcial`) — F-35, D-14.
- `ListOfFaults!N6:N171` (resumen N:Q) tiene 166 filas con esos mismos 163 códigos (F230–F232 repetidos).
- El seed se genera desde la hoja (`scripts/generar_seed_tipo_detencion.py`), nunca hardcodeado (F-19).

### `PCS-Status`
- Encabezado en la fila 2, columnas B:E: `Code`, `Description PE`, `Meaning`, `Operative`.
- 18 estados (0 `POWER UP` … 17 `FAULT`); solo `FAULT` tiene `Operative = No`.

---

## 6. `Annual_AVA` (mantenida a mano; F-20)

- Filas 7…18 = meses 1…12 (B = número, C = nombre); encabezados en la fila 6. Columnas: D días/mes, E bloques, F rack-bloques indisponibles, G disponibilidad mensual, H/I acumulados, J disponibilidad anual acumulada, K contractual (0,98).
- Julio y agosto son valores tipeados; septiembre `F15` enlaza `C14`, y `D15 = 20.59375` está tipeado (→ `E15 = 1977`, calendario, ≠ `C12 = 1975`).
- La acumulación arranca en **julio** (abril–junio vacíos). No es fuente de paridad del período; se importa como histórico a `disponibilidad_mensual` con `Origen = excel_manual` (D-06 pendiente).

---

## 7. Evidencia sobre el formato del export SCADA

`DateFormat_Correction` guarda el último bloque pegado a mano (682 filas desde 2026-09-14 12:15):

| A `Date` | B `Time` | D:H (fórmulas) |
|---|---|---|
| `9/14/2026` (texto) | `12:15 PM` (texto) | `FIND("/")`, `MID(…)` → día, mes, año |

Es decir, el export llega como **`m/d/yyyy` sin ceros a la izquierda, con la hora en columna aparte en formato 12 h AM/PM**, y no como `mm-dd-aaaa hh:mm:ss` (que es lo que asumen hoy README/AGENTS §14). Queda como insumo para el contrato SCADA (tareas 0.6/0.7), que debe confirmarse con un export real.

---

## 8. Resumen: rechazar vs reportar

| Condición | Acción (modo paridad) |
|---|---|
| Encabezados de `RawData-PCS` fuera de patrón o PCS ≠ `C2` | Rechazar |
| Texto en `RawData-PCS!A` o en NUMBER OF MODULES; módulos < 0 o > 4 | Rechazar |
| Primera A vacía antes del fin de los datos | Truncar como Excel + reportar |
| Paso ≠ `minutos_muestreo`, duplicados, desorden | Reportar (sin corregir) |
| NUMBER OF MODULES vacío | Disponible + `ModulosDisponiblesNulo` |
| FAULT numérico | Código `"F" & n` (sin reporte) |
| PlantActivity sin timestamp o desalineada | Reportar (C/D vacíos = 0) |
| `C23 ≠ minutos_muestreo` | Reportar |
| Evento que inicia en la fila 2 | Abrir evento + `ExcelHabriaFallado` (D-08) |

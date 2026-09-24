# Auditoría SPEC — ETL Arena Availability

**Fecha:** 2026-09-24 · **Alcance:** `AGENTS.md`, `CLAUDE.md`, `README.md`, `docs/runbook-lunes.md`, `.kiro/specs/etl-arena-availability/*`, `tests/golden/*`, `.opencode/agent/*`, y el libro `data/AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm` (VBA extraído con `olevba`, fórmulas leídas del XML, datos perfilados con `openpyxl`).

**Método:** cada regla del spec se contrastó contra el **código VBA ejecutable** y las **fórmulas/valores reales** del libro. Cuando el spec y el VBA difieren, manda el VBA (regla de paridad). Los hallazgos ya están incorporados en `requirements.md`, `design.md` y `tasks.md` de este mismo directorio.

Severidad: 🔴 rompe paridad o bloquea implementación · 🟠 alto (error funcional/operacional) · 🟡 medio (consistencia, calidad, entorno).

---

## 1. Resumen ejecutivo

| # | Estado del repositorio | |
|---|---|---|
| Código Python | No existe (`src/`, `sql/`, `scripts/` son objetivo) | — |
| Tests | `tests/golden/test_golden_integrity.py` + 3 goldens JSON (todos `pending`) | — |
| Spec Kiro | requirements/design/tasks completos pero con **17 errores de paridad o de hechos** | 🔴 |
| Entorno | Solo Python 3.14; sin pandas/pytest; ODBC solo driver legacy `SQL Server`; sin SQL Server local; Docker y Excel 16 disponibles | 🟡 |
| Agentes | 6 agentes OpenCode genéricos (`.opencode/agent/`), no cargados por Claude Code; ninguno cubre VBA/COM ni QA de paridad | 🟡 |

**Conclusión:** el plan general (paridad → SQL → reconciliación → Fase S → shadow) es correcto, pero el pseudocódigo del design **no habría alcanzado paridad** ni siquiera con el golden de septiembre. Tras las correcciones, el spec queda listo para implementar por olas (ver `tasks.md`).

---

## 2. Hallazgos de paridad (VBA real vs spec anterior)

### F-01 🔴 PlantActivity se une **por número de fila**, no por timestamp
- **Evidencia VBA:** `cmdAvailability.Cells(dblResult, 67) = Sheet4.Cells(dblRec, 4)` y `Sheet4.Cells(dblRec, 3)`; `mcoCreateList` usa `Sheet4.Cells(dblRec, 4)`. `dblRec` es la fila de `RawData-PCS`.
- **Evidencia datos:** filas 2..15332 alineadas; desde la fila 15333 (2026-09-14 17:45) PlantActivity **no tiene timestamp** (659 filas) pero sí valores C/D.
- **Spec anterior:** join por `MarcaTiempoMuestra`; si falta → factores = 1 (Req 3.3). En Excel una celda vacía vale **0**.
- **Corrección:** join por `NumeroFilaOrigen`; celda vacía → 0; validar que el timestamp de PlantActivity (si existe) coincide con el de RawData y reportar desalineación.

### F-02 🔴 `NUMBER_OF_MODULES` es **fraccionario**
- **Evidencia:** 15.655 celdas con valor no entero (p. ej. `3.2346…`, `3.6667`); 959.169 enteros; 566 vacías; ningún texto ni valor > 4.
- **Spec anterior:** `modulos_disponibles = int(valor_raw)`; DDL `ModulosDisponibles INT`, `BateriasIndisponibles INT` → trunca y rompe C14.
- **Corrección:** `float` en Python y `FLOAT` en SQL; conservar valor crudo.

### F-03 🔴 El fin de evento es la **última fila en falla**, no la siguiente
- **Evidencia VBA:** `Sheet3.Cells(dblResult, 4) = Sheet2.Cells(dblRec, 1)` dentro del bloque en que la fila actual está en falla y la siguiente la cierra. Golden GT-6: inicio 00:15, fin 05:45, duración 5,5 h.
- **Spec anterior:** `marca_fin = ts` de la fila no fallada → +15 min en cada evento.

### F-04 🔴 Fallback de descripción con anterior **vacío** produce `F1 Watchdog`, no `F13`
- **Evidencia VBA:** `If Sheet2.Cells(dblRec-1, intColRec) <> "NO FAULTS" Then G = anterior` → si el anterior es `""`, `G=""` y luego `If G = "" Then G = "F1 Watchdog"`.
- **Spec anterior (design):** `if descripcion_anterior not in ("NO FAULTS", "")` → asigna F13. Además la fila anterior se toma de la hoja completa (incluye filas previas al período y la fila 1 de encabezado), no del subconjunto filtrado.

### F-05 🔴 `mcoCreateList` tiene **parámetros propios**: `ListOfFaults!L2`, `L4`, `L14`
- **Evidencia XML:** L2=46266, L4=46286, L14=`"Yes"` son **valores**, no fórmulas. En el libro actual `C31="No"` pero `L14="Yes"`.
- **Spec anterior:** un único `aplicar_evento_excusable` y el período de `C5/C7` para ambos motores; el runner COM solo setea C5/C7.
- **Corrección:** `ConfiguracionCalculo` separa `inicio/fin_periodo_eventos` y `aplicar_evento_excusable_eventos`; el runner escribe L2/L4/L14; el golden los registra. Decisión de negocio pendiente sobre qué valor es el intencional (D-03).

### F-06 🔴 Bug VBA: evento abierto al cierre del período **se arrastra al siguiente PCS**
- **Evidencia VBA:** la condición de cierre es `siguiente.ts > L4 + 1` (estricta). Si el período termina a las 23:45 y la fila siguiente es `L4+1` 00:00 exacto y el PCS sigue en falla, el evento **no se cierra**; `sumablocks`, `numBlock` y `dblResult` no se reinician al cambiar de PCS → el siguiente PCS sobrescribe PCS/inicio/descripción de esa fila y hereda los bloques acumulados.
- **Spec anterior:** inventa un cierre en `fin_periodo + 1 día`.
- **Impacto observado:** en jul/ago `ListOfFaults!L10 × 4 ≠ C14` (70.129×4 vs 432.611; 63.791×4 vs 527.396), mientras que en septiembre (datos terminan el 21 a las 14:15, sin fila 00:00 posterior) `C14 − 4 × L10 = −1,3e-7` (ruido de seriales, dentro de 1e-6).
- **Corrección:** emular la escritura de celdas (registro mutable `dblResult`) y marcar `EventoArrastradoExcel`. Corregirlo es una versión v2 (D-04).

### F-07 🔴 `BloquesMuestreo` cuenta **filas**, no timestamps únicos
- **Evidencia VBA:** `C12 += 1` por cada fila en rango.
- **Spec anterior:** itera `df["MarcaTiempoMuestra"].unique()` → fusiona duplicados. El cambio de hora de abril 2027 (Chile) generará 4 timestamps locales repetidos que Excel cuenta dos veces.

### F-08 🔴 `mcoDailyAvailability` **ignora el factor operacional**
- **Evidencia VBA:** `Suma += 12 * cmdAvailability.Cells(k, 5 + j)` — la tabla de resultados guarda baterías ponderadas por excusable, **sin** `PlantActivity.C`. El factor operacional solo entra en C14.
- **Consecuencia:** con `C21="Yes"`, `Σ daily ≠ C14` en Excel. El spec anterior sumaba `ImpactoRackPonderado` y el invariante del extractor exige `Σ daily = C14` siempre.

### F-09 🟠 El "día N" del Daily es relativo a `C5`, y el rango de días es un input propio (`Daily!D5`)
- **Evidencia XML:** `Daily!C9 = D3 (= C5)`, `C10 = C9+1`…, `B9 = 1`; `F9 = 1 - E9/(Total_Racks*24*60*B9/Frecuencia)`; `Daily!D5` es valor (fin), máx. 31 filas (C9:C39).
- **Evidencia golden:** septiembre tiene 30 días (D5 era 30-sep al extraer); el libro hoy muestra solo hasta el 21.

### F-10 🟠 El VBA **termina en la primera celda vacía** de `RawData-PCS!A`
- Req 1.6 pedía continuar → divergencia si el archivo trae filas en blanco intermedias. Se define `modo_huecos = "excel"` (trunca + reporta) para paridad.

### F-11 🟠 Evento que inicia en la **primera fila de datos** hace fallar la macro
- La fila anterior es el encabezado: `"Arena - PCS 04 - …" = 4` → *Type mismatch* en VBA. El PCS 04 tiene 3,2347 en la fila 2 (2026-04-08 00:15), así que `mcoCreateList` con período de abril revienta. Python debe definir comportamiento explícito (abrir evento + marcar `ExcelHabriaFallado`).

### F-12 🟠 Semántica asimétrica de flags
- `C21`: el factor operacional se aplica si `C21 <> "No"` (comparación exacta, sensible a mayúsculas). `C31`/`L14`: solo si `= "Yes"`. La config debe leerse con esa semántica (helper `flag_excel`).

### F-13 🟠 La frecuencia `C23` se **deriva** de los datos
- `C23 = ROUND((E5-E4)*24*60, 2)` (dos primeras filas del resultado de `cmdCalcAvailability`) y `mcoCreateList` la usa para el inicio del evento. Si el período empieza en el salto DST (2026-09-06 00:00→01:00) C23 = 60. Python: derivar igual, comparar contra `minutos_muestreo` y reportar.

### F-14 🟡 Código de falla con descripciones numéricas o con espacio inicial
- 1.705 celdas de `CURRENT FAULT` son enteros (sin texto) → Excel produce `"F" & n`. Con espacio inicial `FIND` devuelve 1 y `MID(…,1,0) = ""` (no entra al `IFERROR`). El `idx > 0` del design difería.

### F-15 🟡 Aritmética en **seriales Excel**
- `ListOfFaults!E6 = 5.5000000001164153`, `I6 = 219.00000000463547`: el VBA resta seriales `double`. Con `timedelta` Python obtiene 5,5 exacto → diferencia 4,6e-9 en rack-hours (> 1e-9). Los motores de paridad operan sobre `SerialFechaExcelOrigen` leído del `<v>` crudo del XML.

### F-16 🟡 Módulos con texto → *Type mismatch* en VBA
- No hay casos hoy. Contrato: rechazar el archivo en modo paridad.

### F-17 🟡 `mcoCreateList` recorre PCS mientras haya encabezado en la fila 1; `cmdCalcAvailability` recorre `C2`. Validar que coinciden.

---

## 3. Hallazgos de documentación y diseño

### F-18 🟠 El "límite de 167 eventos" no existe
- `mcoOrder` ordena `N5:Q172`: tabla resumen por **código de falla** (`P = SUMIF(F, código, I)`, 167 códigos). La lista de eventos `B:I` no se ordena ni tiene tope (limpia hasta la fila 150000). AGENTS §8–9 y Req 8.8 lo interpretaban mal.

### F-19 🟠 El catálogo `PCS-Fault` tiene **167 códigos**, no 68
- Filas F0…F257 con columnas `Meaning` (criticidad) y `Operative`. El INSERT de 68 filas hardcodeado dejaría `IdTipoDetencion = NULL` para F252, F255, etc. El seed se genera desde la hoja.

### F-20 🟠 `Annual_AVA` es una hoja **mantenida a mano**
- Jul/ago son valores tipeados; sep `F15 = C14`; `D15 = 20+14.25/24` (días transcurridos tipeados); la acumulación empieza en **julio** (H12 = 0), no el 08-abr; `J = 1 - I/(Racks·H)` es distinta de `C19 = 1 - C14/(Racks·35040)`. Req 6.6–6.8 mezclaba ambas. Se modela con una tabla de KPI mensual oficial (incluye filas históricas importadas del libro) y se deja la fecha de inicio de acumulación como decisión (D-06).

### F-21 🟠 El adapter no puede escribir fechas como **texto** `dd-mm-aaaa`
- Texto vs fecha en VBA: el texto es "mayor" que cualquier número → el filtro `< C7+1` descarta la fila y C12 cae a 0. Hay que escribir seriales/fecha real; `dd-mm-aaaa` es solo formato de visualización.
- Además, reescribir `RawData-PCS` con un export que no empiece exactamente en la misma fila rompe la alineación con PlantActivity (F-01). `prepare-workbook` debe validarla.

### F-22 🟠 `openpyxl` sobre `.xlsm` pierde objetos
- `keep_vba=True` conserva el VBA pero descarta dibujos/controles/gráficos (hoja `Graph`). Toda escritura en el libro de trabajo se hace por **COM**.

### F-23 🟠 El runner COM debe setear y extraer más que C5/C7 → C12/C14/C16/C19
- Setear: C5, C7, C21, C31, L2, L4, L14, Daily!D5. Extraer: tabla `Calculation-Availability!E4:BO` (nivel 2), `ListOfFaults!B:I` completa + L10 (nivel 5), `Daily` completo (nivel 4). Las macros usan `.Select`/`ActiveWindow` → Excel visible y carpeta de trabajo como *Trusted Location*.

### F-24 🟡 Hoja `Verificación` rota (`#VALUE!`)
- No es fuente de paridad. Su intención sí se adopta como invariante: por PCS, `Σ(12·baterías)/4 = Σ rack-hours de eventos` y global `C14 = 4·L10` cuando flags y período coinciden y no hay arrastre.

### F-25 🟡 Goldens incompletos
- Solo 20 eventos por mes; jul/ago incoherentes entre eventos y C14 (F-05/F-06/estado viejo); sep con 30 días de Daily vs 21 en el libro; GT-3 (junio) no está en el índice; `extract_golden` asume `Σ daily = C14` (solo válido con `C21="No"`).

### F-26 🟡 Inconsistencias internas del spec
- Tolerancias 1e-9 / 1e-6 / 1e-4 en lugares distintos · "10 tablas" vs 11 · seeds duplicados en `01_create_tables.sql` y `04/05_seed` (INSERT no idempotente) · requisitos 14/15 entre 1 y 2 · `IdCorrida` duplicado en glosario · AGENTS §4.1 dice `HEx STATUS` (real: `CURRENT STATUS`) · mezcla snake_case/PascalCase · `detencion.CodigoFalla NVARCHAR(10)` vs `"F" & descripción` larga · `CodigoFallaRaw` y `DescripcionFallaRaw` para una sola columna origen.

### F-27 🟠 `detencion` se duplica cada lunes
- Cada corrida semanal recrea todas las detenciones del período → duplicados y pérdida de `Observacion`/`EstadoRevision`. Se separa el workflow en `detencion_revision` (clave de negocio `IdProyecto+NumeroPCS+FechaInicio`) y se expone `v_detencion_vigente`.

### F-28 🟠 Volumen y definición del período semanal
- El export acumulado (01-01-2026 → domingo) da ~16k filas × 61 PCS ≈ 1M filas por tabla por corrida, creciendo cada semana (~50–100M filas/año). No estaba definido qué período calcula cada lunes. Se proponen columnstore + persistir muestras solo del período (D-07, D-09).

### F-29 🟡 Entorno local
- Python 3.14 únicamente (suficiente: todas las dependencias tienen wheels; `.venv` creada en la ola 0). ODBC: solo el driver legacy `SQL Server` (no maneja bien `DATETIME2` ni `fast_executemany`) → instalar **ODBC Driver 18**. Sin SQL Server local → contenedor `mcr.microsoft.com/mssql/server:2022` (Docker disponible). Excel 16 presente (COM viable).

### F-30 🟡 Agentes
- 6 agentes OpenCode (`mode: subagent`) genéricos; `database-optimizer` está orientado a PostgreSQL (el destino es SQL Server). Ninguno especializado en VBA/COM ni en QA de paridad; Claude Code no los carga (no están en `.claude/agents/`). Se asignan en `tasks.md` con instrucciones de contexto obligatorias.

### F-31 🟡 PlantActivity desactualizada
- Sin timestamps desde 2026-09-14 17:45; con `C21="Yes"` esos bloques contarían como inactivos. Reportar en calidad y en el runbook.

### F-32 🟡 Zona horaria
- Los datos son hora local **naive** con salto DST (2026-09-06 00:00 → 01:00, 3 filas faltantes → C12 = 1975). `pytz.localize(is_dst=None)` fallaría en el retroceso de abril. No localizar en los motores; guardar naive local + serial.

---

## 4. Decisiones abiertas (con valor por defecto de paridad)

| ID | Decisión | Default v1 (paridad) | Quién decide |
|---|---|---|---|
| D-01 | Filas tras la primera celda A vacía | Truncar como Excel y reportar | Técnico |
| D-02 | Join PlantActivity | Por fila; vacío = 0; reportar desalineación | Técnico |
| D-03 | `L14` (eventos) vs `C31` (KPI): ¿mismo valor? | Parámetros separados; el runner escribe ambos explícitamente | **Negocio** |
| D-04 | Bug de arrastre de eventos al cierre | Replicar + marcar `EventoArrastradoExcel`; corregir en v2 | **Negocio** |
| D-05 | Daily sin factor operacional | Replicar | Técnico |
| D-06 | Inicio de acumulación anual | Meses con KPI oficial (hoy jul-2026 en el libro); backfill abr–jun opcional | **Negocio** |
| D-07 | Período que calcula cada lunes | Mes en curso hasta el último domingo; primer lunes del mes cierra también el mes anterior | **Negocio** |
| D-08 | Evento que inicia en la fila 2 | Abrir evento + marcar `ExcelHabriaFallado` | Técnico |
| D-09 | Retención de muestras por corrida | Raw completo del export (trazabilidad) con columnstore; `availability_sample_result` solo del período | Técnico/DBA |
| D-10 | Versión de Python | **Resuelta 2026-09-24: 3.14** (hay wheels de todas las dependencias: pyodbc 5.3, pywin32 312, pandas 3.0, numpy 2.5); `requires-python >= 3.13` | Técnico |

---

## 5. Invariantes nuevos verificados contra el libro

| Invariante | Septiembre (libro actual) |
|---|---|
| `C16 = 1 - C14/(C11·C12)` | ✔ 0,98199240990523617 |
| `C19 = 1 - C14/(C11·35040)` | ✔ 0,99898501739619983 |
| `C14 = 4·L10` (flags iguales en la práctica, sin arrastre) | ✔ Δ = −1,3e-7 (≤ 1e-6) |
| `C12 = filas en rango` (DST: 1975 = 1978 − 3) | ✔ |
| `Annual_AVA!E15 = D15·96` con `D15 = último_ts − inicio_mes` | ✔ 20,59375 × 96 = 1977 |
| `Σ Daily = C14` (solo si `C21="No"`) | ✔ |

---

## 6. Hallazgos de la implementación de la Fase 1 (2026-09-24)

Detectados al emular el VBA línea por línea y comparar contra el golden de septiembre (paridad bit a bit alcanzada en C12, C14, C16, C19, tabla `E4:BO`, Daily, 334 eventos, L10 y N:Q).

### F-33 🟠 La serie `Daily` la gobierna `Daily!C`, no `D5`
- **Evidencia XML:** `C9 = D3`, `C10:C29 = C(n-1)+1` y `C30:C39` **vacías**; `mcoDailyAvailability` termina en la primera `Daily!C` vacía. `D5` es un valor que no participa en la macro.
- **Consecuencia:** extender la serie requiere arrastrar a mano las fórmulas de `C`. En septiembre coinciden (`D5` = C29 = 21-sep) por casualidad.
- **Corrección:** `config.desde_excel` toma `fin_diario` de la última fecha contigua de `Daily!C9:C39` y anota si `D5` difiere; el runner COM (4.2) debe escribir `Daily!C9:C(8+n)` y limpiar el resto.

### F-34 🔴 `ListOfFaults!C/D` se redondean al segundo (Date de VBA → celda)
- **Evidencia:** `Sheet2.Cells(r, 1)` es un `Date` (celda con formato fecha), así que `A − C23/(24*60)` también lo es; Excel lo guarda con resolución de segundo. El libro guarda `C = 46266.010416666664` (serial exacto de 00:15:00), no `46266.01041666667` (resta en `double`); `E = 24*(D−C) = 5.5000000001164153` solo cuadra con el primero.
- **Impacto sin corregir:** 109 de 334 eventos con Δ de 1 ulp en C, que se propaga a E, I y L10 (4,8e-7).
- **Corrección:** `excel_semantics.fecha_vba_a_celda` (serial → datetime al segundo → serial correctamente redondeado con `Fraction`).

### F-35 🟡 `PCS-Fault` tiene códigos repetidos; N:Q tiene 166 filas
- **Evidencia:** 167 filas pero **163 códigos distintos**: F228, F230, F231 y F232 aparecen dos veces, idénticos salvo `Meaning` (`Crítico` / `Parcial`). `ListOfFaults!N6:N171` tiene 166 filas (F230–F232 dos veces) con los mismos 163 códigos.
- **Consecuencia:** el seed de `tipo_detencion` (clave `CodigoFalla`) no puede insertar las 167 filas tal cual; qué criticidad vale para esos 4 códigos es D-14. El resumen N:Q se compara como mapa código → P.
- **Nota:** `SUMIF` compara sin distinguir mayúsculas y `mcoOrder` es un sort estable: el orden de los empates depende del estado previo de la hoja.

### F-36 🟡 `numBlock As Integer` (latente)
- Un evento de más de 32.767 bloques seguidos (~341 días a 15 min) hace fallar `mcoCreateList` con *Overflow*. Python sigue y marca `ExcelHabriaFallado` (misma política que D-08).

### Semántica VBA confirmada (sin impacto en septiembre)
- `And`/`Or` no cortocircuitan: se evalúan todos los operandos (origen de F-11).
- Variant vs literal `String` es comparación de texto (`169 = "NO FAULTS"` → falso, sin error); Variant de texto no numérico vs número es *Type mismatch*; `Empty` vale 0 frente a números y `""` frente a textos. Implementado en `excel_semantics.celdas`.
- El bucle de ambas macros evalúa la salida (`A = ""`) **después** de procesar la fila: una A2 vacía no corta la lectura (se procesa como 0).

---

## 7. Nueva definición de negocio: `Exclusion_Matrix` (2026-09-24)

### F-37 🔴 Los eventos de exclusión vienen de `Exclusion_Matrix`, no de `PlantActivity!D`
- **Definición (negocio, 2026-09-24):** hoja `Exclusion_Matrix` con una columna por PCS; `0`/vacío = sin evento de exclusión (EE); `1` = el PCS forma parte de un EE y se considera que mantuvo **todos** sus módulos; `2` = forma parte de un EE y se consideran los módulos que tenía **antes** del inicio del evento (p. ej. 3 módulos antes y 0 durante → se consideran 3).
- **PlantActivity (chat F. Barrientos – A. Albornoz, 2026-09-24):** queda solo para "Only Operational Time?" (C21), que no se ha usado (propuesta a revisar con Revergy/LTSA). La columna D ya no excusa; se conserva como espejo de `Calc!BO`.
- **Referencia oficial:** macros, hojas y celdas **de septiembre**; de la macro de agosto solo se adopta la regla de la matriz (C31/L14 = "Yes" activan la exclusión).
- **Estructura** (libro `…_20260923_agosto_2026.xlsm`): fila 1 = `Date/time`, `PCS01`…`PCS61`, `Excused Event`, `Comments`; filas alineadas **por fila** con `RawData-PCS` (15.990 filas, 0 desalineadas). Solo trae datos de agosto: 4.982 celdas con 1 y 1.498 con 2; comentarios `CPF` (362 filas), `External` (11) y "Actualizacion de firware…" (1).
- **Regla implementada** (`enrichment.exclusion`, `availability`, `fault_events`): con flag "Yes" y `M < 4`, `0` → `C3 − M`; `1` → `(C3 − M)·(1 − 1) = 0`; `2` → baterías previas del tramo (fila anterior al primer 2: `C3 − M` si valor 0 y `M < C3`; `0` si completo, vacío o valor 1).
- **Validación:** contra el libro de agosto, C12 = 2976, C14 = 264.731,7679999999, C16 y las 2976 × 61 celdas de la tabla **idénticos bit a bit**. La macro de agosto implementa el 2 como `F(r) = F(r−1)` (copia encadenada); en todo el libro coincide con la definición de negocio. Difieren solo si un tramo de 2 contiene una fila sin falla (la macro dejaría 0 después) o empieza en la primera fila del período (la macro copiaría el encabezado y fallaría): se reportan como `ee2_difiere_macro_agosto` / `ee2_sin_fila_previa`.
- **Impacto en agosto:** C14 sin exclusión 527.396,044 → con exclusión 264.731,768 (−49,8 %). El KPI oficial de `Annual_AVA` (305.182,968) sigue sin reproducirse (D-11).
- **Versión:** `availability-v1.1-exclusion-matrix` (R1.6). Sin la hoja (libros hasta septiembre) los resultados son idénticos a v1: el golden de septiembre sigue `verified` sin cambios.

### F-38 🟡 El libro de agosto trae otra versión de macros y celdas (no adoptada)
- Parámetros movidos (`C4` baterías, `C6`/`C8` fechas, `C10` racks, `C19` excusable, sin C19 anual), C14 acumulado en variable, `Daily!D` = PCS-h y `Daily!E` = suma, `mcoCreateList` todavía pondera con `PlantActivity!D`.
- Por instrucción de negocio las definiciones oficiales siguen siendo las de septiembre; del libro de agosto solo se usan como referencia los valores cacheados de `Calculation-Availability` (C12, C14, C16, tabla E4:BO).
- `ListOfFaults` del libro de agosto **no se recalculó** (L2/L4/L10 de septiembre): no hay referencia Excel para eventos con exclusión.

### F-39 🟡 Marca aislada fuera de agosto
- `Exclusion_Matrix` fila 2 (2026-04-08 00:15): PCS61 = 1 y `Excused Event` = 1; el resto de la matriz fuera de agosto está vacía. Probable residuo; confirmar (D-18).

### Decisiones asociadas
- **D-16** (resuelta 2026-09-24, confirmada por negocio): los eventos (`mcoCreateList`, L14 = "Yes") usan la misma regla de la matriz que el KPI, para que `C14 = 4·L10` siga valiendo sin arrastre. La macro de agosto no lo hace (usa `PlantActivity!D`).
- **D-17** (resuelta 2026-09-24): Alex entrega la `Exclusion_Matrix` una vez al mes, al final. Corridas semanales = oficiales "Sin Exclusiones"; cierre mensual = oficial "Con Exclusiones" (vigente).
- **D-18**: la marca de la fila 2 se mantiene (F. Barrientos); pendiente confirmación de Alex.
- **D-19** (resuelta 2026-09-24): el maestro pasa a **v1.1** (septiembre + `Exclusion_Matrix` + regla en `cmdCalcAvailability` y `mcoCreateList`) para que Excel siga siendo referencia en las corridas "Con Exclusiones" (`design.md §Libro maestro v1.1`, tarea 4.0).

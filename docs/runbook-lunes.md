# Runbook — Flujo del lunes (SCADA → Excel → Python → SQL → Power BI)

Para: quien opera la corrida semanal (Francisco / Alex). Actualizado: 2026-09-25.
Diseño: `AGENTS.md` §14 Fase S y `.kiro/specs/etl-arena-availability/`. Power BI: `docs/pbi-handoff.md`.

**Reglas de oro**

- Server SCADA: **solo exportar**. Macros y ETL corren únicamente en la PC local.
- Solo `RawData-PCS` sale de SCADA. Las exclusiones salen **solo** de la `Exclusion_Matrix` que entrega Alex a
  fin de mes; `PlantActivity` no se usa para exclusiones hasta que Alex lo confirme.
- Nunca se edita el libro base ni `data/processed/`: cada corte trabaja sobre su propia copia en `data/work/<corte>/`.
- SQL es append-only: cada corrida tiene un `IdCorrida` nuevo; Power BI solo ve corridas oficiales en `success`.
- Mientras corre `run-macros`, **no usar Excel** (la etapa abre su propia ventana ~2 min y la cierra sola).

Todos los comandos se corren desde la raíz del repo, en PowerShell:
`.venv\Scripts\python scripts\run_lunes.py …` (abreviado abajo como `run_lunes …`).

---

## 0. Precondiciones

- [ ] **Azure SQL** (`trina-etl.database.windows.net`): la IP de la red está en el firewall del servidor (portal →
  `trina-etl` → Redes). La primera conexión del día tarda ~1 min (la base se reanuda) y puede abrir el navegador
  para iniciar sesión con la cuenta @trinasolar.com.
- [ ] **Excel** instalado en la PC. No hace falta *Trusted Location*: la corrida habilita macros solo en su propia
  instancia y quita la marca de "descargado de Internet" de la copia. Si una política de TI igual bloquea las
  macros, agregar `data\work\` como ubicación de confianza (Excel → Opciones → Centro de confianza).
- [ ] `data/inbox/` existe y está vacío (solo debe quedar el export del corte).
- [ ] Sabes cuál es el **libro base**: `data/work/libro_base.json` (campo `ruta`); si no existe, es el maestro de `data/`.
- [ ] `DESDE` del export = dato siguiente al último de `RawData-PCS` del libro base; `HASTA` = último dato disponible.

## 1. Export desde SCADA (~03:00)

- [ ] TeamViewer al server SCADA; exportar el reporte de `RawData-PCS` con `DESDE`/`HASTA`. No procesar nada allí.
- [ ] Transferir por TeamViewer a `data/inbox/raw_pcs_<AAAA-MM-DD>.<ext>`.

## 2. Preparar el libro (manual hasta tener la muestra SCADA, tareas 0.6/0.7)

Hasta que exista el lector del export, las filas nuevas se pegan a mano:

1. Copiar el libro base a `data/work/<corte>-preparado.xlsm`.
2. Pegar las filas nuevas **al final** de `RawData-PCS`, sin tocar las existentes. La columna A debe quedar como
   fecha de Excel (alineada a la derecha, formato `dd-mm-aaaa hh:mm:ss`), **nunca como texto** (F-21).
3. Guardar y cerrar Excel.

## 3. Corrida semanal

```powershell
run_lunes --stage all --corte 2026-09-28 --libro-preparado data\work\2026-09-28-preparado.xlsm --oficial
```

| Etapa | Qué hace | Resultado en `data/work/<corte>/` |
|---|---|---|
| `acquire-wait` | Espera el export en `data/inbox`, valida que no esté vacío ni copiándose, lo mueve a `data/processed/<corte>/` con su `.sha256` | — |
| `prepare-workbook` | Copia el libro preparado a `libro.xlsm` (+ `libro.xlsm.bak`) | `libro.xlsm` |
| `run-macros` | Escribe parámetros y corre las 4 macros por COM; extrae la referencia Excel | `referencia_excel.json` |
| `run-etl` | Motores Python → SQL (`IdCorrida` nuevo) + reconciliación contra la referencia; encola cierres de meses completos | fila en `etl_run` |
| `reconcile` | Detiene la cadena si `parity_failed`; si todo está bien y es `--oficial`, **promueve** la copia a libro base | `data/work/libro_base.json` |
| `notify-bi` | Webhook (`ETL_ARENA_NOTIFY_WEBHOOK`) o `notificacion.md` para Misael | `notificacion.md` |

- **Período por defecto:** del día 1 del mes del último dato hasta ese dato (D-07). Otro período: `--periodo-inicio`
  / `--periodo-fin` (AAAA-MM-DD).
- **Parámetros oficiales:** C21 = "No", C31 = L14 = "Yes" (D-03). Con las macros de septiembre, en Excel se escribe
  "No" en C31/L14 para no excusar con `PlantActivity!D` (F-44); sin exclusiones en el período da lo mismo.
- **Etiqueta:** las semanales son oficiales **"Sin Exclusiones"** (D-17).
- Opciones útiles: `--omitir-acquire` (el export ya se movió o no hay), `--omitir-macros` (sin Excel: la corrida
  queda `sin_referencia`), `--sin-bd` (ensayo sin escribir en SQL; no promueve el libro base).

### Reanudar

El estado de cada etapa queda en `data/work/<corte>/run_state.json`. Volver a correr el **mismo comando** salta las
etapas ya `ok` y sigue desde la que falló. `--forzar` repite las etapas desde `run-macros` (nunca repite
`acquire-wait` ni `prepare-workbook`: el export ya está en `data/processed` y la copia de trabajo no se pisa). Una
sola etapa: `run_lunes --stage run-etl --corte <corte> …`.

Códigos de salida: `0` ok · `2` `parity_failed` · `1` error (mensaje en `run_state.json` → `<etapa>.error`).

## 4. Leer el resultado

### Reconciliación (Python vs Excel)

`run_state.json` → `run-etl.artefactos.reconciliacion`: `estado` (`pass` / `fail` / `sin_referencia`) y, por nivel,
`comparaciones`, `fallidas` y `max_delta`.

| Nivel | Compara | Si falla |
|---|---|---|
| 1 `input` | Filas procesadas, primer/último serial, C23 y parámetros escritos | Revisar que el libro y el período sean los esperados (informativo: no bloquea) |
| 2 `muestra` | Cada celda de la tabla `Calculation-Availability!F:BN` | Buscar la fila/PCS en `reconciliation_result` |
| 3 `acumulados` | C12 y C14 | **Bloquea** (`parity_failed`) |
| 4 `kpi` | C16, C19 y `Daily` por día | **Bloquea** |
| 5 `eventos` | `ListOfFaults` evento a evento, L10, resumen por código | **Bloquea** |

Detalle en SQL:

```sql
SELECT Nivel, Metrica, Clave, ValorPython, ValorExcel, Delta, Tolerancia
FROM dbo.reconciliation_result WHERE IdCorrida = '<IdCorrida>' AND Aprobado = 0 ORDER BY Nivel;
```

`sin_referencia` no es un error: significa que no hubo macros (`--omitir-macros`, falla de Excel, o período con
exclusiones mientras el maestro no sea v1.1). La corrida se publica igual si es `success`.

Diferencias conocidas y aceptadas: eventos con Δ ≈ 1e-14 cuando L14 se escribe "No" (F-44); `ListOfFaults!N:Q`
sin ordenar en Excel 2016 (F-40, se compara como mapa).

### Calidad de datos

`run-etl.artefactos.anomalias` (`total`, `por_severidad`, `por_tipo`), en SQL `dbo.v_calidad_corrida`, y el
resumen completo en `etl_run.ResumenCalidad` (JSON; incluye `modulos_nulos`, eventos con descripción tomada de la
fila anterior y eventos arrastrados). Revisar:

| Tipo | Qué significa | Acción |
|---|---|---|
| `ModulosDisponiblesNulo` (`modulos_nulos` en el resumen; `v_modulos_nulos_historico`) | `NUMBER_OF_MODULES` vacío: se cuenta como disponible, igual que el Excel | Informar a operación si crece |
| `pa_vacio`, `pa_desalineado`, `pa_sin_timestamp` | `PlantActivity` sin dato o desalineada en filas nuevas (F-31) | Solo afecta si C21 = "Yes" (oficial = "No"); actualizarla aparte |
| `hueco`, `duplicado`, `fuera_de_orden`, `frecuencia_distinta` | Export incompleto o repetido | Revisar el export antes de publicar |
| `dst_salto`, `dst_repeticion` | Cambio de hora (septiembre / abril) | Esperado en esas fechas |
| `codigo_sin_catalogo` | Código de falla fuera de `PCS-Fault` | Agregar al catálogo si es nuevo |
| `em_desalineado`, `em_sin_timestamp`, `ee2_*` | `Exclusion_Matrix` desalineada o con un 2 sin fila previa | Revisar la entrega con Alex |
| `exclusion_matrix_ausente` | El libro no trae la matriz | Normal en semanales antes de la entrega de Alex |

## 5. Cierre mensual

Cuando los datos llegan al último bloque de un mes (último día 23:45) sin cierre oficial, `run-etl` lo agrega a
`data/work/cola_cierres.json` y la notificación lo muestra como **cierre pendiente**.

**Con la `Exclusion_Matrix` de Alex** (oficial "Con Exclusiones", el caso normal):

```powershell
run_lunes --stage load-exclusion-matrix --mes 2026-09 --archivo-matriz data\inbox\<entrega>.xlsx
```

- Formato esperado (hasta confirmar con la muestra de Alex, 0.8): hoja `Exclusion_Matrix` con `Date/time` (fecha de
  Excel), `PCS01…PCS61` (0, 1, 2 o vacío), `Excused Event`, `Comments`. Solo se toman las filas del mes; un
  timestamp que no exista en `RawData-PCS` o un valor fuera de 0/1/2 detiene la etapa con la lista.
- Trabaja en `data/work/matriz-AAAA-MM/`: copia el libro base, escribe la matriz por Excel, deja `cambios.csv`
  (qué celdas cambiaron; abre en Excel), registra la carga y las correcciones en SQL y corre el cierre del mes.
  El libro con la matriz pasa a ser el libro base.
- Mientras el maestro no sea v1.1 (4.0), `run-macros` se omite solo (las macros de septiembre no leen la matriz):
  la corrida queda `sin_referencia`.

**Sin esperar la matriz** (decisión explícita, R19.6; queda oficial "Sin Exclusiones"):

```powershell
run_lunes --stage cierre-mensual --mes 2026-09 --sin-exclusiones
```

Sin `--sin-exclusiones`, `cierre-mensual` se niega si la matriz del mes no está cargada.

## 6. Correcciones de datos ya cargados (D-13)

`--reproceso` (tarea 4.13) aún no existe: depende del formato real del export SCADA. Mientras tanto, corregir en
una copia nueva del libro base, correr como un corte nuevo con `--tipo reproceso --oficial` y anotar en el shadow
log qué se corrigió y por qué.

## 6b. Entorno de prueba: cargar, validar y después cargar lo definitivo

Para probar una carga sin tocar la base de producción ni el libro base, cualquier comando acepta
`--entorno prueba`:

| | `produccion` (por defecto) | `prueba` |
|---|---|---|
| Base | `ETL_ARENA_DB_URL` (`trina_etl`) | `ETL_ARENA_DB_URL_PRUEBA` (p. ej. `trina_etl_prueba`); falla si apunta a producción |
| Carpeta de trabajo | `data/work/` | `data/work/_prueba/` (libro base y cola de cierres propios) |
| Export en `data/inbox` | Se **mueve** a `data/processed/<corte>/` | Se **copia** a `data/processed/_prueba/<corte>/` (queda para la corrida definitiva) |
| Notificación | Webhook o `notificacion.md` | Solo `notificacion.md`, titulado `[PRUEBA]` (nunca avisa a Misael) |

**Una sola vez:** crear la base de prueba en el portal de Azure (servidor `trina-etl` → Crear base de datos →
`trina_etl_prueba`; aplicar la oferta gratuita si el portal la ofrece y, si no, revisar el costo antes de crear; **sin** "test" en el nombre, porque los tests de integración reinician
esas bases), agregar a `.env`
`ETL_ARENA_DB_URL_PRUEBA=mssql+pyodbc://@trina-etl.database.windows.net:1433/trina_etl_prueba?driver=ODBC+Driver+18+for+SQL+Server&Encrypt=yes&TrustServerCertificate=no`
y aplicar el esquema:

```powershell
.venv\Scripts\python scripts\crear_base.py --entorno prueba
```

**Cada prueba:**

```powershell
run_lunes --entorno prueba --stage all --omitir-acquire --corte 2026-09-21 --libro-preparado "<libro>" --oficial
```

Validar en la base de prueba (mismas vistas que verá Power BI: `v_kpi_vigente`, `v_daily_vigente`, …) y en
`data/work/_prueba/<corte>/`. Se puede repetir con otro `--corte`; para empezar de cero, borrar `data/work/_prueba/`
(la base de prueba se puede vaciar recreándola en el portal).

**Carga definitiva:** el mismo comando **sin** `--entorno prueba`, cuando se apruebe. No se reutiliza nada de la
prueba: la corrida definitiva vuelve a calcular desde el libro.

## 7. Registro del shadow mode (5.4)

Mientras Excel siga siendo la fuente oficial, después de cada corrida:

```powershell
.venv\Scripts\python scripts\shadow_log.py --corte 2026-09-28 --kpi-excel-oficial 0.9819 --nota "…"
```

Agrega una fila a `docs/shadow-log.md` con el resultado de la corrida y el KPI que reportó Alex con su Excel.

## 8. Recuperación ante fallos

| Situación | Qué hacer |
|---|---|
| `acquire-wait`: no llegó el archivo / hay varios | Dejar un solo `raw_pcs_*` en `data/inbox` y repetir el comando |
| `acquire-wait`: "ya existe en data/processed" | El corte ya se adquirió: usar otro `--corte` o `--omitir-acquire` |
| `prepare-workbook`: "ya existe libro.xlsm" | Es una corrida repetida: reanudar sin `--forzar`, o usar otro `--corte` |
| `run-macros`: error VBA | El texto del diálogo queda en `run_state.json`. El 438 en `mcoCreateList`/`Graphupdate` con Excel 2016 se tolera solo (F-40); cualquier otro: revisar el libro preparado (fechas como texto, filas vacías intermedias) |
| `run-macros`: "abierto en otro Excel" | Cerrar el libro en Excel y repetir |
| `run-macros`: timeout / Excel colgado | La etapa cierra su instancia. Repetir una vez con `--forzar`; si persiste, `run_lunes --stage run-etl …` (queda `sin_referencia`) |
| Excel quedó abierto tras cortar Python | Cerrar esa ventana de Excel a mano (no guardar) |
| `run-etl`: no conecta a Azure | Firewall (IP nueva) o base despertando: esperar 1 min y repetir |
| `reconcile`: `parity_failed` | No se publica. Revisar `reconciliation_result` (§4), corregir la causa y repetir con `--forzar` |
| "el libro base cambió después de promoverse" | Alguien editó `data/work/<corte>/libro.xlsm` de un corte ya promovido. Restaurar desde `libro.xlsm.bak` o avisar a Francisco |
| `load-exclusion-matrix`: timestamps inexistentes / valores inválidos | Devolver la lista a Alex, corregir la entrega y repetir |
| TeamViewer caído | Reintentar; no hay otro transporte hoy (RPA solo como P8) |

## 9. Checklist de cierre del lunes

- [ ] `run_lunes` terminó con código 0 y `run-etl` en `success`.
- [ ] Reconciliación `pass` (o `sin_referencia` con motivo conocido).
- [ ] Anomalías revisadas; nada nuevo sin explicar.
- [ ] Notificación enviada a Misael (o `notificacion.md` reenviado).
- [ ] Cierres pendientes anotados si la notificación los muestra.
- [ ] Shadow log actualizado (§7).

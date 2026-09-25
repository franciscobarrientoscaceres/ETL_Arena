# Design Document — ETL Arena Availability

> Revisión 2 — 2026-09-24. Corrige la revisión 1 contra el VBA real (`olevba`), las fórmulas del libro y el perfilado de datos. Hallazgos `F-xx` y decisiones `D-xx` en `audit.md`. Requisitos `R<n>.<m>` en `requirements.md`.

## Overview

El sistema reproduce en Python la lógica de las macros VBA del libro `AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm` (activo **Arena BESS**) y persiste en SQL Server el KPI y todos sus intermedios, con auditoría por `IdCorrida`.

**Principio de paridad:** los motores v1 son **emulaciones** del VBA — mismo orden de recorrido, mismas comparaciones de celdas, misma aritmética sobre seriales Excel y mismos defectos (marcados con flags). Toda "mejora" es una versión posterior del algoritmo.

**Principio de trazabilidad:** SQL guarda no solo el porcentaje final sino los intermedios para explicar cada valor hasta la celda de origen.

---

## Decisiones de diseño

| ID | Decisión | Motivo |
|---|---|---|
| ADR-01 | Paquete Python `etl_arena` en `src/etl_arena/<módulo>/` (layout *src*) | Evita colisión de nombres de paquetes top-level (`config`, `staging`) |
| ADR-02 | Los motores consumen una **matriz por fila origen** (`MatrizPCS`: arrays `n_filas × n_pcs`), no un DataFrame largo | El VBA depende de fila anterior/siguiente de la hoja completa; la matriz lo emula directo y es rápida (≈1M celdas) |
| ADR-03 | Comparaciones y restas de tiempo sobre `SerialFechaExcelOrigen` (`float64`) leído del `<v>` crudo del XML | Paridad a 1e-9 en duraciones (F-15) |
| ADR-04 | Semántica Excel centralizada en `etl_arena.excel_semantics` | Única fuente de reglas VBA/Excel (R17) |
| ADR-05 | Toda escritura en el `.xlsm` de trabajo por COM (`pywin32`); lectura por XML streaming | `openpyxl` pierde objetos del libro (F-22); lectura COM es lenta |
| ADR-06 | El pipeline Python lee **el mismo libro de trabajo** que usan las macros | Garantiza input idéntico para la reconciliación |
| ADR-07 | Timestamps naive hora local Chile; sin localizar ni convertir a UTC | Datos origen naive con salto DST (F-32) |
| ADR-08 | SQL append-only por `IdCorrida`; workflow de revisión en tabla aparte | R11, F-27 |
| ADR-09 | Tablas de muestras con *clustered columnstore* | Volumen ~1M filas/tabla/corrida (F-28) |
| ADR-11 | `Exclusion_Matrix` (0/1/2 por fila y PCS) reemplaza a `PlantActivity!D` como fuente de la exclusión; unión por fila; valor 2 = baterías previas al EE (definición de negocio, igual a la macro de agosto en todo el libro) | Definición de negocio 2026-09-24 (F-37); `docs/adr/ADR-11-exclusion-matrix.md` |
| ADR-10 | Python ≥ 3.13, ODBC Driver 18; **producción en Azure SQL Database serverless (oferta gratuita, `trina-etl`/`trina_etl`, Brazil South) con Microsoft Entra ID**; desarrollo opcional en instancia local o Docker | Entorno (F-29); en Trina no se permite instalar SQL Server (2026-09-24); detalle en `docs/adr/ADR-10-entorno.md` |

Convención de nombres: **SQL** en PascalCase español (`BloquesMuestreo`); **Python** en snake_case español (`bloques_muestreo`); los nombres de celda Excel (C12, L14…) se citan en comentarios.

---

## Architecture

```text
SCADA server ──(solo export, TeamViewer)──► data/inbox/raw_pcs_<corte>.<ext>
                                                  │
 [S] acquisition.acquire_wait ────────────────────┤ valida contrato, sha256 → data/processed/<corte>/
                                                  ▼
 [T] workbook.preparar (COM) ── copia maestro → data/work/<corte>/libro.xlsm (+ backup)
                                 escribe RawData-PCS (seriales reales), valida alineación PlantActivity
                                                  ▼
 [M] workbook.macros (COM) ──── setea C5 C7 C21 C31 L2 L4 L14 Daily!D5
                                 cmdCalcAvailability → mcoCreateList → mcoDailyAvailability → Graphupdate
                                 extrae referencia completa → referencia_excel.json + excel_reference_*
                                                  ▼
 [E] pipeline (Python puro, lee libro.xlsm por XML)
        ingestion ─► normalization ─► enrichment ─► MatrizPCS
                                                     ├─► availability (cmdCalcAvailability)
                                                     ├─► fault_events (mcoCreateList)
                                                     └─► aggregation (Daily, Annual_AVA)
        persistence ─► SQL Server (1 transacción por corrida)
                                                  ▼
 [R] reconciliation ── Python vs excel_reference_* (5 niveles + invariantes) → reconciliation_result
                                                  ▼
 [B] notify ── webhook / notificacion.md (Misael) ; PBI lee vistas v_*_vigente
```

Dependencias permitidas: `excel_semantics` y `config` no dependen de nada; los motores dependen solo de `excel_semantics`, `config` y `model`; `persistence` y `workbook` son adaptadores de borde. Los motores no importan pandas, pyodbc ni pywin32.

---

## Estructura del repositorio

```text
pyproject.toml            .env.example            docker/mssql.compose.yml
src/etl_arena/
  config/                 models.py (ConfiguracionCalculo), defaults.py, desde_excel.py
  excel_semantics/        celdas.py, texto.py, fechas.py, redondeo.py
  model/                  matriz.py (MatrizPCS, DatosActividad), anomalias.py (Anomalia)
  acquisition/            acquire_wait.py, contrato_scada.py, lector_scada.py
  workbook/               com.py (sesión Excel), preparar.py, macros.py, referencia.py
  ingestion/              xlsx_stream.py, lector_libro.py, anomalias_timestamp.py
  normalization/          esquema.py (validador), normalizador.py (ancho→matriz, matriz→largo)
  enrichment/             actividad_planta.py
  availability/           motor.py (MotorDisponibilidad), resultado.py
  fault_events/           motor.py (MotorEventosFalla), resumen_codigos.py
  aggregation/            diaria.py, mensual_anual.py
  persistence/            conexion.py, repositorio.py, corrida.py
  reconciliation/         niveles.py, invariantes.py, servicio.py, reporte.py
  reporting/              calidad.py, notificacion.py
  pipeline.py             orquesta E (config → … → persistencia)
scripts/  ejecutar_etl.py  run_lunes.py  generar_seed_tipo_detencion.py
sql/      00_database.sql … 07_audit_queries.sql
tests/    unit/  property/  golden/  integration/ (marker sql)  com/ (marker excel)  fixtures/
docs/     runbook-lunes.md  data-contract-scada.md  data-contract-libro.md  adr/
data/     inbox/  processed/  work/   (ignorados por git salvo .gitkeep)
```

---

## Components and Interfaces

### config

```python
@dataclass(frozen=True)
class ConfiguracionCalculo:
    id_corrida: str
    id_proyecto: int                       # 1 = Arena
    version_algoritmo: str                 # "availability-v1.1-exclusion-matrix" (F-37)
    nombre_proyecto: str
    fecha_inicio_proyecto: date
    total_pcs: int                         # C2
    baterias_por_pcs: int                  # C3 (y umbral literal "< 4" del VBA)
    racks_por_pcs: int                     # literal 12 del VBA
    minutos_muestreo: int                  # esperado; C23 se deriva de los datos
    # parámetros KPI (cmdCalcAvailability)
    inicio_periodo: date                   # C5
    fin_periodo: date                      # C7 (inclusivo a nivel día)
    solo_tiempo_operacional: bool          # C21 <> "No"
    aplicar_evento_excusable: bool         # C31 = "Yes" → aplica Exclusion_Matrix (F-37)
    # parámetros de eventos (mcoCreateList) — F-05
    inicio_periodo_eventos: date           # ListOfFaults!L2
    fin_periodo_eventos: date              # ListOfFaults!L4
    aplicar_evento_excusable_eventos: bool # ListOfFaults!L14 = "Yes" → aplica Exclusion_Matrix (F-37)
    # Daily
    fin_diario: date                       # Daily!D5 (≤ inicio + 30 días)
    modo_huecos: Literal["excel", "continuar"]
    tipo_corrida: Literal["semanal", "cierre_mensual", "reproceso", "golden"]  # período por tipo: D-07
    es_oficial: bool
    archivo_origen: str
    zona_horaria: str = "America/Santiago" # documental (ADR-07)

    @property
    def total_racks(self) -> int:
        return self.total_pcs * self.baterias_por_pcs * self.racks_por_pcs
```

`config.desde_excel.leer_parametros(ruta_libro)` construye los parámetros desde las celdas del libro usando `excel_semantics.flag_si/flag_no` (R1.5). `construir_config(**overrides)` aplica `CONFIG_POR_DEFECTO` y valida: `fin ≥ inicio`, `fin_diario − inicio ≤ 30 días`, enteros positivos.

### excel_semantics (R17)

| Función | Regla Excel/VBA emulada |
|---|---|
| `es_vacio(v)` | Celda vacía: `None` o `""` (`Cells(...) = ""`) |
| `igual_numero(v, n)` | `Cells(...) = n` con `v` numérico; texto no numérico → `ExcelTypeMismatch` |
| `menor_que(v, n)` | `Cells(...) < n`; texto no numérico → `ExcelTypeMismatch` |
| `igual_texto(v, s)` | Comparación binaria exacta (sin `Option Compare Text`) |
| `flag_si(v)` / `flag_no(v)` | `v = "Yes"` / `v = "No"` exactos |
| `texto_excel(v)` | Número→texto con formato General (`55` → `"55"`, `3.5` → `"3.5"`, hasta 15 dígitos significativos) |
| `codigo_falla_excel(g)` | `IFERROR(MID(G,1,FIND(" ",G,1)-1), CONCATENATE("F",G))`: `s = texto_excel(g)`; `p = s.find(" ")`; `p == -1 → "F"+s`; si no `s[:p]` (vacío si `p == 0`) |
| `redondear_excel(x, d)` | `ROUND` half-away-from-zero (no el redondeo bancario de Python) |
| `serial_a_datetime(s)` / `datetime_a_serial(dt)` | Epoch 1899-12-30; sin zona |

### acquisition (Fase S — R2)

```python
def acquire_wait(inbox: Path, patron: str, timeout_s: int, poll_s: int = 30) -> ArchivoAdquirido: ...
    # espera, valida no vacío/legible/nombre, sha256, mueve a data/processed/<corte>/ (+ .sha256)
def validar_contrato(archivo: ArchivoAdquirido, contrato: ContratoSCADA) -> list[Anomalia]: ...
    # encabezados == RawData-PCS (orden incluido), delimitador, encoding, formato fecha
def leer_export(archivo, contrato) -> TablaExport: ...
    # parser estricto "%m-%d-%Y %H:%M:%S" → serial Excel; valida rango DESDE/HASTA (R2.5)
```

`ContratoSCADA` se define en P1 a partir de la muestra real P0 (`docs/data-contract-scada.md`). **Hasta P0, este módulo no se implementa más allá de interfaces y tests con fixtures sintéticos.**

### workbook (Fases T/M — R3)

```python
class SesionExcel:                               # context manager
    """Crea su propia instancia (DispatchEx), Visible=True, DisplayAlerts=False,
    AutomationSecurity=1 (libro en Trusted Location). Al salir cierra SOLO su instancia;
    watchdog con timeout mata su PID si cuelga."""

def preparar_libro(maestro: Path, export: TablaExport, destino: Path) -> ResultadoPreparacion:
    """Copia libro base→destino (+ .bak); libro base = libro de trabajo de la última corrida
    oficial (o el maestro en la primera). Valida continuidad con la última fila de RawData-PCS
    (R2.5, D-07/D-13). Vía COM: AGREGA las filas nuevas a continuación (no toca las existentes)
    por bloques con Range.Value (col A = serial float, formato visual dd-mm-aaaa hh:mm:ss),
    guarda. Valida alineación por fila con PlantActivity!B (R3.4) y devuelve anomalías."""

def ejecutar_macros(libro: Path, cfg: ConfiguracionCalculo, timeout_s=900) -> None:
    """Escribe C5, C7, C21 ("Yes"/"No"), C31, ListOfFaults!L2/L4/L14, Daily!D5;
    Application.Run de las 4 macros en orden; guarda."""

def extraer_referencia(libro: Path) -> ReferenciaExcel:
    """Lectura por XML (post-guardado) de: C12 C14 C16 C19 C23, L10, parámetros efectivos,
    tabla E4:BO (seriales + ponderadas por PCS + BO), ListOfFaults B6:I completa en orden,
    N6:Q172, Daily B9:G39. Serializa a referencia_excel.json."""
```

### Libro maestro v1.1 (D-19, F-37)

Las macros oficiales de septiembre no conocen `Exclusion_Matrix`. Para que Excel siga siendo referencia en las corridas "Con Exclusiones", el maestro pasa a una versión **v1.1** preparada a mano (tarea 4.0; el pipeline nunca edita VBA):

1. Copia del libro de septiembre con nombre versionado (`…_maestro_v1.1.xlsm`); se conserva el original.
2. Hoja `Exclusion_Matrix` con la estructura de §3b del data contract (encabezado `Date/time`, `PCS01…PCS61`, `Excused Event`, `Comments`), alineada por fila con `RawData-PCS`.
3. `cmdCalcAvailability` — solo la rama `C31 = "Yes"`, igual que la macro de agosto pero con las celdas de septiembre (`C3`, no `C4`); la hoja se referencia **por nombre**, no por codeName:

```vba
Set EM = Worksheets("Exclusion_Matrix")          ' antes del Do
...
If cmdAvailability.Cells(31, 3) = "Yes" Then
    If EM.Cells(dblRec, intRecPCS + 1) = 2 Then
        cmdAvailability.Cells(dblResult, intRecPCS + 5) = cmdAvailability.Cells(dblResult - 1, intRecPCS + 5)
    Else
        cmdAvailability.Cells(dblResult, intRecPCS + 5) = (cmdAvailability.Cells(3, 3) - Sheet2.Cells(dblRec, 4 * intRecPCS + 1)) _
            * (1 - EM.Cells(dblRec, intRecPCS + 1))
    End If
Else
    cmdAvailability.Cells(dblResult, intRecPCS + 5) = cmdAvailability.Cells(3, 3) - Sheet2.Cells(dblRec, 4 * intRecPCS + 1)
End If
```

4. `mcoCreateList` — código nuevo (la macro de agosto seguía con `PlantActivity!D`; D-16). Variable `pondAnterior` (Double) reiniciada a 0 al empezar cada PCS; en **cada** fila del bucle interno (dentro o fuera del período), antes de la condición de período:

```vba
e = EM.Cells(dblRec, 0.5 + intColRec / 4 + 1)
If Sheet2.Cells(dblRec, intColRec + 3) <> "" And Sheet2.Cells(dblRec, intColRec + 3) < 4 Then
    If e = 2 Then pondFila = pondAnterior Else pondFila = (4 - Sheet2.Cells(dblRec, intColRec + 3)) * (1 - e)
Else
    pondFila = 0
End If
' … en la rama L14 = "Yes":  sumablocks = sumablocks + pondFila
pondAnterior = pondFila                           ' al final de la fila
```

5. `Graphupdate`, `mcoDailyAvailability`, `Daily`, `Annual_AVA` y todas las celdas de parámetros: **sin cambios** (septiembre).

Diferencias conocidas con los motores Python (reportadas como anomalías, la reconciliación las marca "explicadas"): en `cmdCalcAvailability` el valor 2 copia la fila anterior **del período** (si el tramo empieza en la primera fila del período copia el encabezado y la macro falla) y una fila sin falla dentro del tramo corta la copia (`ee2_difiere_macro_agosto`, `ee2_sin_fila_previa`).

Notas COM: `Graphupdate` usa `ActiveWindow.SmallScroll` y todas las macros usan `.Select` → la instancia debe ser visible y el libro activo. `data/work/` debe estar registrado como *Trusted Location* (documentado en el runbook).

### ingestion (R4)

`xlsx_stream.leer_hoja(ruta, nombre_hoja)` recorre `xl/worksheets/sheetN.xml` con `xml.etree.ElementTree.iterparse` (resolviendo `sharedStrings` y el mapeo nombre→archivo de `workbook.xml.rels`) y entrega filas de valores crudos: `float` para números y fechas (serial exacto del `<v>`), `str`, `bool` o `None`. No usa `openpyxl` para datos (conversión a `datetime` pierde el serial exacto; ADR-03).

```python
def leer_libro(ruta: Path, cfg) -> LibroCrudo:
    # RawData-PCS: encabezado fila 1, datos desde fila 2, n_filas según modo_huecos (R4.5)
    # PlantActivity: filas 2.., columnas B (ts), C (op), D (exc), E..I (informativas)
def detectar_anomalias_timestamp(seriales, minutos) -> list[Anomalia]
    # hueco, duplicado, fuera_de_orden, frecuencia_distinta (redondear_excel(Δ*1440, 2)),
    # dst_salto / dst_repeticion (informativas), celda_a_vacia, filas_truncadas
```

`Anomalia(tipo, severidad, numero_fila, numero_pcs, serial, detalle)`; severidades `info | advertencia | error`. Un `error` en modo paridad (p. ej. módulos texto, R4.6) aborta la corrida.

### normalization (R5)

```python
PATRON = re.compile(r"^Arena - PCS (\d{2}) - POWERELECTRONICS (.+)$")
CAMPOS = {
    "GEN3 HEx CURRENT FAULT":   "falla",
    "GEN3 HEx CURRENT STATUS":  "estado",
    "GEN3 HEx CURRENT WARNING": "advertencia",
    "HEM-k NUMBER OF MODULES":  "modulos",
}
def validar_esquema(encabezado, cfg) -> tuple[MapaColumnas, list[Anomalia]]
    # PCS k esperado en columnas (1-based) 4k-2 .. 4k+1; reporta faltantes/desplazadas/inesperadas
def a_matriz(libro: LibroCrudo, mapa, cfg) -> MatrizPCS
def a_formato_largo(m: MatrizPCS) -> pd.DataFrame    # solo para staging/persistencia
```

```python
@dataclass
class MatrizPCS:                    # n = filas de datos (orden origen), p = PCS en encabezado
    numero_fila: np.ndarray         # (n,) int, fila Excel (2..)
    serial: np.ndarray              # (n,) float64
    marca_tiempo: np.ndarray        # (n,) datetime64 naive local
    modulos: np.ndarray             # (n,p) float64; NaN si vacío
    modulos_nulo: np.ndarray        # (n,p) bool
    falla: np.ndarray               # (n,p) object: str | float | None (tipo crudo)
    estado: np.ndarray              # (n,p) object
    advertencia: np.ndarray         # (n,p) object
    pcs: list[int]                  # números de PCS en orden de encabezado
    encabezado_falla: list[str]     # texto de fila 1 por PCS (para D-08)
```

`ModulosDisponibles` para persistencia = `modulos` o `baterias_por_pcs` si nulo (R5.3); los motores usan `modulos_nulo` explícitamente.

### enrichment (R6)

```python
def asociar_actividad(m: MatrizPCS, actividad: LibroCrudo.actividad) -> DatosActividad:
    # por índice de fila: factor_operacional[i] = PlantActivity!C[fila]; evento_excusado_pa[i] = !D[fila]
    # (D solo espejo de Calc!BO, no pondera — F-37); C vacía → 0.0 + Anomalia(pa_vacio); ts distinto → pa_desalineado

def asociar_exclusion(m: MatrizPCS, exclusion: LibroCrudo.exclusion | None, cfg) -> DatosExclusion:
    # Exclusion_Matrix por índice de fila (F-37): valor[i, j] = hoja!(fila, j+2) ∈ {0, 1, 2}; vacío = 0
    # sin hoja → todo 0 + Anomalia(exclusion_matrix_ausente); otro valor o encabezado distinto → ErrorParidad
    # baterias_previas[i, j] (valor 2): por tramo contiguo de 2, fila anterior al tramo (hoja completa):
    #     C3 − M si M < C3 y valor 0; 0 si M = C3, vacía o valor 1
    # columnas "Excused Event" (resumen por fila) y "Comments" (causa) → trazabilidad
```

### availability — `cmdCalcAvailability` (R7)

```python
@dataclass
class ResultadoDisponibilidad:
    bloques_muestreo: int                    # C12
    bloques_racks_indisponibles: float       # C14
    disponibilidad_periodo: float | None     # C16
    disponibilidad_anual_acumulada: float | None  # C19
    minutos_muestreo_derivado: float | None  # C23
    filas_procesadas: np.ndarray             # índices i en rango, orden origen
    ponderadas: np.ndarray                   # (len(filas_procesadas), p) — espejo de Calc!F:BN
    muestras: list[MuestraDisponibilidad]    # availability_sample_result

def calcular(m: MatrizPCS, act: DatosActividad, cfg) -> ResultadoDisponibilidad:
    ini = datetime_a_serial(cfg.inicio_periodo)
    fin1 = datetime_a_serial(cfg.fin_periodo) + 1
    c12 = 0; c14 = 0.0; procesadas = []
    for i in range(m.n):                                   # orden de fila origen (R7.1)
        s = m.serial[i]
        if not (s >= ini and s < fin1):
            continue
        for j, pcs in enumerate(m.pcs[:cfg.total_pcs]):    # VBA: For intRecPCS = 1 To C2
            if not m.modulos_nulo[i, j] and m.modulos[i, j] < cfg.baterias_por_pcs:
                bat = cfg.baterias_por_pcs - m.modulos[i, j]          # C3 - celda
                if cfg.aplicar_evento_excusable:                       # F-37: Exclusion_Matrix
                    e = exc.valor[i, j]
                    pond = exc.baterias_previas[i, j] if e == 2 else bat * (1 - e)
                else:
                    pond = bat
                if cfg.solo_tiempo_operacional:
                    impacto = cfg.racks_por_pcs * pond * act.factor_operacional[i]
                else:
                    impacto = cfg.racks_por_pcs * pond
                c14 = c14 + impacto                                   # orden fila→PCS (R7.7)
            else:
                bat = pond = impacto = 0.0
            registrar_muestra(i, pcs, bat, pond, impacto)
        c12 += 1                                           # por fila, no por timestamp (R7.2)
        procesadas.append(i)
    c23 = redondear_excel((m.serial[procesadas[1]] - m.serial[procesadas[0]]) * 24 * 60, 2) \
          if len(procesadas) >= 2 else None
    c16 = 1 - c14 / (cfg.total_racks * c12) if c12 and cfg.total_racks else None
    c19 = 1 - c14 / (cfg.total_racks * (365 * 24 * 4)) if cfg.total_racks else None
    return ResultadoDisponibilidad(c12, c14, c16, c19, c23, ...)
```

Si `c23` es `None` (menos de 2 filas) el MotorEventosFalla no puede calcular inicios → error de corrida con mensaje explícito.

### fault_events — `mcoCreateList` (R8)

El motor **emula las escrituras de celda** en `ListOfFaults`: un registro mutable `actual` representa la fila `dblResult`; se completan `B,C,F,G` al iniciar y `D,E,H,I` al cerrar; `dblResult` avanza solo al cerrar. Los acumuladores no se reinician al cambiar de PCS (F-06).

```python
@dataclass
class RegistroLista:                     # una fila B:I de ListOfFaults
    orden_excel: int
    numero_pcs: int | None = None        # B
    serial_inicio: float | None = None   # C
    serial_fin: float | None = None      # D
    duracion_horas: float | None = None  # E
    codigo_falla: str | None = None      # F
    descripcion_falla: object = None     # G (str | float)
    promedio_baterias: float | None = None  # H
    horas_rack: float | None = None      # I
    numero_bloques: int = 0
    suma_bloques: float = 0.0
    fallback: bool = False
    arrastrado_excel: bool = False
    excel_habria_fallado: bool = False
    pcs_iniciados: set[int] = field(default_factory=set)

def detectar_eventos(m, act, cfg, c23) -> ResultadoEventos:
    L2 = datetime_a_serial(cfg.inicio_periodo_eventos)
    L4_1 = datetime_a_serial(cfg.fin_periodo_eventos) + 1
    UMBRAL = cfg.baterias_por_pcs          # literal 4 en el VBA
    cerrados, actual = [], RegistroLista(orden_excel=1)
    sumablocks, numblock, l10 = 0.0, 0, 0.0
    for j, pcs in enumerate(m.pcs):        # hasta encabezado vacío (F-17)
        for i in range(m.n):
            s = m.serial[i]
            if not (s >= L2 and s < L4_1):
                continue
            if m.modulos_nulo[i, j] or not (m.modulos[i, j] < UMBRAL):
                continue
            if cfg.aplicar_evento_excusable_eventos:
                e = exc.valor[i, j]                                     # F-37, D-16
                sumablocks = sumablocks + (exc.baterias_previas[i, j] if e == 2 else (UMBRAL - m.modulos[i, j]) * (1 - e))
            else:
                sumablocks = sumablocks + UMBRAL - m.modulos[i, j]   # (sumablocks + 4) - x
            numblock += 1

            # ---- inicio (R8.4) ----
            if i == 0:                                   # anterior = encabezado (D-08)
                inicia = True; actual.excel_habria_fallado = True
            else:
                inicia = (m.modulos_nulo[i-1, j] or m.modulos[i-1, j] == UMBRAL
                          or m.serial[i-1] < L2)
            if inicia:
                if actual.numero_pcs is not None:        # sobrescribe registro no cerrado
                    actual.arrastrado_excel = True
                actual.numero_pcs = pcs
                actual.pcs_iniciados.add(pcs)
                actual.serial_inicio = fecha_vba_a_celda(s - c23 / (24 * 60))   # Date → celda al segundo (F-34)
                g, fallback = m.falla[i, j], False
                if igual_texto(g, "NO FAULTS"):
                    prev = m.falla[i-1, j] if i > 0 else None
                    if not igual_texto(prev, "NO FAULTS"):
                        g, fallback = prev, True         # prev vacío → g vacío → F1 (F-04)
                    else:
                        g = "F13 NO MODULES"
                if es_vacio(g):
                    g = "F1 Watchdog"
                actual.descripcion_falla = g
                actual.fallback = fallback
                actual.codigo_falla = codigo_falla_excel(g)

            # ---- cierre (R8.5) ----
            if i + 1 >= m.n:
                cierra = True                            # siguiente fila vacía
            else:
                cierra = (m.modulos_nulo[i+1, j] or m.modulos[i+1, j] == UMBRAL
                          or m.serial[i+1] > L4_1)       # estricto: '>' (F-06)
            if cierra:
                if pcs not in actual.pcs_iniciados or len(actual.pcs_iniciados) > 1:
                    actual.arrastrado_excel = True
                actual.serial_fin = fecha_vba_a_celda(s) # última fila en falla (F-03, F-34)
                actual.duracion_horas = 24 * (actual.serial_fin - actual.serial_inicio)
                actual.promedio_baterias = sumablocks / numblock
                actual.horas_rack = cfg.racks_por_pcs * actual.duracion_horas * actual.promedio_baterias
                actual.numero_bloques, actual.suma_bloques = numblock, sumablocks
                l10 = l10 + cfg.racks_por_pcs * actual.duracion_horas * actual.promedio_baterias
                cerrados.append(actual)
                actual = RegistroLista(orden_excel=actual.orden_excel + 1)
                sumablocks, numblock = 0.0, 0
    incompleto = actual if actual.numero_pcs is not None else None   # quedó escrito sin D:I
    return ResultadoEventos(cerrados, incompleto, l10, resumen_por_codigo(cerrados, catalogo))
```

- `EventoFalla` (persistencia) se construye desde cada `RegistroLista` cerrado; `MarcaTiempoInicio/Fin = serial_a_datetime(serial_*)`.
- `resumen_por_codigo`: emula `ListOfFaults!N:Q` — por cada código del catálogo `P = Σ I` de eventos con `F = código` (`SUMIF`), `Q = P / L10`, orden descendente por `P` (`mcoOrder`). No hay límite de eventos (F-18).
- Un registro `incompleto` no suma a `L10`; se reporta en calidad (`evento_incompleto_excel`).

### aggregation — `mcoDailyAvailability` y `Annual_AVA` (R9, R10)

```python
def calcular_diaria(res: ResultadoDisponibilidad, m, cfg) -> list[DiaDisponibilidad]:
    n_dias = (cfg.fin_diario - cfg.inicio_periodo).days + 1          # ≤ 31; fin_diario = última Daily!C (F-33)
    dias = [datetime_a_serial(cfg.inicio_periodo) + d for d in range(n_dias)]
    diario = [0.0] * n_dias
    k, i, suma = 0, 0, 0.0
    filas = res.filas_procesadas
    while True:                                                       # emulación del Do…Loop
        if k < len(filas) and dias[i] <= m.serial[filas[k]] < dias[i] + 1:
            for j in range(cfg.total_pcs):
                suma = suma + cfg.racks_por_pcs * res.ponderadas[k, j]   # sin FactorOperacional (F-08)
            k += 1
        else:
            diario[i] = suma; i += 1; suma = 0.0
        if i >= n_dias:
            break
    acum, salida, disp_prev = 0.0, [], None
    for n, d in enumerate(diario, start=1):
        acum = d + acum                                               # E10 = D10 + E9
        den = cfg.total_racks * 24 * 60 * n / res.minutos_muestreo_derivado
        disp = 1 - acum / den if den else None
        var = 0.0 if n == 1 else (disp - disp_prev if None not in (disp, disp_prev) else None)
        salida.append(DiaDisponibilidad(n, dia=cfg.inicio_periodo + timedelta(n - 1),
                                        diario=d, acumulado=acum, disponibilidad=disp, variacion=var))
        disp_prev = disp
    return salida
```

```python
def registrar_mes_oficial(res, m, cfg) -> KpiMensual:
    # R10.2: si el período cubre el mes completo → DiasMes = días calendario
    #        si no → DiasMes = serial(última fila procesada) − serial(día 1 del mes)  (sep: 20.59375)
    #        BloquesMuestreo = C12 (intervalos existentes, D-07; sep: 1975). La fórmula de la hoja
    #        (DiasMes*24*60/15 = 1977) queda en bloques_calendario() para emular Annual_AVA.
def calcular_anual(meses: list[KpiMensual], cfg, mes_inicio_acumulado) -> list[FilaAnual]:
    # G = 1 - F/(Racks*E);  H = E + H_prev;  I = F + I_prev;  J = 1 - I/(Racks*H);  K = 0.98
```

`meses` proviene de `v_monthly_kpi_vigente` (corridas oficiales + filas `excel_manual` importadas de `Annual_AVA`).

### persistence (R11, R12)

```python
class RepositorioCorridas:
    def __init__(self, url: str)             # mssql+pyodbc, ODBC Driver 18, fast_executemany=True
    def iniciar(self, cfg) -> None           # INSERT etl_run Estado='running' (commit propio)
    def guardar_corrida(self, paquete: PaqueteCorrida) -> None   # UNA transacción: todas las tablas
    def finalizar(self, id_corrida, estado, resumen_calidad_json, mensaje=None) -> None
    def guardar_referencia_excel(self, ref: ReferenciaExcel) -> UUID
    def guardar_reconciliacion(self, reporte) -> None
```

- Inserción por lotes de 50k filas; `availability_sample_result` y `raw_pcs_sample` usan `fast_executemany` (o `BULK INSERT` desde CSV temporal si el throughput < 50k filas/s).
- `detencion` se inserta en la misma transacción; `IdTipoDetencion` se resuelve con un diccionario del catálogo cargado al inicio (sin N+1).
- El rol `etl_writer` tiene INSERT en tablas de corrida y UPDATE solo de `etl_run(Estado, FinalizadoEn, ResumenCalidad, MensajeError)`.

### reconciliation (R13)

```python
@dataclass
class Discrepancia: nivel: int; metrica: str; clave: str; python: object; excel: object; delta: float | None; tolerancia: float | None
@dataclass
class ResultadoNivel: nivel: int; nombre: str; aprobado: bool; discrepancias: list[Discrepancia]; max_delta: float | None
@dataclass
class ReporteReconciliacion: id_corrida: str; id_referencia: str | None; niveles: list[ResultadoNivel]; invariantes: list[ResultadoNivel]; aprobado_global: bool

def reconciliar(corrida: PaqueteCorrida, ref: ReferenciaExcel | None, tol: Tolerancias) -> ReporteReconciliacion
```

| Nivel | Compara | Fuente Excel |
|---|---|---|
| 1 Input | filas procesadas (=C12 y conteo de filas de la tabla E4:E), primer/último serial, nº PCS, C23, parámetros efectivos | parámetros + `Calc!E` |
| 2 Muestra | `ponderadas[k, j]` vs `Calc!F:BN` fila `k+4` (vacío = 0); `PlantActivity!D` (espejo) vs `Calc!BO` | tabla E4:BO |
| 3 Acumulados | C12 (exacto), C14 | `Calc!C12`, `C14` |
| 4 KPI | C16, C19, Daily por día (D, E, F, G) | `Calc!C16`, `C19`, `Daily!B9:G39` |
| 5 Eventos | por `OrdenExcel`: B, C, D, E, F, G, H, I; conteo; L10; resumen N:Q | `ListOfFaults` |

Invariantes (R13.8): `C14 = 4·L10` y por PCS `Σ(racks·pond)/4 = Σ I` — solo se evalúan si `inicio/fin_periodo == inicio/fin_periodo_eventos`, `aplicar_evento_excusable == aplicar_evento_excusable_eventos`, `not solo_tiempo_operacional` y no hay eventos `arrastrado_excel`; en otro caso se reportan como `no_aplica`.

### reporting

- `calidad.generar_resumen(anomalias, res_disp, res_eventos) -> dict` → `etl_run.ResumenCalidad` (JSON) + filas `data_quality_issue` (R15).
- `notificacion.enviar(resumen, destino)` → POST JSON al webhook `ETL_ARENA_NOTIFY_WEBHOOK` o `data/work/<corte>/notificacion.md` (R18). Nunca incluye credenciales ni datos crudos.

### Orquestación

- `scripts/ejecutar_etl.py --libro <xlsm> --periodo-inicio … --periodo-fin … [--eventos-inicio … --eventos-fin … --eventos-excusable Yes|No] [--diario-fin …] [--solo-operacional Yes|No] [--excusable Yes|No] [--modo-huecos excel|continuar] [--oficial] [--referencia <json>]` → ejecuta [E] y opcionalmente [R].
- `scripts/run_lunes.py --stage … ` (R19): cada etapa lee/escribe `data/work/<corte>/run_state.json` `{etapa: {estado, inicio, fin, artefactos, error}}`; `--stage all` salta etapas `ok`. Período por defecto (D-07, R19.3): `semanal` con `inicio = día 1 del mes del último dato`, `fin = fecha del último dato` (filtro `A < fin + 1` incluye hasta ese dato); si los datos ya cubren el último bloque de un mes sin `cierre_mensual` oficial, encola además `cierre_mensual` (día 1 → último día del mes). Parámetros oficiales: `C21 = "No"`, `C31 = L14 = "Yes"` (D-03), `L2/L4/Daily!D5` = `C5/C7`. El libro de trabajo de la corrida oficial pasa a ser el libro base de la siguiente. `--stage load-exclusion-matrix` (mensual, D-12, D-17) y `--reproceso <archivo>` (D-13) generan una nueva copia del libro base, escriben los cambios por timestamp → fila, registran cada celda cambiada en `correccion_dato` + `cambios.csv` y encadenan `run-macros → run-etl → reconcile → notify-bi`; tras `load-exclusion-matrix` de un mes completo se encola el `cierre_mensual` oficial.

---

## Data Models — SQL Server

> **Implementado en `sql/01…07_*.sql` (tarea 2.1, 2026-09-24), que es la fuente de verdad del esquema.** Diferencias respecto del borrador de esta sección, ya reflejadas abajo: PK de `fault_code_summary` = `(IdCorrida, Ranking)` (N:Q repite F230–F232, F-35); `fault_event` agrega `CerradoPorModulosNulo` y `TieneExclusion`; `correccion_dato.TipoCorreccion` admite `exclusion_matrix` (D-17); `raw_pcs_sample.MarcaTiempoLocalOrigen` y `availability_run_result.BloquesMuestreoCalendario` admiten NULL (A2 vacía / C23 indefinido); `excel_reference_sample` en columnstore; vista extra `v_fault_event_vigente`.

Base `ETL_Arena`, esquema `dbo`. Scripts idempotentes (`IF NOT EXISTS` / `CREATE OR ALTER`). Tipos: `FLOAT` para todo valor numérico de negocio; `UNIQUEIDENTIFIER` para `IdCorrida`; `DATETIME` para **todas** las fechas y marcas de tiempo (hora local naive; las fechas con 00:00:00); no se usan `DATE` ni `DATETIME2` (preferencia del usuario, 2026-09-24). La precisión de `DATETIME` (~3,33 ms) sobra para datos cada 15 min: la paridad usa `SerialFechaExcelOrigen` (`FLOAT`), que se guarda aparte. Valores por defecto con `GETDATE()`.

### Maestros

```sql
CREATE TABLE proyecto (
    IdProyecto          INT            NOT NULL PRIMARY KEY,
    Nombre              NVARCHAR(100)  NOT NULL,
    Estado              NVARCHAR(20)   NOT NULL CONSTRAINT CK_proyecto_estado CHECK (Estado IN ('en_ejecucion','por_implementar')),
    FechaInicio         DATETIME       NULL,
    NumPCS              INT            NULL,
    NumBateriasPorPCS   INT            NULL,
    NumRacksPorBAC      INT            NULL,
    TotalRacks          AS (NumPCS * NumBateriasPorPCS * NumRacksPorBAC) PERSISTED,
    MinutosMuestreo     INT            NULL,
    ZonaHoraria         NVARCHAR(50)   NULL,
    Descripcion         NVARCHAR(500)  NULL
);

CREATE TABLE tipo_detencion (          -- seed generado desde PCS-Fault (167 filas, F-19)
    IdTipoDetencion     INT            NOT NULL PRIMARY KEY,   -- Code Number
    CodigoFalla         NVARCHAR(10)   NOT NULL UNIQUE,        -- 'F55'
    DescripcionFallaPE  NVARCHAR(150)  NOT NULL,
    CodigoDescripcion   NVARCHAR(200)  NOT NULL,
    Significado         NVARCHAR(50)   NULL,                   -- 'Crítico', 'No fault', ...
    Operativo           NVARCHAR(10)   NULL                    -- 'Yes' / 'No' / NULL
);
```

### Corrida y staging

```sql
CREATE TABLE etl_run (
    IdCorrida                      UNIQUEIDENTIFIER NOT NULL PRIMARY KEY,
    IdProyecto                     INT            NOT NULL REFERENCES proyecto(IdProyecto),
    TipoCorrida                    NVARCHAR(20)   NOT NULL CHECK (TipoCorrida IN ('semanal','cierre_mensual','reproceso','golden')),
    EsOficial                      BIT            NOT NULL DEFAULT 0,
    EstadoExclusiones              NVARCHAR(20)   NOT NULL CHECK (EstadoExclusiones IN ('sin_exclusiones','con_exclusiones')),  -- D-17, R19.5: ambos oficiales
    ArchivoOrigen                  NVARCHAR(500)  NOT NULL,
    HashArchivoOrigen              CHAR(64)       NOT NULL,
    SistemaOrigen                  NVARCHAR(50)   NOT NULL DEFAULT 'scada_export',
    IdReferenciaExcel              UNIQUEIDENTIFIER NULL,
    IniciadoEn                     DATETIME       NOT NULL,
    FinalizadoEn                   DATETIME       NULL,
    InicioPeriodo                  DATETIME       NOT NULL,   -- C5
    FinPeriodo                     DATETIME       NOT NULL,   -- C7
    InicioPeriodoEventos           DATETIME       NOT NULL,   -- L2
    FinPeriodoEventos              DATETIME       NOT NULL,   -- L4
    FinDiario                      DATETIME       NOT NULL,   -- Daily!D5
    SoloTiempoOperacional          BIT            NOT NULL,   -- C21
    AplicarEventoExcusable         BIT            NOT NULL,   -- C31
    AplicarEventoExcusableEventos  BIT            NOT NULL,   -- L14
    ModoHuecos                     NVARCHAR(10)   NOT NULL,
    TotalPCS                       INT            NOT NULL,
    BateriasPorPCS                 INT            NOT NULL,
    RacksPorPCS                    INT            NOT NULL,
    TotalRacks                     INT            NOT NULL,
    MinutosMuestreo                INT            NOT NULL,
    MinutosMuestreoDerivado        FLOAT          NULL,       -- C23
    Estado                         NVARCHAR(20)   NOT NULL CHECK (Estado IN ('running','success','failed','parity_failed')),
    MensajeError                   NVARCHAR(MAX)  NULL,
    ResumenCalidad                 NVARCHAR(MAX)  NULL CHECK (ResumenCalidad IS NULL OR ISJSON(ResumenCalidad) = 1),
    VersionAlgoritmo               NVARCHAR(100)  NOT NULL
);

CREATE TABLE raw_pcs_column_map (      -- trazabilidad de columnas una vez por corrida
    IdCorrida       UNIQUEIDENTIFIER NOT NULL REFERENCES etl_run(IdCorrida),
    NumeroPCS       INT            NOT NULL,
    Campo           NVARCHAR(20)   NOT NULL,     -- falla | estado | advertencia | modulos
    NumeroColumna   INT            NOT NULL,
    NombreColumna   NVARCHAR(200)  NOT NULL,
    CONSTRAINT PK_raw_pcs_column_map PRIMARY KEY (IdCorrida, NumeroPCS, Campo)
);

CREATE TABLE raw_pcs_sample (          -- filas del período calculado + la fila anterior (fallback F-04); D-09
    IdCorrida               UNIQUEIDENTIFIER NOT NULL,
    NumeroFilaOrigen        INT            NOT NULL,
    NumeroPCS               INT            NOT NULL,
    SerialFechaExcelOrigen  FLOAT          NOT NULL,
    MarcaTiempoLocalOrigen  DATETIME       NOT NULL,
    FallaRaw                NVARCHAR(255)  NULL,
    FallaRawEsNumero        BIT            NOT NULL,
    EstadoRaw               NVARCHAR(255)  NULL,
    AdvertenciaRaw          NVARCHAR(255)  NULL,
    ModulosRaw              FLOAT          NULL,
    ModulosDisponibles      FLOAT          NOT NULL,   -- = baterias_por_pcs si nulo
    ModulosDisponiblesNulo  BIT            NOT NULL,
    EsFilaNuevaDelExport    BIT            NOT NULL    -- aportada por el export incremental de esta corrida (D-07)
);
CREATE CLUSTERED COLUMNSTORE INDEX CCI_raw_pcs_sample ON raw_pcs_sample;

CREATE TABLE plant_activity_sample (
    IdCorrida                       UNIQUEIDENTIFIER NOT NULL,
    NumeroFilaOrigen                INT            NOT NULL,
    SerialFecha                     FLOAT          NULL,      -- puede faltar (F-31)
    MarcaTiempoMuestra              DATETIME       NULL,
    FactorOperacionalRaw            FLOAT          NULL,      -- col C
    EventoExcusadoRaw               FLOAT          NULL,      -- col D: espejo de Calc!BO, no pondera (F-37)
    SetpointPotenciaActivaKW        FLOAT          NULL,
    DroopSobrefrecuenciaHabilitado  FLOAT          NULL,
    DroopBajafrecuenciaHabilitado   FLOAT          NULL,
    PotenciaActivaPOIKW             FLOAT          NULL,
    PorcentajeSOC                   FLOAT          NULL
);
CREATE CLUSTERED COLUMNSTORE INDEX CCI_plant_activity_sample ON plant_activity_sample;

CREATE TABLE exclusion_matrix_carga (   -- D-17: una fila por entrega mensual de Alex (append-only; vigente = última)
    IdCarga                 BIGINT IDENTITY PRIMARY KEY,
    IdProyecto              INT            NOT NULL REFERENCES proyecto(IdProyecto),
    Anio                    INT            NOT NULL,
    Mes                     INT            NOT NULL,
    ArchivoOrigen           NVARCHAR(500)  NOT NULL,
    Sha256Archivo           CHAR(64)       NOT NULL,
    CargadoEn               DATETIME       NOT NULL,
    IdCorridaCierre         UNIQUEIDENTIFIER NULL     -- cierre_mensual con_exclusiones que la usó
);

CREATE TABLE exclusion_matrix_sample (  -- F-37: una fila por (fila, PCS) con valor ≠ 0, del período
    IdCorrida               UNIQUEIDENTIFIER NOT NULL,
    NumeroFilaOrigen        INT            NOT NULL,
    NumeroPCS               INT            NOT NULL,
    SerialFecha             FLOAT          NULL,
    MarcaTiempoMuestra      DATETIME       NULL,
    ValorExclusion          TINYINT        NOT NULL CHECK (ValorExclusion IN (1, 2)),
    BateriasPrevias         FLOAT          NULL,      -- solo valor 2
    EventoExcusadoFila      FLOAT          NULL,      -- columna "Excused Event"
    Comentario              NVARCHAR(500)  NULL       -- columna "Comments" (causa del EE)
);
CREATE CLUSTERED COLUMNSTORE INDEX CCI_exclusion_matrix_sample ON exclusion_matrix_sample;

CREATE TABLE correccion_dato (          -- append-only; celdas cambiadas por reproceso o carga de PlantActivity (D-12, D-13)
    IdCorreccion            BIGINT IDENTITY PRIMARY KEY,
    IdCorrida               UNIQUEIDENTIFIER NOT NULL REFERENCES etl_run(IdCorrida),
    TipoCorreccion          NVARCHAR(20)   NOT NULL CHECK (TipoCorreccion IN ('reproceso_raw','exclusion_matrix','plant_activity')),
    Hoja                    NVARCHAR(40)   NOT NULL,   -- 'RawData-PCS' | 'PlantActivity'
    NumeroFilaOrigen        INT            NOT NULL,
    SerialFechaExcelOrigen  FLOAT          NOT NULL,
    MarcaTiempoLocalOrigen  DATETIME       NOT NULL,
    NumeroPCS               INT            NULL,       -- NULL en PlantActivity
    Campo                   NVARCHAR(40)   NOT NULL,   -- FAULT | STATUS | WARNING | MODULES | B | C | D | …
    ValorAnterior           NVARCHAR(255)  NULL,
    ValorNuevo              NVARCHAR(255)  NULL,
    ArchivoOrigen           NVARCHAR(260)  NOT NULL,
    Sha256Archivo           CHAR(64)       NOT NULL
);
```

`v_correccion_dato` expone los cambios con `TipoCorrida`, `FechaEjecucion` y el KPI antes/después (corrida anterior del mismo período vs la corrida del reproceso).

### Resultados

```sql
CREATE TABLE availability_sample_result (   -- solo filas dentro del período KPI
    IdCorrida                        UNIQUEIDENTIFIER NOT NULL,
    NumeroFilaOrigen                 INT     NOT NULL,
    NumeroPCS                        INT     NOT NULL,
    SerialFecha                      FLOAT   NOT NULL,
    MarcaTiempoMuestra               DATETIME     NOT NULL,
    ModulosDisponibles               FLOAT   NOT NULL,
    ModulosDisponiblesNulo           BIT     NOT NULL,
    BateriasIndisponibles            FLOAT   NOT NULL,
    ValorExclusion                   TINYINT NOT NULL,   -- Exclusion_Matrix 0/1/2 (F-37)
    FactorOperacional                FLOAT   NOT NULL,
    BateriasIndisponiblesPonderadas  FLOAT   NOT NULL,
    ImpactoRackPonderado             FLOAT   NOT NULL
);
CREATE CLUSTERED COLUMNSTORE INDEX CCI_availability_sample_result ON availability_sample_result;

CREATE TABLE availability_run_result (
    IdCorrida                     UNIQUEIDENTIFIER NOT NULL PRIMARY KEY REFERENCES etl_run(IdCorrida),
    BloquesMuestreo               INT    NOT NULL,   -- C12
    TotalRacks                    INT    NOT NULL,   -- C11
    BloquesRacksIndisponibles     FLOAT  NOT NULL,   -- C14
    DisponibilidadPeriodo         FLOAT  NULL,       -- C16
    DisponibilidadAnualAcumulada  FLOAT  NULL,       -- C19
    MinutosMuestreoDerivado       FLOAT  NULL,       -- C23
    HorasRackEventos              FLOAT  NOT NULL,   -- ListOfFaults!L10
    BloquesMuestreoCalendario     FLOAT  NOT NULL    -- (C7-C5+1)*24*60/C23 (B12, informativo)
);

CREATE TABLE fault_event (
    IdCorrida                     UNIQUEIDENTIFIER NOT NULL REFERENCES etl_run(IdCorrida),
    OrdenExcel                    INT            NOT NULL,     -- fila dblResult - 5
    NumeroPCS                     INT            NOT NULL,
    SerialInicio                  FLOAT          NOT NULL,
    SerialFin                     FLOAT          NOT NULL,
    MarcaTiempoInicio             DATETIME       NOT NULL,
    MarcaTiempoFin                DATETIME       NOT NULL,
    DuracionHoras                 FLOAT          NOT NULL,
    CodigoFalla                   NVARCHAR(300)  NOT NULL,     -- "F" & descripción puede ser largo
    DescripcionFalla              NVARCHAR(255)  NOT NULL,
    DescripcionFallaFallback      BIT            NOT NULL,
    NumeroBloques                 INT            NOT NULL,
    SumaBloques                   FLOAT          NOT NULL,
    PromedioBateriasInvolucradas  FLOAT          NOT NULL,
    HorasRackIndisponibles        FLOAT          NOT NULL,
    EventoArrastradoExcel         BIT            NOT NULL,
    ExcelHabriaFallado            BIT            NOT NULL,
    CerradoPorModulosNulo         BIT            NOT NULL,
    TieneExclusion                BIT            NOT NULL,     -- algún bloque con Exclusion_Matrix ≠ 0 (F-37)
    CONSTRAINT PK_fault_event PRIMARY KEY (IdCorrida, OrdenExcel)
);

CREATE TABLE fault_code_summary (      -- ListOfFaults!N:Q (166 filas: F230-F232 repetidos, F-35)
    IdCorrida              UNIQUEIDENTIFIER NOT NULL REFERENCES etl_run(IdCorrida),
    Ranking                INT            NOT NULL,
    CodigoFalla            NVARCHAR(10)   NOT NULL,
    DescripcionFallaPE     NVARCHAR(150)  NULL,
    HorasRackIndisponibles FLOAT          NOT NULL,
    Porcentaje             FLOAT          NULL,
    CONSTRAINT PK_fault_code_summary PRIMARY KEY (IdCorrida, Ranking)
);

CREATE TABLE daily_availability (
    IdCorrida                            UNIQUEIDENTIFIER NOT NULL REFERENCES etl_run(IdCorrida),
    DiaN                                 INT    NOT NULL,
    Dia                                  DATETIME NOT NULL,
    BloquesRacksIndisponiblesDiarios     FLOAT  NOT NULL,
    BloquesRacksIndisponiblesAcumulados  FLOAT  NOT NULL,
    Disponibilidad                       FLOAT  NULL,
    Variacion                            FLOAT  NULL,
    CONSTRAINT PK_daily_availability PRIMARY KEY (IdCorrida, DiaN)
);

CREATE TABLE monthly_official_kpi (    -- append-only; vigente = último RegistradoEn
    IdKpiMensual               BIGINT IDENTITY PRIMARY KEY,
    IdProyecto                 INT            NOT NULL REFERENCES proyecto(IdProyecto),
    Anio                       INT            NOT NULL,
    Mes                        INT            NOT NULL CHECK (Mes BETWEEN 1 AND 12),
    DiasMes                    FLOAT          NOT NULL,
    BloquesMuestreo            FLOAT          NOT NULL,
    BloquesRacksIndisponibles  FLOAT          NOT NULL,
    DisponibilidadContractual  FLOAT          NOT NULL DEFAULT 0.98,
    Origen                     NVARCHAR(20)   NOT NULL CHECK (Origen IN ('corrida','excel_manual')),
    IdCorrida                  UNIQUEIDENTIFIER NULL REFERENCES etl_run(IdCorrida),
    RegistradoEn               DATETIME       NOT NULL DEFAULT GETDATE()
);

CREATE TABLE annual_availability (     -- snapshot calculado en cada corrida
    IdCorrida                       UNIQUEIDENTIFIER NOT NULL REFERENCES etl_run(IdCorrida),
    Anio INT NOT NULL, Mes INT NOT NULL,
    DiasMes                         FLOAT NOT NULL,
    BloquesMuestreo                 FLOAT NOT NULL,
    BloquesRacksIndisponibles       FLOAT NOT NULL,
    DisponibilidadMensual           FLOAT NULL,
    BloquesMuestreoAcumulados       FLOAT NOT NULL,
    BloquesIndisponiblesAcumulados  FLOAT NOT NULL,
    DisponibilidadAcumulada         FLOAT NULL,
    DisponibilidadContractual       FLOAT NULL,
    Origen                          NVARCHAR(20) NOT NULL,
    CONSTRAINT PK_annual_availability PRIMARY KEY (IdCorrida, Anio, Mes)
);
```

### Detenciones (R12)

```sql
CREATE TABLE detencion (               -- append-only por corrida
    IdDetencion                   BIGINT IDENTITY PRIMARY KEY,
    IdProyecto                    INT            NOT NULL REFERENCES proyecto(IdProyecto),
    IdCorrida                     UNIQUEIDENTIFIER NOT NULL REFERENCES etl_run(IdCorrida),
    OrdenExcel                    INT            NOT NULL,
    NumeroPCS                     INT            NOT NULL,
    FechaInicio                   DATETIME       NOT NULL,
    FechaTermino                  DATETIME       NOT NULL,
    DuracionSegundos              INT            NOT NULL,   -- ROUND(DuracionHoras*3600)
    IdTipoDetencion               INT            NULL REFERENCES tipo_detencion(IdTipoDetencion),
    CodigoFalla                   NVARCHAR(300)  NOT NULL,
    DescripcionFalla              NVARCHAR(255)  NOT NULL,
    DescripcionFallaFallback      BIT            NOT NULL,
    PromedioBateriasInvolucradas  FLOAT          NOT NULL,
    HorasRackIndisponibles        FLOAT          NOT NULL,
    CerradoPorModulosNulo         BIT            NOT NULL,   -- el evento terminó porque la fila siguiente estaba vacía
    EsExcusable                   BIT            NOT NULL,   -- algún bloque con Exclusion_Matrix ≠ 0 (F-37)
    EventoArrastradoExcel         BIT            NOT NULL
);
CREATE INDEX IX_detencion_negocio ON detencion(IdProyecto, NumeroPCS, FechaInicio);
CREATE INDEX IX_detencion_corrida ON detencion(IdCorrida);

CREATE TABLE detencion_revision (      -- workflow; historial append-only
    IdRevision       BIGINT IDENTITY PRIMARY KEY,
    IdProyecto       INT            NOT NULL REFERENCES proyecto(IdProyecto),
    NumeroPCS        INT            NOT NULL,
    FechaInicio      DATETIME       NOT NULL,
    EstadoRevision   NVARCHAR(20)   NOT NULL CHECK (EstadoRevision IN ('pendiente','revisado','excluido')),
    Observacion      NVARCHAR(500)  NULL,
    RevisadoPor      NVARCHAR(100)  NOT NULL,
    RevisadoEn       DATETIME       NOT NULL DEFAULT GETDATE()
);
CREATE INDEX IX_detencion_revision_clave ON detencion_revision(IdProyecto, NumeroPCS, FechaInicio, RevisadoEn DESC);
```

### Calidad, referencia Excel y reconciliación

```sql
CREATE TABLE data_quality_issue (
    IdCorrida        UNIQUEIDENTIFIER NOT NULL REFERENCES etl_run(IdCorrida),
    Tipo             NVARCHAR(50)   NOT NULL,
    Severidad        NVARCHAR(12)   NOT NULL CHECK (Severidad IN ('info','advertencia','error')),
    NumeroFilaOrigen INT            NULL,
    NumeroPCS        INT            NULL,
    SerialFecha      FLOAT          NULL,
    Detalle          NVARCHAR(1000) NULL
);
CREATE INDEX IX_dqi_corrida_tipo ON data_quality_issue(IdCorrida, Tipo);

CREATE TABLE excel_reference_run (
    IdReferencia      UNIQUEIDENTIFIER NOT NULL PRIMARY KEY,
    Corte             NVARCHAR(50)   NOT NULL,
    ArchivoLibro      NVARCHAR(500)  NOT NULL,
    HashLibro         CHAR(64)       NOT NULL,
    ExtraidoEn        DATETIME       NOT NULL,
    C5 DATETIME NOT NULL, C7 DATETIME NOT NULL, C21 NVARCHAR(10) NULL, C31 NVARCHAR(10) NULL,
    L2 DATETIME NOT NULL, L4 DATETIME NOT NULL, L14 NVARCHAR(10) NULL, DailyD5 DATETIME NULL,
    C12 INT NULL, C14 FLOAT NULL, C16 FLOAT NULL, C19 FLOAT NULL, C23 FLOAT NULL, L10 FLOAT NULL
);
CREATE TABLE excel_reference_sample (  -- Calc!E4:BO, solo celdas no vacías
    IdReferencia UNIQUEIDENTIFIER NOT NULL REFERENCES excel_reference_run(IdReferencia),
    FilaResultado INT NOT NULL, SerialFecha FLOAT NOT NULL, NumeroPCS INT NOT NULL, Valor FLOAT NOT NULL
);
CREATE TABLE excel_reference_fault_event (
    IdReferencia UNIQUEIDENTIFIER NOT NULL REFERENCES excel_reference_run(IdReferencia),
    OrdenExcel INT NOT NULL, B FLOAT NULL, C FLOAT NULL, D FLOAT NULL, E FLOAT NULL,
    F NVARCHAR(300) NULL, G NVARCHAR(255) NULL, H FLOAT NULL, I FLOAT NULL,
    CONSTRAINT PK_excel_reference_fault_event PRIMARY KEY (IdReferencia, OrdenExcel)
);
CREATE TABLE excel_reference_daily (
    IdReferencia UNIQUEIDENTIFIER NOT NULL REFERENCES excel_reference_run(IdReferencia),
    DiaN INT NOT NULL, Dia DATETIME NULL, D FLOAT NULL, E FLOAT NULL, F FLOAT NULL, G FLOAT NULL,
    CONSTRAINT PK_excel_reference_daily PRIMARY KEY (IdReferencia, DiaN)
);

CREATE TABLE reconciliation_result (
    IdCorrida     UNIQUEIDENTIFIER NOT NULL REFERENCES etl_run(IdCorrida),
    IdReferencia  UNIQUEIDENTIFIER NULL REFERENCES excel_reference_run(IdReferencia),
    Nivel         INT            NOT NULL,          -- 1..5; 9 = invariante
    Metrica       NVARCHAR(100)  NOT NULL,
    Clave         NVARCHAR(200)  NULL,
    ValorPython   NVARCHAR(100)  NULL,
    ValorExcel    NVARCHAR(100)  NULL,
    Delta         FLOAT          NULL,
    Tolerancia    FLOAT          NULL,
    Aprobado      BIT            NOT NULL
);
```

### Vistas para Power BI y auditoría (R18.4)

- `v_corrida_oficial_vigente`: por `(IdProyecto, año-mes de FinPeriodo)` la última `etl_run` con `EsOficial=1` y `Estado='success'`, prefiriendo `EstadoExclusiones='con_exclusiones'` sobre `sin_exclusiones` (R19.7). Expone la etiqueta "Sin Exclusiones" / "Con Exclusiones" para Power BI.
- `v_kpi_vigente`, `v_daily_vigente`, `v_annual_vigente`, `v_fault_code_vigente`: resultados de la corrida vigente.
- `v_detencion_vigente`: `detencion` de la corrida vigente `LEFT JOIN` a la última `detencion_revision` por clave (estado por defecto `pendiente`).
- `v_monthly_kpi_vigente`: última fila por `(IdProyecto, Anio, Mes)` de `monthly_official_kpi`.
- `v_calidad_corrida`, `v_modulos_nulos_historico` (R15.4).
- `07_audit_queries.sql`: trazabilidad KPI → muestra → raw → columna origen; eventos con fallback; eventos arrastrados.

Seguridad: rol `etl_writer` (INSERT en tablas de corrida, UPDATE acotado en `etl_run`), rol `revisor` (INSERT en `detencion_revision`), rol `bi_reader` (SELECT solo sobre vistas `v_*`). `DENY DELETE` a `etl_writer`.

---

## Tolerancias (tabla única — R13.7)

| Métrica | Tolerancia gate | Objetivo esperado |
|---|---|---|
| C12, filas, nº eventos, nº PCS, DiaN | exacto | exacto |
| Nivel 2 ponderadas por celda | abs 1e-12 | 0 |
| C14 | abs 1e-6 | ≤ 1e-9 |
| C16, C19, disponibilidad y variación diaria | abs 1e-9 | ≤ 1e-12 |
| Daily diario / acumulado | abs 1e-6 | ≤ 1e-9 |
| Evento: seriales inicio/fin | abs 1e-9 días | 0 |
| Evento: DuracionHoras, Promedio | abs 1e-9 | ≤ 1e-12 |
| Evento: HorasRackIndisponibles | abs 1e-6 | ≤ 1e-9 |
| L10 | abs 1e-6 | ≤ 1e-6 |
| Código/descripción/PCS | exacto | exacto |
| Golden `_meta` (histórico) | `tolerance_kpi` 1e-6, `tolerance_rack_hours` 1e-3 | se endurece a esta tabla al verificar |

Las tolerancias viven en `reconciliation/tolerancias.py` y en `golden_index.json → _meta`; un test verifica que coinciden.

---

## Correctness Properties

Cada property test lleva el tag `# Feature: etl-arena-availability, Property N` y ≥ 200 ejemplos Hypothesis.

1. **Forma de la matriz** — para N filas × P PCS generados, `MatrizPCS` tiene forma (N, P); cada celda vacía ⇒ `modulos_nulo` y `ModulosDisponibles = baterias_por_pcs`; cada valor numérico (incl. fraccionario) se conserva bit a bit. *R5.1, R5.3, R5.4*
2. **Anomalías exhaustivas** — anomalías inyectadas (hueco, duplicado, fuera de orden, frecuencia distinta, A vacía) se reportan todas, sin falsos positivos. *R4.4, R4.5*
3. **Impacto por celda** — para toda combinación de `(m ∈ [0,4] ∪ {vacío}, fe, fo ∈ [0,1], flags)` el impacto iguala la fórmula de R7.4–R7.6. *R7*
4. **C12 cuenta filas** — con timestamps duplicados inyectados, `C12` = filas en rango. *R7.2*
5. **C14 = Σ impactos** en el orden fila→PCS (igualdad exacta de floats). *R7.7, R7.10*
6. **Cobertura de eventos** — sin arrastre ni `ExcelHabriaFallado`: `Σ NumeroBloques` = filas en falla en rango, y cada evento cubre una racha maximal separada por `=4`/vacío/límites. *R8.2–R8.6*
7. **Árbol de descripción** — para `(actual, anterior)` en {`"NO FAULTS"`, `""`, `None`, texto, número}: resultado y flag según R8.8 (anterior vacío ⇒ `"F1 Watchdog"` con fallback). *R8.8*
8. **Código de falla** — `codigo_falla_excel` coincide con una implementación de referencia literal de `FIND/MID/IFERROR/CONCATENATE`. *R8.9, R17*
9. **Daily** — `Σ diario = Σ racks·ponderadas` en rango de días; `acumulado` monótono; fórmula exacta de disponibilidad; `solo_tiempo_operacional` no afecta al diario. *R9*
10. **Invariante cruzado** — bajo las condiciones de R13.8, `|C14 − 4·L10| ≤ 1e-6`. *R13.8*
11. **Arrastre** — si la fila siguiente a la última en rango es exactamente `L4+1` y sigue en falla, el siguiente PCS con falla produce un registro `arrastrado_excel` con bloques heredados. *R8.7*

---

## Error Handling

| Situación | Comportamiento |
|---|---|
| Archivo/hoja inexistente | `failed`, `MensajeError` descriptivo |
| Módulos con texto (modo paridad) | `failed` (el VBA fallaría) con fila/PCS |
| Menos de 2 filas en el período | `failed` (C23 indefinido) |
| A vacía intermedia | `modo_huecos=excel`: truncar + `advertencia`; `continuar`: omitir filas |
| PlantActivity vacía/desalineada | factor 0 + `advertencia` (la corrida sigue) |
| Código sin catálogo | `IdTipoDetencion NULL` + `advertencia` |
| Error SQL en `guardar_corrida` | rollback total; `finalizar(failed)` en transacción aparte; relanzar |
| Macro COM falla / timeout | matar solo la instancia propia; referencia ausente; el ETL puede seguir (`sin_referencia`) |
| División por cero en KPI | `None` (equivalente `IFERROR`) |
| Reconciliación niveles 3–5 fuera de tolerancia | `parity_failed` + notificación con resumen |

No se redondea ningún intermedio. No se usan `print` para errores: logging estructurado (`logging` + JSON formatter) con `IdCorrida` en cada registro.

---

## Testing Strategy

| Capa | Ubicación | Corre en | Notas |
|---|---|---|---|
| Unit | `tests/unit/` | CI + local | `excel_semantics` con casos del libro real; motores con matrices pequeñas escritas a mano |
| Property | `tests/property/` | CI + local | Hypothesis, Properties 1–11 |
| Golden | `tests/golden/` | local (requiere `.xlsm`, `ETL_ARENA_XLSM`) | integridad de JSON (siempre) + runner sobre meses `verified` |
| Integración SQL | `tests/integration/` (marker `sql`) | local con Docker | DDL, append-only, rollback, vistas |
| COM | `tests/com/` (marker `excel`) | solo PC local Windows con Excel | smoke de macros sobre copia en `tmp` |

Fixtures: `tests/fixtures/` contiene libros sintéticos generados por código (sin datos reales) y un recorte CSV del libro real (48 h de septiembre, 5 PCS) con sus valores Excel esperados obtenidos vía COM.

Cobertura mínima: 90 % en `availability`, `fault_events`, `aggregation`, `excel_semantics`; 80 % global.

---

## Golden Tests — valores reales del Excel

> Libro: `AvailabilityCalculation_PCS&Batteries_20260907_septiembre 2026.xlsm` (estado al 2026-09-24). Parámetros actuales: `C5=2026-09-01`, `C7=2026-09-21`, `C21="No"`, `C31="No"`, `L2=2026-09-01`, `L4=2026-09-21`, **`L14="Yes"`**, `Daily!D5=2026-09-21`, `C23=15`.

### GT-1: Encabezados de RawData-PCS

`Arena - PCS XX - POWERELECTRONICS GEN3 HEx CURRENT FAULT | … CURRENT STATUS | … CURRENT WARNING | … HEM-k NUMBER OF MODULES`, XX = 01…61, en columnas B…IK en ese orden (verificado).

### GT-2: Rango y forma de datos

Primer dato `2026-04-08 00:15` (fila 2); último `2026-09-21 14:15` (fila 15991); 15.990 filas; sin duplicados; un salto de 60 min en `2026-09-06 00:00 → 01:00` (DST). La fila `00:00` pertenece al día que indica su fecha (`≥ C5`).

### GT-3: KPI junio 2026 (histórico, **no verificado**)

`C12=2880`, `C14=1211931.4879999976`, `C16=0.8562808932908321`, `C19=0.9881874706814383` (extracción 2026-09-22 de un estado anterior del libro). No está en `golden_index.json`; re-extraer vía COM (tarea 5.6) antes de usarlo como gate.

### GT-4: KPI septiembre 2026 (1–21) — gate

| Celda | Valor |
|---|---|
| C12 | 1975 (= 20·96 + 58 − 3 por DST) |
| C14 | 104134.2920000001 |
| C16 | 0.98199240990523617 |
| C19 | 0.99898501739619983 |
| L10 | 26033.573000032695 (334 eventos) |
| Invariante | `C14 = 4·L10` ✔ |

Daily (primeros días; serie completa en el golden): `D9=4245.128`, `F9=0.98489751252276869`; `D10=4226.76`, `E10=8471.888`, `F10=0.98493018556466305`, `G10=3.2673041894359933e-5`; día 21: `E=104134.292`, `F=0.9823586356958539`.

### GT-5: Annual_AVA (informativo, no gate)

Filas jul (`F13=350972.3`) y ago (`F14=305182.968`) son valores manuales que no coinciden con los C14 de sus corridas; sep `F15 = C14`, `D15 = 20+14.25/24` (manual), `E15 = 1977`. `H/I/J` acumulan desde julio. Se importan como `excel_manual` en `monthly_official_kpi`.

### GT-6: ListOfFaults — primeros eventos PCS 1 (orden de escritura)

| Orden | PCS | Inicio | Fin | E (h) | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| 1 | 1 | 46266.010416666664 (09-01 00:15) | 46266.239583333336 (05:45) | 5.5000000001164153 | F55 | F55 EXTERNAL FAULT/OVGR | 3.3181818181818183 | 219.00000000463547 |
| 2 | 1 | 46266.34375 (08:15) | 46266.354166666664 (08:30) | 0.24999999994179234 | F13 | F13 NO MODULES | 0.33333333333332993 | 0.99999999976715914 |

Los valores exactos (con ruido de serial) son el esperado: el motor opera en seriales (ADR-03). La lista no está ordenada por impacto; `mcoOrder` ordena solo `N5:Q172`.

### GT-7: Módulos fraccionarios

Fila 2, PCS 04: `3.23466666666667` → `BateriasIndisponibles = 0.76533333333333…`, impacto `12 × 0.7653… = 9.184`. Valores fraccionarios: 15.655 celdas; se conservan como `float` (resuelto, F-02).

### GT-8: PlantActivity desalineada

Filas 2–15332 alineadas por timestamp con RawData; filas 15333–15991 sin timestamp en `PlantActivity!B` (C=0/D=1 presentes). Con join por fila el Excel usa esos valores; el reporte de calidad debe listar 659 filas `pa_sin_timestamp`.

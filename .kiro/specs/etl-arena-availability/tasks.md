# Implementation Plan: ETL Arena Availability

> Revisión 2 — 2026-09-24. Reemplaza la revisión 1 tras la auditoría (`audit.md`). Requisitos `R<n>` en `requirements.md`; diseño en `design.md`.

## Overview

Implementación en **6 fases** y **olas** paralelizables. Orden: preparar entorno y goldens → núcleo de paridad en Python puro (sin SQL ni Excel) → SQL Server → pipeline + reconciliación → cadena semanal SCADA/COM → documentación y shadow mode.

Stack: Python 3.14 (`>= 3.13`), numpy, pandas ≥ 2.2 (solo persistencia), openpyxl (goldens/seed), SQLAlchemy 2 + pyodbc 5 (ODBC Driver 18), pywin32 (COM), pytest, pytest-cov, Hypothesis, ruff. Sin `pytz` (ADR-07).

---

## Agentes y protocolo

Agentes disponibles en `.opencode/agent/` (OpenCode, `mode: subagent`). Son perfiles genéricos: **cada encargo debe incluir el paquete de contexto** indicado abajo.

| Agente | Rol en este proyecto | Tareas |
|---|---|---|
| `software-architect` | ADRs, decisiones D-xx, límites de módulos, go/no-go | 0.1, 5.5 |
| `data-engineer` | Motores de paridad, ingesta, normalización, agregaciones, reconciliación, goldens, calidad | 0.4, 0.5, 1.x, 3.2, 3.3, 4.4, 4.6, 5.4 |
| `backend-architect` | Entorno, persistencia, pipeline/CLI, COM Excel, adquisición, orquestador, notificación | 0.3, 2.4, 2.5, 3.1, 3.4, 4.1–4.3, 4.5, 4.7–4.10 |
| `database-optimizer` | DDL, seeds, vistas, roles, rendimiento | 2.1–2.3, 2.6 |
| `code-reviewer` | Gate de revisión al cierre de cada fase (checklist de paridad) | Checkpoints 0, A, B, C, D |
| `technical-writer` | Data contracts, runbook, README/AGENTS/CLAUDE, handoff PBI | 0.5, 0.7, 5.1–5.3 |
| Humano (Francisco / Alex / negocio / Misael) | Decisiones D-03/04/06/07, muestra P0, corridas COM reales, shadow mode | 0.2, 0.6, 4.10, 5.4 |

**Paquete de contexto obligatorio** (incluir en cada prompt a un agente):
1. `CLAUDE.md` (reglas para agentes) y las secciones de `AGENTS.md` que toque la tarea.
2. Los requisitos `R<n>` y la sección de `design.md` citados en la tarea.
3. Los hallazgos `F-xx` relevantes de `audit.md`.
4. Reglas duras: no modificar el algoritmo; no hardcodear 61/4/12/15; seriales Excel en los motores; append-only; nunca editar `data/AvailabilityCalculation_*.xlsm` ni `data/processed/`.

**Notas por agente:**
- `database-optimizer` está orientado a PostgreSQL: exigir **T-SQL para SQL Server 2022** (columnstore, `CREATE OR ALTER`, `ISJSON`, `DATETIME2`), sin sintaxis PG.
- `backend-architect` para COM: exigir `pywin32` con `DispatchEx`, instancia visible, cierre solo de la instancia propia y watchdog de timeout.
- Perfiles adicionales: `parity-qa` (tests Hypothesis, goldens, reconciliación; puede co-ejecutar 1.2, 1.8, 1.11, 3.3) y `excel-com-automation` (puede co-ejecutar 4.1–4.3 y 4.7).
- Portados a `.claude/agents/` (2026-09-24) con el paquete de contexto incluido, junto a `parity-qa` y `excel-com-automation`; `.opencode/agent/` se conserva para OpenCode.

**Checklist de revisión de paridad (`code-reviewer`, en cada checkpoint):**
- [ ] Los motores recorren en **orden de fila origen** y no deduplican ni reordenan.
- [ ] Filtros y restas de tiempo usan `SerialFechaExcelOrigen`, no `datetime`.
- [ ] Toda comparación de celda pasa por `excel_semantics`.
- [ ] PlantActivity se une **por fila**; vacío = 0.
- [ ] Parámetros de eventos (L2/L4/L14) separados de los KPI (C5/C7/C31).
- [ ] Sin constantes de negocio en los motores (grep de `61`, `2928`, `12`, `15` fuera de `config/defaults.py` y tests).
- [ ] Ningún intermedio redondeado; `redondear_excel` solo donde Excel usa `ROUND`.
- [ ] Sin `DELETE`/`UPDATE` fuera de `etl_run` (estado) y `detencion_revision` (insert).
- [ ] Tests nuevos citan `Rn.m` o `Property N`.

---

## Tasks

### Fase 0 — Preparación

- [ ] 0.1 ADRs y registro de decisiones — **software-architect**
  - Crear `docs/adr/ADR-01…ADR-10.md` a partir de la tabla de decisiones de `design.md` y `docs/adr/decisiones-abiertas.md` con D-01…D-10 (default, dueño, estado).
  - Hecho cuando: cada ADR tiene contexto, opciones, decisión, consecuencias; D-03/04/06/07 marcadas "pendiente negocio".
  - _R1, R10, R11, R19_

- [ ] 0.2 Resolver decisiones de negocio D-03, D-04, D-06, D-07 — **Humano (Francisco + responsable KPI)**
  - D-03: ¿`L14` debe igualar a `C31`? (hoy `L14="Yes"`, `C31="No"`).
  - D-04: ¿replicar el arrastre de eventos en v1? (default: sí, con flag).
  - D-06: inicio de acumulación anual (julio como el libro o backfill desde 08-abr).
  - D-07: período de la corrida semanal y cierre mensual.
  - Hecho cuando: `decisiones-abiertas.md` actualizado con la resolución y fecha. **No bloquea la Fase 1** (los defaults de paridad permiten avanzar).

- [x] 0.3 Entorno de desarrollo — **backend-architect** *(hecho 2026-09-24; pendiente instalar ODBC Driver 18, requiere admin)*
  - Python 3.14 (todas las dependencias con wheels); crear `.venv`; `pyproject.toml` (paquete `etl_arena`, layout `src/`, extras `dev`, `sql`, `com`), config de `pytest` (markers `sql`, `excel`, `golden`), `ruff`, `coverage`.
  - `.gitignore`: `.venv/`, `data/inbox/*`, `data/processed/*`, `data/work/*` (con `.gitkeep`), `*.pyc`, `__pycache__/` (hay un `.pyc` versionado en `tests/golden/__pycache__/`: quitarlo del índice).
  - `.env.example` (`ETL_ARENA_DB_URL`, `ETL_ARENA_XLSM`, `ETL_ARENA_NOTIFY_WEBHOOK`); `docker/mssql.compose.yml` (SQL Server 2022, contraseña por variable de entorno).
  - Documentar instalación de **ODBC Driver 18** (el equipo solo tiene el driver legacy `SQL Server`).
  - Hecho cuando: `pip install -e .[dev]` y `pytest tests/golden/test_golden_integrity.py` pasan en la venv; `docker compose up` levanta SQL y `sqlcmd`/pyodbc conecta.
  - _F-29, ADR-10_

- [x] 0.4 Extractor de goldens v2 + golden de septiembre — **data-engineer** *(hecho 2026-09-24: schema 1.2, 334 eventos, tabla 1975 filas, 37 tests de integridad OK)*
  - Extender `tests/golden/extract_golden.py`: parámetros efectivos (C5, C7, C21, C31, L2, L4, L14, `Daily!D5`, C23), L10, **todos** los eventos en orden de escritura, resumen `N6:Q172`, tabla `Calculation-Availability!E4:BO` (archivo `golden_<mes>_calc.json.gz`).
  - Invariante `Σ daily = C14` solo si `C21="No"`; agregar aviso `C14 ≈ 4·L10` cuando aplica (R13.8).
  - Regenerar `golden_2026_09_september.json` desde el libro actual (21 días de Daily, 334 eventos) y actualizar `golden_index.json` a `schema_version 1.2` (campo `effective_parameters`).
  - Actualizar `test_golden_integrity.py` (nuevos campos; eventos completos; tolerancias de la tabla única).
  - Hecho cuando: el extractor corre sin `--force` sobre el libro actual y los tests de integridad pasan.
  - _R14, F-05, F-08, F-09, F-25_

- [ ] 0.5 Data contract del libro — **technical-writer** (con **data-engineer**)
  - `docs/data-contract-libro.md`: hojas, columnas, tipos observados (módulos fraccionarios, fallas numéricas, vacíos), convención de fila 00:00, alineación por fila con PlantActivity, DST, celdas de parámetros.
  - _R4, R5, R6, GT-1/2/7/8_

- [ ] 0.6 **P0 (bloqueante solo para 4.6/4.7)** — muestra real del export SCADA — **Humano (Alex / Francisco)**
  - Guardar en `tests/fixtures/scada/` un export real (recortado si es grande) + nota con nombre de archivo, delimitador, encoding, formato de fecha, filas de encabezado.

- [ ] 0.7 **P1** — data contract SCADA — **technical-writer**
  - `docs/data-contract-scada.md` a partir de 0.6: columnas (= encabezados RawData-PCS), tipos, formato `mm-dd-aaaa hh:mm:ss`, convención DESDE/HASTA, nombre de archivo.
  - _R2.4, R2.5_

- [ ] **Checkpoint 0 — code-reviewer**: revisar 0.3 y 0.4 (extractor, esquema del golden, `.gitignore`). Preguntar al usuario si hay ajustes.

### Fase 1 — Núcleo de paridad (Python puro)

- [ ] 1.1 `config` — **data-engineer**
  - `ConfiguracionCalculo`, `CONFIG_POR_DEFECTO`, `construir_config()`, `desde_excel.leer_parametros()` con semántica `flag_si/flag_no`.
  - Tests: `total_racks`, defaults, validaciones, parámetros de eventos por defecto = KPI (R1.4), `"no"`≠`"No"` (R1.5).
  - _R1_

- [ ] 1.2 `excel_semantics` — **data-engineer** (revisión reforzada de **code-reviewer**)
  - Funciones de `design.md §excel_semantics`. Casos de test tomados del libro: descripción numérica `55`→`"F55"`, `"F55 EXTERNAL FAULT/OVGR"`→`"F55"`, `" X"`→`""`, `"ABC"`→`"FABC"`, `ROUND` half-up, serial↔datetime de `46266.010416666664`.
  - Property 8.
  - _R17, F-12, F-14_

- [ ] 1.3 `model` — **data-engineer**
  - `MatrizPCS`, `DatosActividad`, `Anomalia`, `RegistroLista`, `EventoFalla`, `MuestraDisponibilidad`, `DiaDisponibilidad`, `KpiMensual`.

- [ ] 1.4 `ingestion` — **data-engineer**
  - `xlsx_stream` (iterparse + sharedStrings + rels), `lector_libro` (RawData-PCS, PlantActivity, `modo_huecos`), `anomalias_timestamp`.
  - Test: serial exacto de fila 2 = serial del XML; 15.990 filas del libro real (marker `golden`); anomalía DST en 2026-09-06; rechazo de módulos texto con fixture sintético.
  - Property 2.
  - _R4, R16, F-10, F-16_

- [ ] 1.5 `normalization` — **data-engineer**
  - `validar_esquema`, `a_matriz`, `a_formato_largo`. Property 1.
  - _R5, F-02_

- [ ] 1.6 `enrichment` — **data-engineer**
  - Unión por fila; vacío = 0; anomalías `pa_vacio`, `pa_sin_timestamp`, `pa_desalineado`. Test contra libro real: 659 filas `pa_sin_timestamp` (GT-8).
  - _R6, F-01, F-31_

- [ ] 1.7 `availability` — **data-engineer**
  - `MotorDisponibilidad.calcular` según `design.md`. Properties 3, 4, 5. Unit: 3 filas × 2 PCS a mano; fila fuera de rango; C23 derivado; `solo_tiempo_operacional` con factor 0.
  - _R7, F-07, F-13_

- [ ] 1.8 `fault_events` — **data-engineer** (revisión reforzada de **code-reviewer**)
  - Emulación de `mcoCreateList` con `RegistroLista`; `resumen_por_codigo`.
  - Unit: los 4 caminos de descripción (incl. anterior vacío → `F1 Watchdog`); fin = última fila en falla; evento partido por celda vacía; cierre estricto `> L4+1`; arrastre entre PCS; `ExcelHabriaFallado` en fila 2; orden de suma `(s+4)-m`.
  - Properties 6, 7, 11.
  - _R8, F-03, F-04, F-06, F-11, F-18_

- [ ] 1.9 `aggregation.diaria` — **data-engineer**
  - Emulación de `mcoDailyAvailability` (sin factor operacional), días desde `inicio_periodo` hasta `fin_diario`. Property 9.
  - _R9, F-08, F-09_

- [ ] 1.10 `aggregation.mensual_anual` — **data-engineer**
  - `registrar_mes_oficial` (DiasMes = 20.59375 y 1977 bloques para sep) y `calcular_anual` (fórmulas G/H/I/J de Annual_AVA). Test con las filas jul/ago/sep de GT-5 → `J15 = 0.96725164144625086`.
  - _R10, F-20_

- [ ] 1.11 Golden runner — **data-engineer**
  - `tests/golden/conftest.py` + `test_golden_runner.py`: para cada mes `verified`, ejecutar ingesta→motores con los **parámetros efectivos del golden** y comparar niveles 2–5 con la tabla de tolerancias. Sep marcado `verified` al pasar.
  - Hecho cuando: septiembre pasa C12 exacto, C14/C16/C19, los 21 días de Daily, los 334 eventos campo a campo, L10 y el resumen N:Q.
  - _R14, R13.7_

- [ ] **Checkpoint A — code-reviewer**: checklist de paridad completo sobre Fase 1; cobertura ≥ 90 % en motores y `excel_semantics`. Preguntar al usuario si hay ajustes.

### Fase 2 — SQL Server

- [ ] 2.1 DDL — **database-optimizer**
  - `sql/00_database.sql`, `01_maestros.sql`, `02_corrida.sql`, `03_indices.sql` según `design.md §Data Models` (idempotentes, `CREATE OR ALTER` / `IF NOT EXISTS`).
  - _R11, R12_

- [ ] 2.2 Seeds — **database-optimizer** + **data-engineer**
  - `scripts/generar_seed_tipo_detencion.py` lee `PCS-Fault` (167 filas) y genera `sql/05_seed_tipo_detencion.sql` con `MERGE`; `sql/05_seed_proyecto.sql` (4 proyectos); `sql/06_seed_annual_manual.sql` (jul/ago de Annual_AVA como `excel_manual`, según D-06).
  - _R12.1, R12.3, R10.4, F-19_

- [ ] 2.3 Vistas, roles y consultas de auditoría — **database-optimizer**
  - `sql/04_vistas.sql` (`v_*_vigente`, `v_calidad_corrida`, `v_modulos_nulos_historico`), `sql/06_roles.sql` (`etl_writer`, `revisor`, `bi_reader`, `DENY DELETE`), `sql/07_audit_queries.sql`.
  - _R12.6, R15, R18.4_

- [ ] 2.4 `persistence` — **backend-architect**
  - `conexion.py` (URL desde `ETL_ARENA_DB_URL`, `fast_executemany`), `repositorio.py` (`iniciar`, `guardar_corrida` en una transacción, `finalizar`, `guardar_referencia_excel`, `guardar_reconciliacion`), resolución de `IdTipoDetencion` en memoria, `detencion` desde eventos.
  - _R11, R12.4_

- [ ] 2.5 Tests de integración SQL (marker `sql`, Docker) — **backend-architect**
  - DDL aplica dos veces sin error; dos corridas coexisten (append-only); fallo inyectado a mitad → rollback total y `etl_run.Estado='failed'`; `v_detencion_vigente` conserva la revisión entre corridas; `bi_reader` no puede leer tablas base.
  - _R11.1–R11.3, R12.5, R12.6_

- [ ] 2.6 Rendimiento — **database-optimizer**
  - Cargar una corrida real completa (~1M filas raw + ~120k muestras de período). Objetivo: `guardar_corrida` < 3 min en el equipo local; si no, `BULK INSERT` desde CSV temporal. Documentar tamaños con columnstore.
  - _F-28, ADR-09_

- [ ] **Checkpoint B — code-reviewer**: DDL vs design, append-only, seguridad, sin credenciales en repo.

### Fase 3 — Pipeline, calidad y reconciliación

- [ ] 3.1 `pipeline.py` + `scripts/ejecutar_etl.py` — **backend-architect**
  - Orden: config → ingesta → normalización → enriquecimiento → disponibilidad → eventos → agregaciones → calidad → persistencia → (reconciliación si `--referencia`). Logging estructurado con `IdCorrida`; exit ≠ 0 ante error.
  - _R11.2, R15, R19.4_

- [ ] 3.2 `reporting.calidad` — **data-engineer**
  - Resumen JSON (R15.3) + filas `data_quality_issue`.
  - _R15_

- [ ] 3.3 `reconciliation` — **data-engineer**
  - Niveles 1–5 + invariantes (con condiciones de aplicabilidad), `tolerancias.py` sincronizado con `golden_index._meta` (test), `ReporteReconciliacion` → `reconciliation_result`, `parity_failed` si fallan niveles 3–5, `sin_referencia` si no hay referencia.
  - Tests: datos idénticos → pass; C14 alterado → fail nivel 3; evento con D +15 min → fail nivel 5; invariante `no_aplica` cuando `L14 ≠ C31`.
  - _R13, F-24_

- [ ] 3.4 E2E local — **backend-architect**
  - `ejecutar_etl.py` sobre el libro real (sep) → SQL en Docker → reconciliar contra la referencia del golden v2 → `success` y todos los niveles `pass`.
  - _R11, R13, R14_

- [ ] **Checkpoint C — code-reviewer**: revisión del pipeline y de la reconciliación; confirmar con el usuario antes de tocar Excel real vía COM.

### Fase 4 — Cadena semanal (S → T → M → E → R → B)

- [ ] 4.1 `workbook.com.SesionExcel` — **backend-architect**
  - `DispatchEx`, visible, `DisplayAlerts=False`, watchdog por PID, cierre solo de la instancia propia; verificación de *Trusted Location*.
  - _R3.7, F-23_

- [ ] 4.2 `workbook.macros` + `workbook.referencia` — **backend-architect**
  - Escribir C5/C7/C21/C31/L2/L4/L14/`Daily!D5`, ejecutar las 4 macros en orden, guardar; extraer referencia completa por XML → `referencia_excel.json` + `excel_reference_*`.
  - _R3.5, R3.6, F-05, F-23_

- [ ] 4.3 Tests COM (marker `excel`) — **backend-architect**
  - Sobre una copia en `tmp` del libro: correr macros para sep 1–21 y verificar que la referencia extraída coincide con el golden v2.

- [ ] 4.4 Re-extraer goldens jul/ago (y jun) vía COM — **data-engineer** + **Humano**
  - Con los parámetros resueltos en D-03, generar `golden_2026_07/08` v2 (y junio si se quiere GT-3 como gate); correr el golden runner; marcar `verified` o registrar discrepancias. Esperado: evidencia del arrastre F-06 en los cierres de mes.
  - _R14, F-06, F-25_

- [ ] 4.5 `acquisition.acquire_wait` — **backend-architect**
  - Espera con timeout, validación de nombre/no vacío, sha256, movimiento a `data/processed/<corte>/`. Tests con archivos temporales.
  - _R2.1–R2.3_

- [ ] 4.6 `acquisition.contrato_scada` + `lector_scada` — **data-engineer** *(requiere 0.6/0.7)*
  - Validación de encabezados, parser de fecha estricto, rango DESDE/HASTA. Tests con el fixture P0.
  - _R2.4, R2.5, R3.3_

- [ ] 4.7 `workbook.preparar` — **backend-architect** *(requiere 4.1, 4.6)*
  - Copia + backup, escritura COM de RawData-PCS con seriales reales, validación de alineación con PlantActivity. Test COM: el libro preparado con el export P0 produce en macros el mismo C12 que el libro original para el mismo rango.
  - _R3.1–R3.4, F-21, F-22_

- [ ] 4.8 `reporting.notificacion` — **backend-architect**
  - Webhook configurable o `notificacion.md`; contenido R18.1; sin datos sensibles.
  - _R18_

- [ ] 4.9 `scripts/run_lunes.py` — **backend-architect**
  - Etapas, `run_state.json` reanudable, período por defecto D-07, `--oficial`, logs por etapa.
  - _R19_

- [ ] 4.10 Ensayo del lunes completo — **Humano (Francisco / Alex)** + **backend-architect**
  - `run_lunes.py --stage all` con un export real; registrar tiempos y problemas en `docs/shadow-log.md`.

- [ ] 4.11 *(Opcional, P8)* Fallback RPA TeamViewer — **backend-architect** — solo si TeamViewer falla de forma recurrente; fuera del camino crítico.

- [ ] **Checkpoint D — code-reviewer**: COM (cierre de instancias, rutas, no toca el maestro), orquestador reanudable, manejo de errores.

### Fase 5 — Documentación, shadow mode y handoff

- [ ] 5.1 Runbook — **technical-writer**
  - Actualizar `docs/runbook-lunes.md`: *Trusted Location*, Excel visible, etapas y reanudación, lectura del reporte de reconciliación y de calidad, PlantActivity desactualizada (F-31), recuperación ante fallos.

- [ ] 5.2 Sincronizar README / AGENTS / CLAUDE — **technical-writer**
  - Reflejar layout `src/etl_arena/`, tablas nuevas, estado de fases; eliminar menciones al límite de 167 eventos y al catálogo de 68 códigos.

- [ ] 5.3 Handoff Power BI — **technical-writer**
  - `docs/pbi-handoff.md` para Misael: vistas `v_*_vigente`, columnas, semántica C19 vs acumulado Annual, refresh por notificación.

- [ ] 5.4 Shadow mode — **Humano** + **data-engineer**
  - Mínimo 4 lunes consecutivos + 1 cierre mensual con reconciliación `pass` (o discrepancias explicadas por flags de Excel). Registro en `docs/shadow-log.md`.

- [ ] 5.5 Go/no-go — **software-architect** + **code-reviewer**
  - Verificar criterios de aceptación de `AGENTS.md §22`; decidir retiro del Excel como fuente oficial y abrir `availability-v2` (corrección de F-06, F-08, D-06).

---

## Task Dependency Graph

```json
{
  "waves": [
    { "id": 0, "tasks": ["0.1", "0.2", "0.3", "0.5", "0.6"] },
    { "id": 1, "tasks": ["0.4", "0.7", "1.1", "1.2", "1.3"] },
    { "id": 2, "tasks": ["checkpoint-0", "1.4", "2.1"] },
    { "id": 3, "tasks": ["1.5", "1.6", "2.2", "2.3"] },
    { "id": 4, "tasks": ["1.7", "2.4"] },
    { "id": 5, "tasks": ["1.8", "1.9", "1.10", "2.5"] },
    { "id": 6, "tasks": ["1.11", "2.6", "3.2"] },
    { "id": 7, "tasks": ["checkpoint-A", "checkpoint-B", "3.3"] },
    { "id": 8, "tasks": ["3.1"] },
    { "id": 9, "tasks": ["3.4", "4.1", "4.5", "4.8"] },
    { "id": 10, "tasks": ["checkpoint-C", "4.2", "4.6"] },
    { "id": 11, "tasks": ["4.3", "4.4", "4.7"] },
    { "id": 12, "tasks": ["4.9"] },
    { "id": 13, "tasks": ["4.10", "checkpoint-D", "5.1", "5.2", "5.3"] },
    { "id": 14, "tasks": ["5.4"] },
    { "id": 15, "tasks": ["5.5"] }
  ],
  "blocking_external": {
    "0.6 (P0 muestra SCADA)": ["0.7", "4.6", "4.7"],
    "0.2 (decisiones negocio)": ["4.4", "4.9 (período por defecto)", "5.5"]
  }
}
```

Camino crítico: 0.3 → 1.2 → 1.4 → 1.5/1.6 → 1.7 → 1.8 → 1.11 → 3.3 → 3.1 → 3.4 → 4.2 → 4.9 → 4.10 → 5.4.

---

## Notes

- La Fase 1 no necesita SQL ni Excel COM: todo se valida contra el golden de septiembre v2 (0.4).
- P0 bloquea solo la adquisición/adapter (4.6, 4.7), no los motores ni SQL.
- Property tests: Hypothesis, ≥ 200 ejemplos por propiedad, tag `# Feature: etl-arena-availability, Property N`.
- Tolerancias: única tabla en `design.md §Tolerancias`; `golden_index._meta` se sincroniza con ella.
- `VersionAlgoritmo = "availability-v1-excel-parity"` para toda corrida de esta fase.
- Server SCADA: solo exportar. Macros COM y ETL solo en la PC local.

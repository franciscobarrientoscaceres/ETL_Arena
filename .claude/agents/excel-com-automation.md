---
name: excel-com-automation
description: Especialista en automatización de Excel vía COM (pywin32) para el ETL Arena: copia de trabajo del .xlsm, escritura de RawData-PCS, ejecución de las macros VBA y extracción de la referencia. Usar para las tareas de workbook/ (4.1–4.3, 4.7) y para depurar fallos de macros.
---

## Contexto obligatorio del proyecto ETL Arena

Antes de trabajar lee: `CLAUDE.md`, las secciones pertinentes de `AGENTS.md` y del spec revisión 2 en `.kiro/specs/etl-arena-availability/` (`audit.md`, `requirements.md`, `design.md`, `tasks.md`). La tarea que recibas citará `R<n>`, secciones de `design.md` y hallazgos `F-xx`: respétalos al pie de la letra.

Reglas duras (no negociables en la fase de paridad `availability-v1-excel-parity`):
- No modificar ni "mejorar" el algoritmo del Excel: se emula, incluidos sus defectos (marcados con flags).
- No hardcodear 61, 4, 12, 15 ni 2928 en los motores: vienen de `ConfiguracionCalculo`.
- Los motores recorren en orden de fila origen, operan sobre seriales Excel y comparan celdas solo vía `etl_arena.excel_semantics`.
- `PlantActivity` se une por número de fila; celda vacía = 0.
- SQL Server con estado vigente por mes (ADR-12): `publicar_mes` reemplaza el mes solo si la corrida es `success`; el registro (`etl_run`, `correccion_dato`, `reconciliation_result`…) es append-only.
- Nunca editar `data/AvailabilityCalculation_*.xlsm` ni `data/processed/`; trabajar sobre copias en `data/work/` o temporales.
- Código y documentación en español, con la convención del design (SQL PascalCase, Python snake_case).
- Al terminar, informa qué archivos cambiaste, qué tests corriste y su resultado real.


---

# Excel COM Automation

Automatizas Excel 16 en Windows desde Python con `pywin32` de forma robusta y sin efectos colaterales.

## Reglas
- `win32com.client.DispatchEx("Excel.Application")`: instancia propia; `Visible=True` (las macros usan `.Select` y `ActiveWindow`), `DisplayAlerts=False`, `ScreenUpdating` a gusto.
- Nunca abras el maestro `data/AvailabilityCalculation_*.xlsm`: copia a `data/work/<corte>/` (o a un temporal en tests) y respalda antes de escribir.
- Escribe rangos en bloque con `Range.Value = tupla_de_tuplas`; fechas como serial `float` y formato visual `dd-mm-aaaa hh:mm:ss` (nunca texto: rompe el filtro de la macro, F-21).
- Setea todos los parámetros: `C5`, `C7`, `C21`, `C31`, `ListOfFaults!L2`, `L4`, `L14`, `Daily!D5` (F-05, F-23), luego `Application.Run` de `cmdCalcAvailability`, `mcoCreateList`, `mcoDailyAvailability`, `Graphupdate` en ese orden, y guarda.
- La carpeta de trabajo debe ser *Trusted Location*; si `Run` falla por macros deshabilitadas, informa el paso exacto en lugar de bajar la seguridad.
- Cierre: `Workbook.Close(SaveChanges=...)`, `Application.Quit()`, liberar referencias; watchdog por PID que termine **solo** la instancia creada si excede el timeout. Jamás `taskkill /IM EXCEL.EXE`.
- La extracción de la referencia se hace leyendo el XML del libro guardado (más rápido y exacto), no celda a celda por COM.

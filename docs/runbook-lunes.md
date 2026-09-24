# Runbook — Flujo del lunes (SCADA → ETL → SQL → Power BI)

Checklist operativa semanal. Detalle de diseño: `AGENTS.md` §14 Fase S.

**Reglas de oro**

- Server SCADA: **solo extraer**. Nunca correr macros ni ETL allí.
- Solo `RawData-PCS` sale de SCADA; `PlantActivity` se mantiene aparte.
- SQL: `IdCorrida` nuevo por lunes; append-only.
- Power BI: owner **Misael**; hasta que haya service principal → **notificar**, no API propia.

---

## 0. Precondiciones (antes de las ~03:00 del lunes)

- [ ] Ventana libre en SCADA / se coordinó con operaciones.
- [ ] PC local: Excel instalado (macros COM), Python del proyecto, TeamViewer.
- [ ] Existe `data/inbox/` (crear si no).
- [ ] Conoces el reporte a exportar y la convención de fechas:
  - `DESDE` = dato siguiente al último cargado en el libro base (ver `RawData-PCS`, última fila; hoy 2026-09-21 14:15 → desde 14:30)
  - `HASTA` = último dato disponible al momento del export
- [ ] Si el período es distinto al default, anotar DESDE/HASTA para `prepare-workbook`.

---

## 1. Export desde SCADA (~03:00 AM)

- [ ] TeamViewer al server SCADA.
- [ ] Abrir Dashboard / reporte de export de `RawData-PCS` (columnas = hoja destino).
- [ ] Setear fechas del reporte (manual).
- [ ] Exportar (CSV/XLSX — confirmar formato con **P0** muestra real).
- [ ] **No** procesar nada en el server.
- [ ] File transfer por TeamViewer → PC local:
  - destino: `data/inbox/raw_pcs_<corte>.<ext>`
  - convención de nombre: `raw_pcs_YYYY-MM-DD_HHMMSS.csv` (o la definida en P1).

---

## 2. Orquestador local (PC de desarrollo)

Desde la raíz del repo:

```powershell
# esperar/validar archivo
python scripts/run_lunes.py --stage acquire-wait --inbox data/inbox

# transformar fechas + escribir copia de trabajo del .xlsm
python scripts/run_lunes.py --stage prepare-workbook --inbox data/inbox --work data/work

# macros COM (referencia Excel)
python scripts/run_lunes.py --stage run-macros --work data/work --period-start ... --period-end ...

# ETL Python → SQL
python scripts/run_lunes.py --stage run-etl --work data/work --period-start ... --period-end ...

# reconciliación
python scripts/run_lunes.py --stage reconcile --work data/work

# notificar refresh PBI
python scripts/run_lunes.py --stage notify-bi

# o todo de corrida
python scripts/run_lunes.py --stage all --inbox data/inbox --work data/work
```

> Mientras no se cargue la `Exclusion_Matrix` del mes (D-17), el KPI semanal es **preliminar** respecto a eventos de exclusión (`ExcusablesPendientes = 1`); la notificación lo indica.

### Cierre mensual (cuando se entrega la `Exclusion_Matrix` del mes — D-17, F-37)

```powershell
# carga B/C/D por timestamp, registra cada celda cambiada y encola el cierre del mes
python scripts/run_lunes.py --stage load-exclusion-matrix --inbox data/inbox --work data/work --oficial
```

- [ ] `Exclusion_Matrix` alineada por fila con `RawData-PCS!A` y sin valores fuera de 0/1/2.
- [ ] Revisar `data/work/<corte>/cambios.csv` (qué celdas de la matriz cambiaron).
- [ ] `cierre_mensual` oficial con reconciliación `pass`.

### Corrección de datos ya cargados (D-13)

```powershell
python scripts/run_lunes.py --reproceso data/inbox/<tramo_corregido>.<ext> --work data/work --oficial
```

- [ ] Revisar `cambios.csv` / `v_correccion_dato`: cada celda cambiada con valor anterior y nuevo, y KPI antes/después.

> `run_lunes.py` puede no existir aún (tareas del SDD). Hasta entonces: exportar a mano, transformar fechas, correr macros y `ejecutar_etl.py` siguiendo AGENTS §14.

---

## 3. Validación tras la corrida

- [ ] `acquire-wait`: archivo presente, no vacío, primer dato = siguiente al último cargado (sin hueco ni solape distinto), sha256 registrado en log.
- [ ] `Exclusion_Matrix` del mes cargada (0/1/2 por PCS, `Comments` con la causa): con `C31 = L14 = "Yes"`, un evento de exclusión no marcado **no** se descuenta (F-37, D-17). PlantActivity solo importa si `C21 = "Yes"`.
- [ ] `scada_adapter`: no quedaron fechas `mm-dd-aaaa` ni fechas como texto en la hoja de destino (deben ser serial Excel — F-21); columnas = mapping `RawData-PCS`.
- [ ] Macros: `C12`, `C14`, `C16`, `C19` extraídos y guardados como referencia de la corrida.
- [ ] ETL: `etl_run.Status = success`; nuevo `IdCorrida`.
- [ ] `reconcile`: paridad en niveles 3–4 (C12/C14/C16/C19) dentro de tolerancia.
- [ ] Eventos (`ListOfFaults` / `fault_event`): mismos conteos o discrepancias documentadas.
- [ ] Reporte de calidad: `ModulosDisponiblesNulo`, fallback de descripción, join-misses PlantActivity.

---

## 4. Handoff Power BI

- [ ] SQL actualizado con el nuevo `IdCorrida`.
- [ ] Notificar a **Misael** (mensaje / webhook / canal acordado) para refresh del reporte.
- [ ] No intentar refresh vía API hasta que exista service principal (fase posterior).

---

## 5. Fallbacks

| Situación | Acción |
|---|---|
| TeamViewer caído | Reintentar; no hay alternativa de red hoy. RPA TeamViewer solo como P8 (no camino crítico). Pedir share UNC o API GPM a GPM. |
| Formato de fecha inesperado | Abortar adapter; documentar muestra para P1; no “arreglar a ciegas”. |
| PlantActivity desactualizada | Continuar ETL; reportar join-misses; actualizar fuente aparte. |
| Macros fallan (COM) | Guardar logs, reintentar una vez; si persiste, correr ETL de todos modos y marcar corrida con advertencia de referencia Excel ausente. |

---

## 6. Checklist de cierre semanal

- [ ] Archivo en `data/processed/` con sha256 (si ya existe la fase A/B).
- [ ] Referencia Excel (C12/C14/C16/C19) archivada con `IdCorrida`.
- [ ] Log de etapas (`run_lunes` o notas) guardado.
- [ ] PBI notificado / refresh en proceso.
- [ ] Incidencias del lunes anotadas (discrepancias, calidad de datos).

---

## Enlaces

- Plan: `AGENTS.md` §14 Fase S, §21 Etapa 0
- SDD: `.kiro/specs/etl-arena-availability/` (requirements, design, tasks)
- README: sección “Flujo semanal (lunes)”

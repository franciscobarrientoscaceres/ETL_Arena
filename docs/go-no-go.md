# Go/no-go — retiro del Excel como fuente oficial (tarea 5.5)

Para: Francisco, Alex y negocio (decisión) · revisión técnica: software-architect + code-reviewer.
Estado: **preparado 2026-09-25, decisión pendiente**. Se decide al terminar el shadow mode (`docs/shadow-log.md`).

---

## 1. Recomendación actual: **todavía no (no-go)**

La paridad con el Excel está demostrada bit a bit en tres meses. Faltan dos cosas: la operación real de varias
semanas (shadow mode) y las piezas que dependen de insumos externos. Condiciones para pasar a **go**:

1. Shadow mode cumplido: `shadow_log.py --resumen` devuelve `cumple: true` (4 lunes consecutivos + 1 cierre).
2. Lector del export SCADA (4.6/4.7) en uso, sin pegar filas a mano, y ensayo completo con un export real (4.10).
3. Carga de la `Exclusion_Matrix` probada con una entrega real de Alex (0.8), al menos en un cierre.
4. Checkpoint D cerrado.
5. Decisiones abiertas cerradas o aceptadas explícitamente como riesgo: D-11 (agosto), D-14 (códigos duplicados),
   D-18 (marca PCS61) y el uso de `PlantActivity` para exclusiones (hoy **no** se usa; espera confirmación de Alex).
6. Power BI conectado a las vistas `v_*_vigente` y acordado con Misael (`docs/pbi-handoff.md` §8).

## 2. Criterios de aceptación (`AGENTS.md §22`)

| # | Criterio | Estado | Evidencia |
|---|---|---|---|
| 1 | Una corrida Python reproduce la corrida Excel con el mismo input | ✅ | `tests/golden/test_golden_runner.py`: septiembre 1–21, julio y agosto completos, paridad bit a bit; `tests/com/test_macros_com.py`: macros reales por COM + reconciliación 5/5 |
| 2 | C12 reproducible | ✅ | 1.975 (sep, con cambio de hora), 2.976 (jul, ago) exactos |
| 3 | C14 reproducible | ✅ | Exacto en los tres meses; con `Exclusion_Matrix`, 264.731,768 exacto contra el libro de agosto (tabla E4:BO celda a celda) |
| 4 | C16 reproducible | ✅ | Exacto (0,9819924099052362 en sep) |
| 5 | C19 reproducible | ✅ | Exacto. Ojo: C19 no es el acumulado anual (ver `pbi-handoff.md` §5) |
| 6 | Eventos de `ListOfFaults` reconciliables | ✅ | 334 / 473 / 595 eventos campo a campo; Δ ≤ 5,7e-14 cuando L14 se escribe "No" (F-44) |
| 7 | Disponibilidad diaria reconciliable | ✅ | `Daily` completo (21 y 31 días) exacto tras escribir la plantilla B/F/G (F-41) |
| 8 | Acumulado anual reconciliable | ⚠️ | Cálculo de `Annual_AVA` reproducido con los meses vigentes. Julio se explica con la regla histórica (F-42); **agosto no se reproduce** (D-11, F-43) y queda importado como `excel_manual` |
| 9 | Períodos con cambio de hora validados | ✅ | Septiembre 2026 (salto del 06-09 00:00 → 01:00): C12 = 1.975. El cambio de abril es anterior al inicio de operación |
| 10 | Auditoría KPI → acumulado → muestra → input raw | ✅ | `sql/07_audit_queries.sql`, `tests/integration/test_sql.py` (C14 = Σ muestras, trazabilidad a archivo, hash, fila y columna) |
| 11 | Intervalos con `ModulosDisponiblesNulo` identificados en la BD | ✅ | `raw_pcs_sample.ModulosDisponiblesNulo`, vista `v_modulos_nulos_historico`, resumen de calidad |

## 3. Criterios operativos (spec revisión 2)

| Criterio | Estado | Referencia |
|---|---|---|
| Shadow mode: 4 lunes + 1 cierre | ⏳ Sin registros | `docs/shadow-log.md`, tarea 5.4 |
| Lector y contrato del export SCADA | ⏳ Falta la muestra | Tareas 0.6, 0.7, 4.6, 4.7 |
| Ensayo del lunes con export real | ⏳ | Tarea 4.10 (ensayo con libro preparado: ok, 2026-09-24) |
| Carga mensual de la `Exclusion_Matrix` | 🟡 Implementada; falta la muestra real | Tarea 4.12, 0.8 |
| Reproceso con registro de cambios | ⏳ Depende del formato SCADA | Tarea 4.13 |
| Maestro Excel v1.1 (referencia con exclusiones) | ⏳ Tarea humana | Tarea 4.0; sin él, los cierres "Con Exclusiones" quedan `sin_referencia` |
| Checkpoint D (COM + orquestador) | 🟡 Revisión parcial hecha | `tasks.md` |
| Power BI conectado | ⏳ | `docs/pbi-handoff.md` §8 |
| Persistencia en Azure SQL | ✅ | Carga de un mes real en ~46 s; esquema aplicado en `trina_etl` |

## 4. Riesgos que se aceptan al pasar a go

- **Paridad con defectos del VBA.** Python reproduce a propósito comportamientos discutibles del Excel (sección 5).
  Corregirlos cambia el KPI y requiere acuerdo de negocio.
- **Excel 2016 en la PC local** (F-40): `ListOfFaults!N:Q` queda sin ordenar y la hoja `Graph` sin actualizar. No
  afecta el KPI.
- **Transporte solo por TeamViewer** y proceso en una PC local (una persona, una máquina).
- **Oferta gratuita de Azure SQL**: la base se pausa y tiene límites mensuales de cómputo.

## 5. Después del go: `availability-v2`

Con el Excel retirado, la paridad deja de ser obligatoria y se pueden corregir los defectos heredados. Cada cambio
cambia el KPI: requiere aprobación de negocio, `VersionAlgoritmo` nuevo y recálculo comparativo.

| Tema | Hoy (v1.1, igual al Excel) | Propuesta a evaluar |
|---|---|---|
| F-06 | Un evento abierto al cierre del período se arrastra al PCS siguiente | Cerrar el evento en el PCS correcto |
| F-08 | `Daily` no aplica el factor operacional (difiere de C14 si C21 = "Yes") | Aplicarlo igual que C14 |
| D-06 | Acumulado anual 2026 desde julio | Definir con negocio el inicio del año contractual |
| F-03 | Fin de evento = última fila en falla (la duración omite el último bloque) | Fin = primera fila sin falla |
| F-04 | Descripción con fila anterior vacía → "F1 Watchdog" | Descripción real del código |
| C19 | `1 − C14 / (Racks × 365 × 24 × 4)` con constantes fijas | Acumulado anual real o retirar la métrica |
| F-36 | `numBlock As Integer` puede desbordar | Irrelevante en Python; documentar |

## 6. Decisión

| Fecha | Decisión | Quién | Condiciones / notas |
|---|---|---|---|
| | | | |

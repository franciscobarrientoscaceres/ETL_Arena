# Registro del periodo de prueba en paralelo (shadow mode)

Para: Francisco, Alex y quienes decidan si se deja de usar el Excel ([go/no-go](./go-no-go.md)).

## ¿Qué es esto?

Antes de apagar el Excel, hay que demostrar **en la vida real** que el nuevo sistema da lo mismo. Durante unas
semanas se trabaja "en paralelo":

- **El Excel sigue siendo el oficial.** Alex calcula y reporta la disponibilidad como siempre.
- **El nuevo sistema corre al lado**, cada lunes, con los mismos datos.
- **Aquí se anota** cada semana si los dos dieron lo mismo.

## ¿Cuándo termina?

Cuando la tabla de abajo tenga:

- ✅ **4 lunes seguidos** que terminaron bien (`success`) y dieron igual que el Excel (`pass`), o cuya diferencia
  está explicada en la nota;
- ✅ **1 cierre de mes** que terminó bien (idealmente "Con Exclusiones", con la matriz de Alex);
- ✅ en cada fila, el número de Alex igual al del nuevo sistema (o la diferencia explicada).

Para saber cuánto falta:

```powershell
.venv\Scripts\python scripts\shadow_log.py --resumen
```

## Cómo anotar una semana

Después de cada corrida (paso 5 de la [guía del lunes](./runbook-lunes.md)), con el número que Alex obtuvo en su
Excel (por ejemplo 0.9819 = 98,19 %):

```powershell
.venv\Scripts\python scripts\shadow_log.py --corte 2026-09-28 --kpi-excel-oficial 0.9819 --nota "sin novedades"
```

El programa agrega la fila solo. **No edites la tabla a mano** (el programa la lee para calcular el avance).

Qué significa cada columna:

| Columna | Significa |
|---|---|
| Registrado | Cuándo se anotó |
| Corte / Tipo / Período | Qué corrida es (semanal o cierre de mes) y qué fechas calculó |
| Etiqueta | "Sin Exclusiones" o "Con Exclusiones" |
| Estado | `success` = terminó bien |
| Reconciliación | Comparación automática con las macros del Excel: `pass` = iguales |
| C12 | Cantidad de bloques de 15 minutos del período |
| C14 Python / C14 Excel | Racks-bloque indisponibles según cada uno |
| C16 Python | La disponibilidad que calculó el nuevo sistema |
| KPI Alex | La disponibilidad que reportó Alex con su Excel (se ingresa a mano) |
| Δ | La diferencia: `C16 Python − KPI Alex` (idealmente 0) |
| N° corrida | El número correlativo de la corrida (1, 2, 3…), el que conviene usar al conversar |
| IdCorrida | El código interno de la corrida en la base de datos |
| Nota | Observaciones o explicación de diferencias |

## Registro

| Registrado | Corte | Tipo | Período | Etiqueta | Estado | Reconciliación | C12 | C14 Python | C14 Excel | C16 Python | KPI Alex | Δ | N° corrida | IdCorrida | Nota |
|---|---|---|---|---|---|---|---:|---:|---:|---:|---:|---:|---:|---|---|

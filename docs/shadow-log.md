# Shadow log — Excel oficial vs Python/SQL (tarea 5.4)

Para: Francisco, Alex y quien decida el go/no-go (`docs/go-no-go.md`).

Durante el shadow mode, **el Excel sigue siendo la fuente oficial**: Alex calcula y reporta el KPI como siempre, y
la cadena Python/SQL corre en paralelo cada lunes. Este registro acumula la evidencia para retirar el Excel.

## Criterio de salida

- **4 lunes consecutivos** con corrida semanal oficial en `success` y reconciliación `pass` contra la referencia
  Excel, o con cada discrepancia explicada (p. ej. F-40, F-44).
- **1 cierre mensual** en `success`, idealmente "Con Exclusiones" con la `Exclusion_Matrix` de Alex.
- En cada fila, el KPI que reportó Alex coincide con `C16 Python`, o la diferencia está explicada en la nota.

## Cómo registrar

Después de cada corrida (`docs/runbook-lunes.md` §7):

```powershell
.venv\Scripts\python scripts\shadow_log.py --corte <corte> --kpi-excel-oficial <KPI de Alex> --nota "<observaciones>"
.venv\Scripts\python scripts\shadow_log.py --resumen      # avance contra el criterio de salida
```

`Reconciliación` = resultado contra la referencia Excel que genera la propia corrida (macros por COM).
`KPI Alex` = valor que reportó Alex con su Excel ese día (se ingresa a mano). `Δ` = `C16 Python − KPI Alex`.

## Registro

| Registrado | Corte | Tipo | Período | Etiqueta | Estado | Reconciliación | C12 | C14 Python | C14 Excel | C16 Python | KPI Alex | Δ | IdCorrida | Nota |
|---|---|---|---|---|---|---|---:|---:|---:|---:|---:|---:|---|---|

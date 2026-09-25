# Handoff Power BI — Disponibilidad Arena BESS

Para: **Misael** (owner del reporte Power BI). Mantiene: Francisco Barrientos (ETL).
Estado: 2026-09-25, previo al shadow mode (tarea 5.4). Fuente de verdad del esquema: `sql/04_vistas.sql`.

> **En palabras simples:** cada lunes (y a fin de mes) el nuevo sistema calcula la disponibilidad y la guarda en
> una base de datos en la nube. Power BI solo tiene que leer unas **vistas** (tablas ya preparadas) que siempre
> muestran **un único resultado oficial por mes**. No hay que elegir corridas ni filtrar nada: si el cálculo de
> una semana sale mal, la vista sigue mostrando el último resultado bueno.
>
> Palabras técnicas: [glosario](./glosario.md). Cómo se relacionan los datos: [modelo de datos](./modelo-datos.md)
> (§6 muestra las vistas en un dibujo).

---

## 1. Qué cambia

El KPI de disponibilidad de Arena BESS (61 PCS × 4 baterías × 12 racks = 2.928 racks) dejaba sus
resultados en el libro Excel. Ahora el ETL escribe cada corrida en **Azure SQL**, y Power BI lee
**solo vistas** `v_*`. Esas vistas ya resuelven qué corrida es la oficial de cada mes, así que el
reporte no necesita conocer `IdCorrida` ni filtrar corridas.

| Cuándo | Qué corre | Qué muestran las vistas |
|---|---|---|
| Cada lunes | Corrida **semanal**: del día 1 del mes hasta el último dato exportado | El mes en curso, etiquetado **"Sin Exclusiones"** (oficial) |
| Fin de mes, cuando Alex entrega la `Exclusion_Matrix` | **Cierre mensual**: el mes completo con eventos excusados | Ese mes pasa a **"Con Exclusiones"** y reemplaza a la semanal |

Las dos etiquetas son oficiales. "Con Exclusiones" descuenta los eventos que Alex marca como
excusables, así que normalmente muestra una disponibilidad mayor.

## 2. Conexión

| Parámetro | Valor |
|---|---|
| Conector | SQL Server database (Azure SQL Database) |
| Servidor | `trina-etl.database.windows.net` |
| Base | `trina_etl` (**PROD**, la oficial). Para validar el reporte con datos de práctica: `trina_etl_prueba` (**TEST/QA**, mismas vistas) |
| Autenticación | Cuenta Microsoft / Entra ID (@trinasolar.com). No hay usuario ni contraseña SQL |
| Permiso | Rol `bi_reader`: `SELECT` solo sobre las vistas `v_*`; las tablas base no son visibles |
| Modo recomendado | Import (el volumen es chico y el dato cambia una vez por semana) |

**Alta de usuario:** la hace Francisco, como admin Entra de la base, una sola vez por persona:

```sql
CREATE USER [misael.xxx@trinasolar.com] FROM EXTERNAL PROVIDER;
ALTER ROLE bi_reader ADD MEMBER [misael.xxx@trinasolar.com];
```

**Tres cosas a tener en cuenta:**
- **Firewall.** Desde Power BI Desktop, la IP de tu red debe estar autorizada en el servidor (Francisco
  la agrega). Para refrescar desde el **servicio** Power BI hay que habilitar "Permitir que los
  servicios y recursos de Azure accedan a este servidor". Por ser Azure SQL no hace falta gateway.
- **La base se pausa sola** (serverless, oferta gratuita). La primera conexión después de un rato
  tarda **~1 minuto** mientras despierta y puede dar timeout: reintentar una vez.
- **Credenciales en el servicio:** configurar el dataset con OAuth2 (cuenta organizacional).

## 3. Cómo se elige la corrida vigente

Todas las vistas `*_vigente` salen de `v_corrida_oficial_vigente`, que toma **una corrida por
(proyecto, año, mes)** con estas reglas:

1. Solo corridas **oficiales** y en estado **`success`**. Una corrida que falla, o que no cuadra con
   Excel (`parity_failed`), nunca se publica: el reporte sigue mostrando la anterior.
2. Gana **"Con Exclusiones"** sobre "Sin Exclusiones".
3. A igualdad, gana la de período más largo (fin más tardío) y luego la más reciente.

El mes lo define la **fecha de fin del período**. Las corridas de prueba (`golden`) nunca aparecen.

## 4. Vistas

Todas traen `IdProyecto` (1 = Arena), `Anio` y `Mes`, que sirven para relacionar vistas entre sí. La
mayoría trae además `EtiquetaExclusiones` ("Sin Exclusiones" / "Con Exclusiones") para mostrarla en
el reporte.

| Vista | Grano | Para qué |
|---|---|---|
| `v_kpi_vigente` | 1 fila por mes | KPI del período: disponibilidad, bloques, racks |
| `v_daily_vigente` | 1 fila por día del mes | Curva diaria acumulada y variación |
| `v_monthly_kpi_vigente` | 1 fila por mes | KPI mensual oficial (incluye jul/ago importados de Excel) |
| `v_annual_vigente` | 1 fila por mes del año | Tabla anual (equivale a la hoja `Annual_AVA`) |
| `v_fault_event_vigente` | 1 fila por evento de falla | Lista de eventos (hoja `ListOfFaults`) |
| `v_fault_code_vigente` | 1 fila por código de falla | Pareto de horas-rack por código |
| `v_detencion_vigente` | 1 fila por detención | Eventos con catálogo y estado de revisión |
| `v_calidad_corrida` | corrida × tipo de anomalía | Salud de cada corrida (todas, no solo vigentes) |
| `v_modulos_nulos_historico` | intervalo × PCS | Muestras sin dato de módulos (auditoría) |
| `v_correccion_dato` | celda corregida | Cambios de datos por reproceso o matriz de exclusión |
| `v_corrida_oficial_vigente` | 1 fila por mes | Qué corrida alimenta cada mes (período, tipo, archivo) |

### `v_kpi_vigente`

| Columna | Significado | Celda Excel |
|---|---|---|
| `InicioPeriodo`, `FinPeriodo` | Período calculado (fechas, `DATETIME`) | C5, C7 |
| `TipoCorrida` | `semanal` o `cierre_mensual` | — |
| `BloquesMuestreo` | Intervalos de 15 min con dato en el período | C12 |
| `TotalRacks` | 2.928 | C11 |
| `BloquesRacksIndisponibles` | Σ racks indisponibles × intervalos | C14 |
| `DisponibilidadPeriodo` | **KPI principal**: `1 − C14 / (TotalRacks × C12)` | C16 |
| `DisponibilidadAnualAcumulada` | Ver §5: **no es** el acumulado anual | C19 |
| `HorasRackEventos` | Σ horas-rack de la lista de eventos | `ListOfFaults!L10` |

### `v_daily_vigente`

`DiaN` (1…n), `Dia`, `BloquesRacksIndisponiblesDiarios`, `BloquesRacksIndisponiblesAcumulados`,
`Disponibilidad` (acumulada desde el día 1 hasta ese día) y `Variacion` (respecto al día anterior;
0 el día 1). Es la hoja `Daily`. El último día de un mes cerrado coincide con `DisponibilidadPeriodo`
del mes.

### `v_monthly_kpi_vigente` y `v_annual_vigente`

- `v_monthly_kpi_vigente`: `DiasMes`, `BloquesMuestreo`, `BloquesRacksIndisponibles`,
  `DisponibilidadMensual`, `DisponibilidadContractual` (0,98), `Origen` (`corrida` o `excel_manual`).
- `v_annual_vigente`: agrega `BloquesMuestreoAcumulados`, `BloquesIndisponiblesAcumulados` y
  **`DisponibilidadAcumulada`**, que es el acumulado del año contra el que se compara el 0,98 contractual.

### Eventos y detenciones

- `v_fault_event_vigente`: `NumeroPCS`, `MarcaTiempoInicio`, `MarcaTiempoFin`, `DuracionHoras`,
  `CodigoFalla`, `DescripcionFalla`, `PromedioBateriasInvolucradas`, `HorasRackIndisponibles` y banderas
  de auditoría (`DescripcionFallaFallback`, `EventoArrastradoExcel`, `TieneExclusion`).
- `v_detencion_vigente`: lo mismo más el catálogo (`Significado`, `Operativo`) y la revisión
  (`EstadoRevision` = `pendiente` si nadie la revisó, `Observacion`, `RevisadoPor`, `RevisadoEn`). La
  revisión se conserva aunque la corrida del mes se reemplace.
- `v_fault_code_vigente`: `Ranking`, `CodigoFalla`, `DescripcionFallaPE`, `HorasRackIndisponibles`,
  `Porcentaje`.

## 5. Semántica que conviene no confundir

- **C19 (`DisponibilidadAnualAcumulada` en `v_kpi_vigente`) no es el acumulado del año.** El Excel la
  calcula como `1 − C14 / (TotalRacks × 365 × 24 × 4)`: la indisponibilidad **del período** repartida
  en un año completo. Por eso siempre está muy cerca de 1 (septiembre 1–21: 0,99899). Se publica por
  paridad con el Excel. Para el KPI anual, usar `v_annual_vigente.DisponibilidadAcumulada`.
- **Acumulado anual.** En 2026 parte en **julio** (como el libro); desde 2027 parte en enero.
  Acumula bloques y bloques indisponibles de cada mes vigente.
- **Mes en curso.** Mientras el mes no se cierra, su fila en `v_monthly_kpi_vigente` y
  `v_annual_vigente` es **parcial** (día 1 → último dato; `DiasMes` puede ser fraccionario, p. ej.
  20,59). Cambia cada lunes y queda fija con el cierre mensual.
- **Julio y agosto 2026** vienen importados del Excel (`Origen = excel_manual`): son los valores
  oficiales que se reportaron. El de agosto no se puede recalcular con ninguna regla conocida y está en
  revisión con Alex (D-11). No aparecen en `v_kpi_vigente` ni en `v_daily_vigente`, porque esas vistas
  necesitan una corrida.
- **Intervalos existentes.** `BloquesMuestreo` cuenta los intervalos que tienen dato, no los del
  calendario. Por ejemplo, el cambio de hora de septiembre deja 1.975 intervalos en lugar de 1.977.
  Coincide con el Excel.
- **Horas.** Las marcas de tiempo son hora local de Chile, tal como vienen del SCADA (sin conversión a UTC).

## 6. Refresh

Hoy el refresh es **por notificación**: Power BI no se refresca por API hasta tener un service principal.

1. El lunes, al terminar la cadena, llega un aviso: mensaje en el canal/webhook acordado o, si no
   hay webhook, `notificacion.md` que reenvía Francisco. El aviso trae el período, la etiqueta, el
   estado, los KPI y los cierres mensuales pendientes.
2. Si el estado es **`success`**, refrescar el dataset.
3. Si el estado **no** es `success`, no hace falta refrescar: las vistas siguen mostrando la corrida
   vigente anterior. Francisco revisa y vuelve a correr.

Opcional: agendar un refresh diario en el servicio. Es inocuo, porque las vistas solo cambian cuando
entra una corrida vigente nueva.

## 7. Recomendaciones para el modelo

- Relacionar por (`IdProyecto`, `Anio`, `Mes`). No usar `IdCorrida` como filtro ni como clave de
  negocio: cambia cuando un mes se reemplaza, por ejemplo al pasar a "Con Exclusiones".
- Mostrar siempre `EtiquetaExclusiones` junto al KPI del mes.
- Tipos: disponibilidades `FLOAT` entre 0 y 1 (formatear como %); fechas `DATETIME`.
- Para validar el dataset: septiembre 1–21 de 2026 debe dar C12 = 1.975, C14 = 104.134,292,
  C16 = 0,981992 (mismo valor que el Excel).

## 8. Por acordar con Misael

- [ ] Canal de la notificación (Teams / Power Automate webhook → variable `ETL_ARENA_NOTIFY_WEBHOOK`).
- [ ] Cuentas que necesitan `bi_reader` y si el refresh se hará desde el servicio (firewall de Azure).
- [ ] Si el reporte muestra el mes en curso (parcial) o solo meses cerrados.
- [ ] Fecha de inicio del shadow mode (Excel y SQL en paralelo) y criterio para dejar el Excel.

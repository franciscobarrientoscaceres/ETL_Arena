# Handoff Power BI — Disponibilidad Arena BESS

Para: **Misael** (owner del reporte Power BI). Mantiene: Francisco Barrientos (ETL).
Estado: 2026-09-26, modelo "estado vigente por mes" ([ADR-12](./adr/ADR-12-estado-vigente.md)), previo al shadow mode
(tarea 5.4). Fuente de verdad del esquema: `sql/02_corrida.sql` y `sql/04_vistas.sql`.

> **En palabras simples:** cada lunes (y a fin de mes) el nuevo sistema calcula la disponibilidad y la guarda en
> una base de datos en la nube. La base tiene **una sola versión de cada mes**: cuando se vuelve a calcular un mes,
> sus datos se **reemplazan**. Power BI lee esas tablas directamente; no hay corridas que elegir ni filtros que
> poner. Si un cálculo sale mal, el mes no se toca y sigue mostrando el último resultado bueno.
>
> Palabras técnicas: [glosario](./glosario.md). Cómo se relacionan los datos: [modelo de datos](./modelo-datos.md)
> (§6 muestra lo que lee Power BI en un dibujo).

---

## 1. Qué cambia

El KPI de disponibilidad de Arena BESS (61 PCS × 4 baterías × 12 racks = 2.928 racks) dejaba sus resultados en el
libro Excel. Ahora el ETL escribe cada mes en **Azure SQL**, en tablas que siempre tienen **solo lo vigente**.

| Cuándo | Qué corre | Qué queda en las tablas |
|---|---|---|
| Cada lunes | Carga **semanal**: del día 1 del mes hasta el último dato exportado | El mes en curso se **reemplaza** (ahora con más días), etiquetado **"Sin Exclusiones"** |
| Fin de mes, cuando Alex entrega la `Exclusion_Matrix` | **Cierre mensual**: el mes completo con eventos excusados | Ese mes se reemplaza por la versión **"Con Exclusiones"** |

Las dos etiquetas son oficiales. "Con Exclusiones" descuenta los eventos que Alex marca como excusables, así que
normalmente muestra una disponibilidad mayor.

## 2. Conexión

| Parámetro | Valor |
|---|---|
| Conector | SQL Server database (Azure SQL Database) |
| Servidor | `trina-etl.database.windows.net` |
| Base | `trina_etl` (**PROD**, la oficial). Para validar el reporte con datos de práctica: `trina_etl_prueba` (**TEST/QA**, mismas tablas) |
| Autenticación | Cuenta Microsoft / Entra ID (@trinasolar.com). No hay usuario ni contraseña SQL |
| Permiso | Rol `bi_reader`: `SELECT` sobre las tablas de estado, el catálogo `tipo_detencion` y las vistas `v_*` (solo lectura) |
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

## 3. Qué hay en las tablas

Solo se guarda un mes cuando el cálculo **cuadra con el Excel** (`success`). Si no cuadra, falla o es una prueba,
el mes queda como estaba. El reemplazo es de "todo o nada": nunca se ve un mes a medias.

Todas las tablas traen `IdProyecto` (1 = Arena), `Anio` y `Mes` para relacionarlas, y `NumCorrida` (qué carga
escribió esa fila por última vez; solo informativo).

| Tabla / vista | Grano | Para qué | Antes (Excel) |
|---|---|---|---|
| `disponibilidad_mensual` | 1 fila por mes | **KPI del mes**: C12, C14, C16, C19, L10; hasta qué dato llega y si está completo | `Calculation-Availability` |
| `v_disponibilidad_anual` | 1 fila por mes del año | Acumulado del año contra el 98 % contractual | `Annual_AVA` |
| `disponibilidad_diaria` | 1 fila por día | Curva diaria acumulada y variación | `Daily` |
| `v_detencion` | 1 fila por detención | Eventos de falla con catálogo y estado de revisión | `ListOfFaults` |
| `v_resumen_codigo_mensual` | código de falla × mes | Pareto de horas-rack por código | `ListOfFaults!N:Q` |
| `muestra_pcs` | bloque de 15 min × PCS | Detalle fino: dato de SCADA y aporte al C14 (auditoría) | `RawData-PCS` |
| `calidad_dato` | anomalía | Problemas de datos del mes (huecos, cambio de hora…) | — |
| `v_modulos_nulos` | bloque × PCS | Bloques sin dato de módulos (auditoría) | — |
| `v_correccion_dato` | dato cambiado | Qué datos cambiaron al recargar o al cargar la matriz | — |
| `v_ejecuciones` | 1 fila por carga | Registro de cargas: cuándo, qué archivo, si publicó | — |

### `disponibilidad_mensual`

| Columna | Significado | Celda Excel |
|---|---|---|
| `InicioPeriodo`, `UltimoDato` | Desde el día 1 hasta el último dato cargado | C5, C7 |
| `MesCompleto` | 1 cuando el mes tiene todos sus días | — |
| `BloquesMuestreo` | Intervalos de 15 min con dato | C12 |
| `TotalRacks` | 2.928 | C11 |
| `BloquesRacksIndisponibles` | Σ racks indisponibles × intervalos | C14 |
| `DisponibilidadMensual` | **KPI principal**: `1 − C14 / (TotalRacks × C12)` | C16 |
| `DisponibilidadAnualAcumulada` | Ver §4: **no es** el acumulado anual | C19 |
| `HorasRackEventos` | Σ horas-rack de la lista de eventos | `ListOfFaults!L10` |
| `DisponibilidadContractual` | 0,98 | — |
| `EstadoExclusiones` | `sin_exclusiones` o `con_exclusiones` (mostrar como "Sin/Con Exclusiones") | — |
| `TipoCorrida` | `semanal` o `cierre_mensual` | — |
| `Origen` | `corrida` o `excel_manual` (julio y agosto 2026) | — |

### `v_disponibilidad_anual`

Por cada mes: `DisponibilidadMensual`, `BloquesMuestreoAcumulados`, `BloquesIndisponiblesAcumulados` y
**`DisponibilidadAcumulada`**, que es el acumulado del año contra el que se compara el 0,98 contractual.

### `disponibilidad_diaria`

`DiaN` (1…n), `Dia`, `BloquesRacksIndisponiblesDiarios`, `BloquesRacksIndisponiblesAcumulados`, `Disponibilidad`
(acumulada desde el día 1 hasta ese día) y `Variacion` (respecto al día anterior; 0 el día 1). El último día de un
mes cerrado coincide con `DisponibilidadMensual`.

### Detenciones

- `v_detencion`: `IdDetencion`, `NumeroPCS`, `FechaInicio`, `FechaTermino`, `DuracionSegundos`, `DuracionHoras`,
  `CodigoFalla`, `DescripcionFalla`, `Significado`, `Operativo`, `PromedioBateriasInvolucradas`,
  `HorasRackIndisponibles`, `EsExcusable` y la revisión (`EstadoRevision` = `pendiente` si nadie la revisó,
  `Observacion`, `RevisadoPor`, `RevisadoEn`).
- **`IdDetencion` es estable**: cuando el mes se recarga, una detención que sigue existiendo conserva su número y su
  revisión. Si una corrección de datos la hace desaparecer, se quita.
- Una detención pertenece al mes que la calculó (`Anio`, `Mes`), aunque su `FechaInicio` sea el 23:45 del último
  día del mes anterior (así lo hace el Excel).
- `v_resumen_codigo_mensual`: `CodigoFalla`, `DescripcionFallaPE`, `Detenciones`, `DuracionHoras`,
  `HorasRackIndisponibles`, `Porcentaje` (del mes).

## 4. Semántica que conviene no confundir

- **C19 (`DisponibilidadAnualAcumulada`) no es el acumulado del año.** El Excel la calcula como
  `1 − C14 / (TotalRacks × 365 × 24 × 4)`: la indisponibilidad **del período** repartida en un año completo. Por eso
  siempre está muy cerca de 1 (septiembre 1–21: 0,99899). Se publica por paridad con el Excel. Para el KPI anual,
  usar `v_disponibilidad_anual.DisponibilidadAcumulada`.
- **Acumulado anual.** En 2026 parte en **julio** (como el libro); desde 2027 parte en enero.
- **Mes en curso.** Mientras el mes no se cierra, su fila es **parcial** (día 1 → `UltimoDato`, `MesCompleto = 0`;
  `DiasMes` puede ser fraccionario, p. ej. 20,59). Se reemplaza cada lunes y queda fija con el cierre mensual.
- **Julio y agosto 2026** vienen importados del Excel (`Origen = excel_manual`): son los valores oficiales que se
  reportaron. El de agosto no se puede recalcular con ninguna regla conocida y está en revisión con Alex (D-11).
  Esos meses no tienen días ni detenciones en la base (solo la fila mensual).
- **Intervalos existentes.** `BloquesMuestreo` cuenta los intervalos que tienen dato, no los del calendario. Por
  ejemplo, el cambio de hora de septiembre deja 1.975 intervalos en lugar de 1.977. Coincide con el Excel.
- **Horas.** Las marcas de tiempo son hora local de Chile, tal como vienen del SCADA (sin conversión a UTC).

## 5. Refresh

Hoy el refresh es **por notificación**: Power BI no se refresca por API hasta tener un service principal.

1. El lunes, al terminar la cadena, llega un aviso: mensaje en el canal/webhook acordado o, si no hay webhook,
   `notificacion.md` que reenvía Francisco. El aviso dice **qué mes se reemplazó**, la etiqueta, el estado, los KPI y
   los cierres mensuales pendientes.
2. Si dice **"Publicado: mes AAAA-MM reemplazado"**, refrescar el dataset.
3. Si dice **"No publicado"**, no hace falta refrescar: las tablas siguen con la versión anterior. Francisco revisa y
   vuelve a correr.

Opcional: agendar un refresh diario en el servicio. Es inocuo, porque las tablas solo cambian cuando se publica un
mes.

## 6. Recomendaciones para el modelo

- Relacionar por (`IdProyecto`, `Anio`, `Mes`) y, para detenciones, por `IdDetencion`.
- No usar `NumCorrida` como filtro: cambia cada vez que el mes se recarga. Sirve para mostrar "actualizado por la
  carga N° 15" y buscar esa carga en `v_ejecuciones`.
- Mostrar siempre la etiqueta de exclusiones (`EstadoExclusiones`) junto al KPI del mes.
- Tipos: disponibilidades `FLOAT` entre 0 y 1 (formatear como %); fechas `DATETIME`.
- Para validar el dataset: septiembre 1–21 de 2026 debe dar C12 = 1.975, C14 = 104.134,292, C16 = 0,981992 (mismo
  valor que el Excel).

## 7. Por acordar con Misael

- [ ] Canal de la notificación (Teams / Power Automate webhook → variable `ETL_ARENA_NOTIFY_WEBHOOK`).
- [ ] Cuentas que necesitan `bi_reader` y si el refresh se hará desde el servicio (firewall de Azure).
- [ ] Si el reporte muestra el mes en curso (parcial) o solo meses cerrados.
- [ ] Fecha de inicio del shadow mode (Excel y SQL en paralelo) y criterio para dejar el Excel.

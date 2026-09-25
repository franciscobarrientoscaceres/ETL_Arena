# Glosario — las palabras del proyecto, explicadas en simple

Si una palabra de la documentación no se entiende, está aquí. Van en orden alfabético dentro de cada grupo.

---

## La planta

| Palabra | Qué significa |
|---|---|
| **Arena BESS** | La planta de baterías que medimos. *BESS* = *Battery Energy Storage System*: un sistema que guarda energía eléctrica en baterías y la entrega cuando se necesita. Opera desde el 08-04-2026. |
| **PCS** | *Power Conversion System*. Es el equipo que convierte la energía entre las baterías (corriente continua) y la red (corriente alterna). Arena tiene **61 PCS**, numerados del 01 al 61. |
| **BEC / batería / módulo** | Cada PCS tiene conectadas **4 baterías** (módulos BEC). En total: 61 × 4 = **244 baterías**. |
| **Rack** | Cada batería está formada por **12 racks** (estantes con celdas). En total: 61 × 4 × 12 = **2.928 racks**. El KPI se mide en racks. |
| **`NUMBER_OF_MODULES`** | El dato más importante. SCADA informa, cada 15 minutos, **cuántas de las 4 baterías de un PCS están funcionando**. Si dice 4, el PCS está completo. Si dice menos de 4, faltan baterías y hay racks indisponibles. Puede venir con decimales (p. ej. 3,2347) y se usa tal cual. Si viene **vacío**, se cuenta como 4 (igual que hacía el Excel) y se marca para revisarlo. |
| **Código de falla (F13, F55…)** | El código que el PCS informa cuando algo anda mal (`CURRENT FAULT`). Sirve para **explicar** una falla, pero **no decide** si hay indisponibilidad: eso lo decide `NUMBER_OF_MODULES`. El catálogo completo está en la hoja `PCS-Fault`. |

## Los datos y el Excel

| Palabra | Qué significa |
|---|---|
| **SCADA** | El sistema de monitoreo de la planta. Guarda lo que informa cada equipo. Nosotros **solo exportamos** datos desde el server SCADA; nunca procesamos nada allí. |
| **Export** | El archivo que se saca de SCADA cada lunes con los datos nuevos de `RawData-PCS`. Se trae a la PC por **TeamViewer** (programa de acceso remoto). |
| **Bloque de 15 minutos (muestra, intervalo)** | SCADA guarda un dato cada 15 minutos. Cada uno de esos momentos es un bloque. Un día normal tiene 96 bloques. |
| **Libro / libro Excel** | El archivo `.xlsm` donde está todo el cálculo original (con macros). El oficial es el de septiembre 2026 (carpeta `data/`). |
| **Hojas del libro** | `RawData-PCS` (datos de SCADA), `Calculation-Availability` (parámetros y resultados), `ListOfFaults` (eventos), `Daily` (día a día), `Annual_AVA` (año), `Exclusion_Matrix` (exclusiones), `PlantActivity` (actividad de la planta), `PCS-Fault` (catálogo de códigos). |
| **Macro / VBA** | Pequeños programas guardados dentro del Excel que hacen el cálculo (`cmdCalcAvailability`, `mcoCreateList`, `mcoDailyAvailability`, `Graphupdate`). VBA es el lenguaje en que están escritos. **El proyecto nunca modifica las macros.** |
| **COM** | La forma en que Python "maneja" Excel desde afuera: abre el libro, escribe los parámetros y aprieta el botón de las macros por nosotros. |
| **C5, C7, C12, C14, C16, C19…** | Celdas de la hoja `Calculation-Availability`. Las más importantes: **C12** = cuántos bloques hay en el período; **C14** = racks indisponibles acumulados; **C16** = disponibilidad del período (el KPI); **C19** = una medida anual del Excel (ver más abajo). |
| **C21 / C31 / L14** | Interruptores "Yes/No" del Excel. **C21**: contar solo el tiempo en que la planta operó (oficial: "No"). **C31** (KPI) y **L14** (eventos): aplicar las exclusiones (oficial: "Yes"). |
| **PlantActivity** | Hoja con la actividad de la planta. Hoy solo se usa para C21. **No se usa para exclusiones** hasta que Alex lo confirme. |

## El cálculo

| Palabra | Qué significa |
|---|---|
| **Disponibilidad (KPI)** | Qué porcentaje del tiempo estuvieron disponibles los racks. Fórmula: `1 − C14 / (2.928 × C12)`. Si nunca falla nada, da 100 %. Es el número contractual. |
| **Racks-bloque indisponibles (C14)** | Se suma, bloque por bloque y PCS por PCS, cuántos racks faltaban: `(4 − NUMBER_OF_MODULES) × 12`. Ejemplo: si un PCS tiene 3 baterías durante un bloque, faltan 12 racks en ese bloque. |
| **Evento de falla** | Varios bloques seguidos en que un mismo PCS tuvo menos de 4 baterías se juntan en un solo evento, con inicio, fin, duración, código de falla y horas-rack perdidas. Es la hoja `ListOfFaults`. |
| **Horas-rack** | Otra forma de medir lo perdido: racks indisponibles × horas. |
| **Exclusión / evento excusable** | Una falla que **no** es culpa de la planta (p. ej. un corte de la red eléctrica) y por eso no debe restar disponibilidad. |
| **`Exclusion_Matrix`** | La planilla que entrega **Alex a fin de mes** marcando las exclusiones, con una columna por PCS: **0 o vacío** = no se excusa; **1** = se excusa todo (se cuentan las 4 baterías como disponibles); **2** = se excusa solo lo nuevo (se mantiene la falla que ya existía antes del evento). |
| **Sin Exclusiones / Con Exclusiones** | Etiqueta de cada resultado. Las corridas semanales salen **Sin Exclusiones** (la matriz aún no llega). El cierre del mes sale **Con Exclusiones**. Las dos son oficiales. |
| **C19 ("acumulada anual" del Excel)** | Ojo: **no** es la disponibilidad del año. Es la falla del período repartida en un año completo, por eso siempre da cerca de 100 %. La disponibilidad del año de verdad está en `Annual_AVA` / `v_annual_vigente`. |
| **Cambio de hora (DST)** | En Chile, en septiembre el reloj salta de 00:00 a 01:00 (en 2026, el 06-09) y en los datos faltan esos bloques. Por eso septiembre 1–21 tiene **1.975** bloques en `RawData-PCS` y no los 1.977 que diría el calendario. Se cuentan los bloques que existen, igual que el Excel. |

## El proceso

| Palabra | Qué significa |
|---|---|
| **Corte** | El nombre de una corrida en disco, normalmente la fecha (p. ej. `2026-09-28`). Cada corte tiene su carpeta en `data/work/`. |
| **Corrida (`NumCorrida`, `IdCorrida`)** | Cada vez que se calcula algo y se guarda en la base de datos es una corrida. Tiene un **número** para las personas (`NumCorrida`: 1, 2, 3…, "la corrida 12") y un **código interno** único para el sistema (`IdCorrida`, como `4d763e71-…`). Nunca se borra ni se pisa una corrida anterior. |
| **Semanal / cierre mensual** | La **semanal** se corre cada lunes: desde el día 1 del mes hasta el último dato. El **cierre mensual** calcula el mes completo cuando ya terminó. |
| **Oficial / vigente** | Una corrida **oficial** puede publicarse. La **vigente** es la que Power BI muestra para cada mes: la oficial más reciente que terminó bien, prefiriendo "Con Exclusiones". |
| **Libro base / maestro / copia de trabajo** | El **maestro** es el libro original de `data/`. El **libro base** es el libro de la última corrida oficial (la próxima parte desde ahí). La **copia de trabajo** es la copia que usa cada corte; los originales nunca se tocan. |
| **Referencia Excel** | Los resultados que calcula el Excel (vía macros) para el mismo período. Sirven para comparar. |
| **Reconciliación** | La comparación automática entre lo que calcula Python y lo que calcula el Excel. `pass` = son iguales; `fail` / `parity_failed` = hay diferencias y **no se publica**; `sin_referencia` = no hubo Excel para comparar (no es error). |
| **Paridad** | Que Python dé **exactamente** lo mismo que el Excel. Hoy se cumple bit a bit en julio, agosto y septiembre. |
| **Golden** | Resultados guardados del Excel que se usan como "respuestas correctas" en las pruebas automáticas. |
| **Shadow mode** | Período de prueba en paralelo: el Excel sigue siendo el oficial y Python corre al lado cada lunes, para comprobar que dan lo mismo. |
| **Go / no-go** | La decisión final de si ya se puede dejar de usar el Excel como fuente oficial. |
| **Ambiente PROD / TEST-QA** | Las dos bases de datos del proyecto. **PROD** (`trina_etl`) guarda lo oficial. **TEST/QA** (`trina_etl_prueba`) es un "cajón de arena" con las mismas tablas, más la carpeta `data/work/_prueba`, para practicar y validar sin tocar lo oficial. Se elige con `--entorno prueba`. |

## Computación

| Palabra | Qué significa |
|---|---|
| **Python** | El lenguaje en que está escrito el nuevo cálculo. |
| **Entorno virtual (`.venv`)** | Una carpeta con un Python propio del proyecto y sus librerías, para no mezclarlo con otros programas del computador. |
| **Git / repositorio / clonar** | Git guarda todas las versiones del proyecto. El **repositorio** es el proyecto completo en GitHub; **clonar** es descargarlo al computador. |
| **Terminal / PowerShell** | La ventana negra/azul de Windows donde se escriben comandos. |
| **`.env`** | Archivo de configuración local (no se sube a GitHub) con la dirección de la base de datos. No lleva contraseñas. |
| **Base de datos / Azure SQL** | Donde se guardan todos los resultados, en la nube de Microsoft (servidor `trina-etl.database.windows.net`; bases `trina_etl` = PROD y `trina_etl_prueba` = TEST/QA). |
| **Tabla / vista** | Una **tabla** guarda datos (como una hoja de Excel con columnas fijas). Una **vista** es una "consulta guardada" que junta tablas y muestra solo lo útil; Power BI lee vistas. |
| **Append-only** | En la base de datos **solo se agrega**, nunca se borra ni se modifica lo anterior. Así siempre se puede auditar qué pasó. |
| **Entra ID** | El inicio de sesión de Microsoft con la cuenta @trinasolar.com. Reemplaza a las contraseñas. |
| **Firewall de Azure** | La "puerta" del servidor: solo deja entrar a las direcciones de internet (IP) autorizadas. |
| **ODBC Driver 18** | Un programa de Microsoft que permite que Python "hable" con la base de datos. |
| **Power BI** | La herramienta de reportes que ve la gerencia. La maneja **Misael**. |

# ADR-10 — Entorno de desarrollo

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R11

## Contexto
F-29: el PC de oficina tenía solo Python 3.14, el driver ODBC legacy `SQL Server` (no maneja bien `DATETIME2` ni `fast_executemany`) y no tenía SQL Server local. El PC `FRANCISCO-PC` (2026-09-24) tiene Python 3.13, ODBC Driver 17 y 18, y SQL Server 2025 Developer (instancia `SQLSERVER2025DEV`).

## Opciones
- Python: fijar una versión, o `requires-python >= 3.13` (hay wheels de todas las dependencias en 3.13 y 3.14).
- SQL: instancia local (el contenedor Docker de la primera versión se retiró el 2026-09-25: no se usaba).
- Autenticación de desarrollo: `sa` con contraseña en `.env`, o autenticación Windows.

## Decisión
- `requires-python >= 3.13` (D-10). Una `.venv` por equipo.
- **ODBC Driver 18 for SQL Server** obligatorio.
- SQL de desarrollo: la instancia local (`FRANCISCO-PC\SQLSERVER2025DEV`) para los tests de integración; producción y pruebas en Azure (ver actualizaciones abajo).
- Desarrollo con **autenticación Windows** (`trusted_connection=yes`): no hay contraseñas en `.env`. `sa` no se usa desde el código; la carga productiva usa el login de mínimo privilegio `etl_writer` (tarea 2.3).

## Consecuencias
- `ETL_ARENA_DB_URL` varía por equipo y vive solo en `.env` (no versionado).
- El DDL debe ser compatible con SQL Server 2022 y 2025.
- Instancia nombrada: si la conexión desde otro equipo falla, verificar el servicio *SQL Server Browser* o fijar el puerto.

## Actualización 2026-09-24 — Azure SQL Database (serverless, oferta gratuita)

**Contexto:** en Trina no se permite instalar SQL Server en la laptop de trabajo.

**Decisión:** la base de producción pasa a **Azure SQL Database, General Purpose serverless (Gen5), oferta gratuita** (100.000 vCore-segundos + 32 GB al mes, sin cargo al superar: pausa hasta el mes siguiente):
- Servidor `trina-etl.database.windows.net`, base `trina_etl`, región **Brazil South** (Chile Central no admite la oferta gratuita), grupo de recursos `rg_trina_etl`, suscripción *Azure subscription 1* de F. Barrientos (cuenta @trinasolar.com).
- Autenticación **Microsoft Entra ID** con token de `azure-identity` (`ETL_ARENA_DB_AUTH=entra`, `persistence/entra.py`): sin contraseñas en `.env`; el navegador se abre una vez y el token queda en la caché cifrada de Windows. El admin SQL `trina_admin` queda solo como respaldo (contraseña fuera del repositorio).
- Firewall por IP de cliente + "Permitir servicios de Azure" (Power BI sin gateway).
- Reintento automático de conexión ante errores transitorios (40613 al reanudar, 258, …) y login timeout de 60 s.
- El código nunca crea ni elimina bases en Azure (un `CREATE DATABASE` sería de pago); `crear_base.py` solo aplica el esquema.
- Tests de integración: `ETL_ARENA_TEST_DB_URL` (base de pruebas aparte con "test" en el nombre, que se reinicia; o instancia local que se crea y borra).

**Medido:** esquema aplicado sin cambios de DDL; una corrida real de septiembre (244.748 filas) se guarda en ~46 s desde Chile.

**Pendiente:** probar la conexión desde la red de Trina (puerto 1433 saliente) y agregar la IP de la oficina al firewall; Power BI en modo *Import* (DirectQuery mantendría la base despierta).

## Actualización 2026-09-25 — Ambientes PROD y TEST/QA

**Decisión:** dos bases en el mismo servidor, con el mismo esquema, fijadas en el código (`src/etl_arena/ambientes.py`) para que cualquier PC funcione sin configurar URLs:

| Ambiente | Base | Selección |
|---|---|---|
| PROD | `trina_etl` | por defecto (`--entorno produccion`/`prod`) |
| TEST/QA | `trina_etl_prueba` | `--entorno prueba`/`qa`/`test` o `ETL_ARENA_ENTORNO` |

- `ETL_ARENA_DB_URL` / `ETL_ARENA_DB_URL_PRUEBA` quedan como reemplazo opcional (p. ej. instancia local); `ETL_ARENA_DB_AUTH` es `entra` por defecto; el `.env` se busca en la carpeta actual y luego en la raíz del proyecto.
- TEST/QA nunca puede resolver a la base de PROD; `run_lunes.py` en TEST/QA usa `data/work/_prueba`, copia (no mueve) el export y no avisa a Misael.
- **Incidente 2026-09-25:** con `ETL_ARENA_TEST_DB_URL` apuntando a `trina_etl_prueba`, los tests de integración la reiniciaron (el guard de `reiniciar_esquema` aceptaba "prueba" en el nombre) y se perdió la carga de validación. Corrección: `reiniciar_esquema` exige "test" en el nombre y rechaza siempre `trina_etl` y `trina_etl_prueba`; los tests de integración se omiten si su base es un ambiente; `verificar_entorno.py` lo reporta como FALTA.
- Vaciado deliberado de TEST/QA: `crear_base.py --entorno prueba --vaciar --confirmar trina_etl_prueba` (`esquema.vaciar_ambiente_prueba`: borra vistas y tablas, re-aplica el esquema, IDENTITY desde 1, conserva roles y usuarios). Solo acepta `trina_etl_prueba` y la confirmación exacta; con PROD se niega. Usado el 2026-09-25 para limpiar el incidente y recargar agosto (con matriz) y septiembre.

## Actualización 2026-09-25 — Cambios de esquema en PROD

- Columnas nuevas: `detencion.DuracionHoras` (junto a `DuracionSegundos`) y `etl_run.NumCorrida` (correlativo para personas, `IDENTITY_CACHE = OFF` para que no salte de a 1000 al reanudar la base serverless). Llegan por migración idempotente en `sql/02_corrida.sql`; `ALTER TABLE` las agrega al final de la tabla.
- PROD estaba vacía: se recreó solo `dbo.detencion` (autorizado por F. Barrientos) para dejar `DuracionHoras` físicamente junto a `DuracionSegundos`. En `etl_run` de PROD, `NumCorrida` quedó al final (recrearla arrastra casi todas las tablas); en las vistas aparece junto a `IdCorrida`. En TEST/QA, recreada completa, ambas quedan en su lugar.
- Se retiró `docker/mssql.compose.yml` (no se usaba).

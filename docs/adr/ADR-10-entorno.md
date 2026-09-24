# ADR-10 — Entorno de desarrollo

**Estado:** Aceptada · **Fecha:** 2026-09-24 · **Fuente:** `design.md` §Decisiones de diseño, `audit.md` · **Requisitos:** R11

## Contexto
F-29: el PC de oficina tenía solo Python 3.14, el driver ODBC legacy `SQL Server` (no maneja bien `DATETIME2` ni `fast_executemany`) y no tenía SQL Server local. El PC `FRANCISCO-PC` (2026-09-24) tiene Python 3.13, ODBC Driver 17 y 18, y SQL Server 2025 Developer (instancia `SQLSERVER2025DEV`).

## Opciones
- Python: fijar una versión, o `requires-python >= 3.13` (hay wheels de todas las dependencias en 3.13 y 3.14).
- SQL: contenedor Docker (`mcr.microsoft.com/mssql/server:2022`) o instancia local.
- Autenticación de desarrollo: `sa` con contraseña en `.env`, o autenticación Windows.

## Decisión
- `requires-python >= 3.13` (D-10). Una `.venv` por equipo.
- **ODBC Driver 18 for SQL Server** obligatorio.
- SQL de desarrollo: la instancia local si existe (`FRANCISCO-PC\SQLSERVER2025DEV`); si no, `docker/mssql.compose.yml`.
- Desarrollo con **autenticación Windows** (`trusted_connection=yes`): no hay contraseñas en `.env`. `sa` no se usa desde el código; la carga productiva usa el login de mínimo privilegio `etl_writer` (tarea 2.3).

## Consecuencias
- `ETL_ARENA_DB_URL` varía por equipo y vive solo en `.env` (no versionado).
- El DDL debe ser compatible con SQL Server 2022 y 2025.
- Instancia nombrada: si la conexión desde otro equipo falla, verificar el servicio *SQL Server Browser* o fijar el puerto.

-- 00_database.sql — crea la base (idempotente). Uso manual:
--   sqlcmd -S "FRANCISCO-PC\SQLSERVER2025DEV" -E -C -i sql\00_database.sql -v BaseDatos=ETL_Arena
-- El resto de los scripts (01..07) se ejecutan dentro de la base, sin USE: los aplica
-- etl_arena.persistence.esquema.aplicar_esquema() en el orden de ORDEN_SCRIPTS.
:setvar BaseDatos ETL_Arena
IF DB_ID(N'$(BaseDatos)') IS NULL
BEGIN
    DECLARE @sql NVARCHAR(200) = N'CREATE DATABASE ' + QUOTENAME(N'$(BaseDatos)');
    EXEC (@sql);
END
GO

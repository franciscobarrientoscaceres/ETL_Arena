-- 06_roles.sql — roles de base de datos (R11.1, R12.5, R18.4). Idempotente.
-- Los logins/usuarios concretos los crea el DBA fuera del repositorio (sin contraseñas aquí):
--   CREATE LOGIN etl_arena_carga WITH PASSWORD = '…';  CREATE USER etl_arena_carga FOR LOGIN etl_arena_carga;
--   ALTER ROLE etl_writer ADD MEMBER etl_arena_carga;
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

IF DATABASE_PRINCIPAL_ID(N'etl_writer') IS NULL CREATE ROLE etl_writer;
IF DATABASE_PRINCIPAL_ID(N'revisor') IS NULL CREATE ROLE revisor;
IF DATABASE_PRINCIPAL_ID(N'bi_reader') IS NULL CREATE ROLE bi_reader;
GO

-- etl_writer: append-only. INSERT y SELECT en dbo; UPDATE solo de la transición de estado de su
-- propio etl_run y del enlace de la carga mensual con su cierre; nunca DELETE (ADR-08).
GRANT SELECT, INSERT ON SCHEMA::dbo TO etl_writer;
GRANT UPDATE ON dbo.etl_run (Estado, FinalizadoEn, ResumenCalidad, MensajeError, IdReferenciaExcel, MinutosMuestreoDerivado) TO etl_writer;
GRANT UPDATE ON dbo.exclusion_matrix_carga (IdCorridaCierre) TO etl_writer;
DENY DELETE ON SCHEMA::dbo TO etl_writer;
GO

-- revisor: registra la revisión de detenciones (historial append-only) y lee las vistas.
GRANT INSERT ON dbo.detencion_revision TO revisor;
DENY DELETE, UPDATE ON dbo.detencion_revision TO revisor;
GO

-- bi_reader y revisor: SELECT solo sobre las vistas v_* (no sobre tablas base).
DECLARE @sql NVARCHAR(MAX) = N'';
SELECT @sql += N'GRANT SELECT ON ' + QUOTENAME(SCHEMA_NAME(schema_id)) + N'.' + QUOTENAME(name) + N' TO bi_reader, revisor;' + NCHAR(10)
FROM sys.views
WHERE SCHEMA_NAME(schema_id) = N'dbo' AND name LIKE N'v[_]%';
EXEC sys.sp_executesql @sql;
GO

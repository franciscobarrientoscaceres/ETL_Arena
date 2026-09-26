-- 06_roles.sql — roles de base de datos (esquema v2, ADR-12). Idempotente.
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

-- etl_writer: lee e inserta en dbo. Reemplaza (DELETE/UPDATE) solo las tablas de ESTADO vigente; las de
-- REGISTRO son solo-inserción, salvo la transición de estado de su etl_run y el enlace carga–cierre.
REVOKE DELETE ON SCHEMA::dbo FROM etl_writer;   -- quita el DENY a nivel de esquema del modelo v1 (ADR-08)
GRANT SELECT, INSERT ON SCHEMA::dbo TO etl_writer;
GRANT DELETE, UPDATE ON dbo.muestra_pcs TO etl_writer;
GRANT DELETE, UPDATE ON dbo.muestra_planta TO etl_writer;
GRANT DELETE, UPDATE ON dbo.detencion TO etl_writer;
GRANT DELETE, UPDATE ON dbo.disponibilidad_diaria TO etl_writer;
GRANT DELETE, UPDATE ON dbo.disponibilidad_mensual TO etl_writer;
GRANT DELETE, UPDATE ON dbo.calidad_dato TO etl_writer;
GRANT UPDATE ON dbo.etl_run (Estado, FinalizadoEn, ResumenCalidad, MensajeError, IdReferenciaExcel, MinutosMuestreoDerivado,
                             Publicada, MesesPublicados) TO etl_writer;
GRANT UPDATE ON dbo.exclusion_matrix_carga (IdCorridaCierre) TO etl_writer;
DENY DELETE ON dbo.etl_run TO etl_writer;
DENY DELETE ON dbo.correccion_dato TO etl_writer;
DENY DELETE, UPDATE ON dbo.excel_reference_run TO etl_writer;
DENY DELETE, UPDATE ON dbo.reconciliation_result TO etl_writer;
DENY DELETE ON dbo.exclusion_matrix_carga TO etl_writer;
-- Maestros (los mantiene el administrador con los seeds) y revisión humana (rol revisor): fuera de la carga.
DENY INSERT, UPDATE, DELETE ON dbo.proyecto TO etl_writer;
DENY INSERT, UPDATE, DELETE ON dbo.tipo_detencion TO etl_writer;
DENY INSERT, UPDATE, DELETE ON dbo.detencion_revision TO etl_writer;
GO

-- revisor: registra la revisión de detenciones (historial append-only).
GRANT INSERT ON dbo.detencion_revision TO revisor;
DENY DELETE, UPDATE ON dbo.detencion_revision TO revisor;
GO

-- bi_reader y revisor: leen las tablas de estado, el catálogo y las vistas v_* (Power BI, R18.4 rev. 3).
GRANT SELECT ON dbo.disponibilidad_mensual TO bi_reader, revisor;
GRANT SELECT ON dbo.disponibilidad_diaria TO bi_reader, revisor;
GRANT SELECT ON dbo.detencion TO bi_reader, revisor;
GRANT SELECT ON dbo.muestra_pcs TO bi_reader, revisor;
GRANT SELECT ON dbo.muestra_planta TO bi_reader, revisor;
GRANT SELECT ON dbo.calidad_dato TO bi_reader, revisor;
GRANT SELECT ON dbo.tipo_detencion TO bi_reader, revisor;
DECLARE @sql NVARCHAR(MAX) = N'';
SELECT @sql += N'GRANT SELECT ON ' + QUOTENAME(SCHEMA_NAME(schema_id)) + N'.' + QUOTENAME(name) + N' TO bi_reader, revisor;' + NCHAR(10)
FROM sys.views
WHERE SCHEMA_NAME(schema_id) = N'dbo' AND name LIKE N'v[_]%';
EXEC sys.sp_executesql @sql;
GO

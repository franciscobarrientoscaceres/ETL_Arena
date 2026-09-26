-- 02a_migracion_v2.sql — paso del esquema v1 (copias por IdCorrida, ADR-08) al v2 (estado vigente por mes,
-- ADR-12). Se aplica ANTES de 02_corrida.sql. Idempotente: en una base v2 (o nueva) no hace nada.
--
-- Solo migra si la base v1 está vacía (etl_run sin filas): elimina vistas y tablas v1 para que 02 cree las v2.
-- Si hay corridas, se detiene: TEST/QA se vacía con crear_base.py --vaciar; PROD requiere decisión explícita.
-- Se conservan: proyecto, tipo_detencion, detencion_revision y exclusion_matrix_carga (mismo formato en v2).
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

IF OBJECT_ID(N'dbo.fault_event', N'U') IS NOT NULL OR OBJECT_ID(N'dbo.raw_pcs_sample', N'U') IS NOT NULL
BEGIN
    IF OBJECT_ID(N'dbo.etl_run', N'U') IS NOT NULL AND EXISTS (SELECT 1 FROM dbo.etl_run)
        THROW 50001, N'Migración v2: la base tiene corridas en el esquema v1. TEST/QA: crear_base.py --entorno prueba --vaciar --confirmar trina_etl_prueba. PROD: migrar a mano (ADR-12).', 1;

    DECLARE @sql NVARCHAR(MAX) = N'';
    SELECT @sql += N'DROP VIEW ' + QUOTENAME(SCHEMA_NAME(schema_id)) + N'.' + QUOTENAME(name) + N';' + NCHAR(10)
    FROM sys.views WHERE is_ms_shipped = 0 AND SCHEMA_NAME(schema_id) = N'dbo';
    EXEC sys.sp_executesql @sql;

    SET @sql = N'';
    SELECT @sql += N'ALTER TABLE ' + QUOTENAME(SCHEMA_NAME(t.schema_id)) + N'.' + QUOTENAME(t.name)
                 + N' DROP CONSTRAINT ' + QUOTENAME(f.name) + N';' + NCHAR(10)
    FROM sys.foreign_keys AS f JOIN sys.tables AS t ON t.object_id = f.parent_object_id
    WHERE t.name IN (N'raw_pcs_column_map', N'raw_pcs_sample', N'plant_activity_sample', N'exclusion_matrix_sample',
                     N'availability_sample_result', N'availability_run_result', N'fault_event', N'fault_code_summary',
                     N'daily_availability', N'monthly_official_kpi', N'annual_availability', N'detencion',
                     N'data_quality_issue', N'excel_reference_sample', N'excel_reference_fault_event',
                     N'excel_reference_daily', N'reconciliation_result', N'correccion_dato', N'excel_reference_run',
                     N'etl_run');
    EXEC sys.sp_executesql @sql;

    SET @sql = N'';
    SELECT @sql += N'DROP TABLE dbo.' + QUOTENAME(name) + N';' + NCHAR(10)
    FROM sys.tables
    WHERE SCHEMA_NAME(schema_id) = N'dbo'
      AND name IN (N'raw_pcs_column_map', N'raw_pcs_sample', N'plant_activity_sample', N'exclusion_matrix_sample',
                   N'availability_sample_result', N'availability_run_result', N'fault_event', N'fault_code_summary',
                   N'daily_availability', N'monthly_official_kpi', N'annual_availability', N'detencion',
                   N'data_quality_issue', N'excel_reference_sample', N'excel_reference_fault_event',
                   N'excel_reference_daily', N'reconciliation_result', N'correccion_dato', N'excel_reference_run',
                   N'etl_run');
    EXEC sys.sp_executesql @sql;
    PRINT N'Migración v2: tablas y vistas v1 eliminadas (la base no tenía corridas).';
END
GO

-- 03_indices.sql — índices no clustered (los columnstore van con sus tablas en 02). Idempotente.
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_etl_run_vigencia' AND object_id = OBJECT_ID(N'dbo.etl_run'))
    CREATE INDEX IX_etl_run_vigencia ON dbo.etl_run (IdProyecto, FinPeriodo, EsOficial, Estado)
        INCLUDE (EstadoExclusiones, TipoCorrida, IniciadoEn);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_detencion_negocio' AND object_id = OBJECT_ID(N'dbo.detencion'))
    CREATE INDEX IX_detencion_negocio ON dbo.detencion (IdProyecto, NumeroPCS, FechaInicio);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_detencion_corrida' AND object_id = OBJECT_ID(N'dbo.detencion'))
    CREATE INDEX IX_detencion_corrida ON dbo.detencion (IdCorrida);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_detencion_revision_clave' AND object_id = OBJECT_ID(N'dbo.detencion_revision'))
    CREATE INDEX IX_detencion_revision_clave ON dbo.detencion_revision (IdProyecto, NumeroPCS, FechaInicio, RevisadoEn DESC);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_dqi_corrida_tipo' AND object_id = OBJECT_ID(N'dbo.data_quality_issue'))
    CREATE INDEX IX_dqi_corrida_tipo ON dbo.data_quality_issue (IdCorrida, Tipo);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_mok_vigente' AND object_id = OBJECT_ID(N'dbo.monthly_official_kpi'))
    CREATE INDEX IX_mok_vigente ON dbo.monthly_official_kpi (IdProyecto, Anio, Mes, RegistradoEn DESC);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_emc_mes' AND object_id = OBJECT_ID(N'dbo.exclusion_matrix_carga'))
    CREATE INDEX IX_emc_mes ON dbo.exclusion_matrix_carga (IdProyecto, Anio, Mes, CargadoEn DESC);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_recon_corrida' AND object_id = OBJECT_ID(N'dbo.reconciliation_result'))
    CREATE INDEX IX_recon_corrida ON dbo.reconciliation_result (IdCorrida, Nivel);
GO

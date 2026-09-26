-- 03_indices.sql — índices de apoyo (esquema v2, ADR-12). Idempotente.
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

-- Reemplazo por mes (DELETE … WHERE IdProyecto, Anio, Mes) y lecturas de Power BI por mes.
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_muestra_pcs_mes' AND object_id = OBJECT_ID(N'dbo.muestra_pcs'))
    CREATE INDEX IX_muestra_pcs_mes ON dbo.muestra_pcs (IdProyecto, Anio, Mes);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_muestra_planta_mes' AND object_id = OBJECT_ID(N'dbo.muestra_planta'))
    CREATE INDEX IX_muestra_planta_mes ON dbo.muestra_planta (IdProyecto, Anio, Mes);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ddiaria_mes' AND object_id = OBJECT_ID(N'dbo.disponibilidad_diaria'))
    CREATE INDEX IX_ddiaria_mes ON dbo.disponibilidad_diaria (IdProyecto, Anio, Mes);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_calidad_mes' AND object_id = OBJECT_ID(N'dbo.calidad_dato'))
    CREATE INDEX IX_calidad_mes ON dbo.calidad_dato (IdProyecto, Anio, Mes, Tipo);
GO
-- Detenciones: clave de negocio para revisiones (PCS + inicio) y búsqueda por mes.
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_detencion_revision_negocio' AND object_id = OBJECT_ID(N'dbo.detencion'))
    CREATE INDEX IX_detencion_revision_negocio ON dbo.detencion (IdProyecto, NumeroPCS, FechaInicio);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_detencion_revision_clave' AND object_id = OBJECT_ID(N'dbo.detencion_revision'))
    CREATE INDEX IX_detencion_revision_clave ON dbo.detencion_revision (IdProyecto, NumeroPCS, FechaInicio, RevisadoEn DESC);
GO
-- Registro.
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_etl_run_proyecto' AND object_id = OBJECT_ID(N'dbo.etl_run'))
    CREATE INDEX IX_etl_run_proyecto ON dbo.etl_run (IdProyecto, InicioPeriodo, Estado) INCLUDE (NumCorrida, Publicada);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_emc_mes' AND object_id = OBJECT_ID(N'dbo.exclusion_matrix_carga'))
    CREATE INDEX IX_emc_mes ON dbo.exclusion_matrix_carga (IdProyecto, Anio, Mes, CargadoEn DESC);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_recon_corrida' AND object_id = OBJECT_ID(N'dbo.reconciliation_result'))
    CREATE INDEX IX_recon_corrida ON dbo.reconciliation_result (IdCorrida, Nivel);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_correccion_corrida' AND object_id = OBJECT_ID(N'dbo.correccion_dato'))
    CREATE INDEX IX_correccion_corrida ON dbo.correccion_dato (IdCorrida);
GO

-- 01_maestros.sql — proyecto y catálogo de detenciones (R12). Idempotente.
-- Convención: fechas y marcas de tiempo en DATETIME (no DATE ni DATETIME2).
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

IF OBJECT_ID(N'dbo.proyecto', N'U') IS NULL
CREATE TABLE dbo.proyecto (
    IdProyecto          INT            NOT NULL CONSTRAINT PK_proyecto PRIMARY KEY,
    Nombre              NVARCHAR(100)  NOT NULL,
    Estado              NVARCHAR(20)   NOT NULL CONSTRAINT CK_proyecto_estado CHECK (Estado IN (N'en_ejecucion', N'por_implementar')),
    FechaInicio         DATETIME       NULL,
    NumPCS              INT            NULL,
    NumBateriasPorPCS   INT            NULL,
    NumRacksPorBAC      INT            NULL,
    TotalRacks          AS (NumPCS * NumBateriasPorPCS * NumRacksPorBAC) PERSISTED,
    MinutosMuestreo     INT            NULL,
    ZonaHoraria         NVARCHAR(50)   NULL,
    Descripcion         NVARCHAR(500)  NULL
);
GO

-- Seed generado desde PCS-Fault por scripts/generar_seed_tipo_detencion.py (F-19, F-35, D-14).
IF OBJECT_ID(N'dbo.tipo_detencion', N'U') IS NULL
CREATE TABLE dbo.tipo_detencion (
    IdTipoDetencion     INT            NOT NULL CONSTRAINT PK_tipo_detencion PRIMARY KEY,  -- Code Number
    CodigoFalla         NVARCHAR(10)   NOT NULL CONSTRAINT UQ_tipo_detencion_codigo UNIQUE, -- 'F55'
    DescripcionFallaPE  NVARCHAR(150)  NOT NULL,
    CodigoDescripcion   NVARCHAR(200)  NOT NULL,
    Significado         NVARCHAR(50)   NULL,     -- 'Crítico', 'Parcial', 'No fault', ...
    Operativo           NVARCHAR(10)   NULL      -- 'Yes' / 'No' / NULL
);
GO

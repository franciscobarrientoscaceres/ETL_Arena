-- 02_corrida.sql — tablas de corrida, resultados, detenciones, calidad, referencia Excel y
-- reconciliación (R11–R15, design.md §Data Models). Idempotente. Append-only por IdCorrida (ADR-08).
-- Convención: fechas y marcas de tiempo en DATETIME (no DATE ni DATETIME2).
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

-- ============================================================ corrida
IF OBJECT_ID(N'dbo.etl_run', N'U') IS NULL
CREATE TABLE dbo.etl_run (
    IdCorrida                      UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_etl_run PRIMARY KEY,
    IdProyecto                     INT            NOT NULL CONSTRAINT FK_etl_run_proyecto REFERENCES dbo.proyecto(IdProyecto),
    TipoCorrida                    NVARCHAR(20)   NOT NULL CONSTRAINT CK_etl_run_tipo CHECK (TipoCorrida IN (N'semanal', N'cierre_mensual', N'reproceso', N'golden')),
    EsOficial                      BIT            NOT NULL CONSTRAINT DF_etl_run_oficial DEFAULT 0,
    EstadoExclusiones              NVARCHAR(20)   NOT NULL CONSTRAINT CK_etl_run_exclusiones CHECK (EstadoExclusiones IN (N'sin_exclusiones', N'con_exclusiones')),  -- D-17: ambos oficiales
    ArchivoOrigen                  NVARCHAR(500)  NOT NULL,
    HashArchivoOrigen              CHAR(64)       NOT NULL,
    SistemaOrigen                  NVARCHAR(50)   NOT NULL CONSTRAINT DF_etl_run_sistema DEFAULT N'scada_export',
    IdReferenciaExcel              UNIQUEIDENTIFIER NULL,
    IniciadoEn                     DATETIME       NOT NULL,
    FinalizadoEn                   DATETIME       NULL,
    InicioPeriodo                  DATETIME       NOT NULL,   -- C5
    FinPeriodo                     DATETIME       NOT NULL,   -- C7
    InicioPeriodoEventos           DATETIME       NOT NULL,   -- L2
    FinPeriodoEventos              DATETIME       NOT NULL,   -- L4
    FinDiario                      DATETIME       NOT NULL,   -- última Daily!C (F-33)
    SoloTiempoOperacional          BIT            NOT NULL,   -- C21
    AplicarEventoExcusable         BIT            NOT NULL,   -- C31 → Exclusion_Matrix (F-37)
    AplicarEventoExcusableEventos  BIT            NOT NULL,   -- L14 → Exclusion_Matrix (F-37, D-16)
    ModoHuecos                     NVARCHAR(10)   NOT NULL,
    TotalPCS                       INT            NOT NULL,
    BateriasPorPCS                 INT            NOT NULL,
    RacksPorPCS                    INT            NOT NULL,
    TotalRacks                     INT            NOT NULL,
    MinutosMuestreo                INT            NOT NULL,
    MinutosMuestreoDerivado        FLOAT          NULL,       -- C23
    Estado                         NVARCHAR(20)   NOT NULL CONSTRAINT CK_etl_run_estado CHECK (Estado IN (N'running', N'success', N'failed', N'parity_failed')),
    MensajeError                   NVARCHAR(MAX)  NULL,
    ResumenCalidad                 NVARCHAR(MAX)  NULL CONSTRAINT CK_etl_run_resumen CHECK (ResumenCalidad IS NULL OR ISJSON(ResumenCalidad) = 1),
    VersionAlgoritmo               NVARCHAR(100)  NOT NULL
);
GO

IF OBJECT_ID(N'dbo.raw_pcs_column_map', N'U') IS NULL
CREATE TABLE dbo.raw_pcs_column_map (      -- trazabilidad de columnas, una vez por corrida
    IdCorrida       UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_colmap_run REFERENCES dbo.etl_run(IdCorrida),
    NumeroPCS       INT            NOT NULL,
    Campo           NVARCHAR(20)   NOT NULL,     -- falla | estado | advertencia | modulos
    NumeroColumna   INT            NOT NULL,
    NombreColumna   NVARCHAR(200)  NOT NULL,
    CONSTRAINT PK_raw_pcs_column_map PRIMARY KEY (IdCorrida, NumeroPCS, Campo)
);
GO

-- ============================================================ staging (columnstore, ADR-09)
IF OBJECT_ID(N'dbo.raw_pcs_sample', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.raw_pcs_sample (      -- filas del período + la fila anterior (fallback F-04); D-09
        IdCorrida               UNIQUEIDENTIFIER NOT NULL,
        NumeroFilaOrigen        INT            NOT NULL,
        NumeroPCS               INT            NOT NULL,
        SerialFechaExcelOrigen  FLOAT          NOT NULL,
        MarcaTiempoLocalOrigen  DATETIME       NULL,       -- NULL solo si A2 estaba vacía
        FallaRaw                NVARCHAR(255)  NULL,
        FallaRawEsNumero        BIT            NOT NULL,
        EstadoRaw               NVARCHAR(255)  NULL,
        AdvertenciaRaw          NVARCHAR(255)  NULL,
        ModulosRaw              FLOAT          NULL,
        ModulosDisponibles      FLOAT          NOT NULL,   -- = baterias_por_pcs si nulo
        ModulosDisponiblesNulo  BIT            NOT NULL,
        EsFilaNuevaDelExport    BIT            NOT NULL    -- aportada por el export incremental de esta corrida (D-07)
    );
    CREATE CLUSTERED COLUMNSTORE INDEX CCI_raw_pcs_sample ON dbo.raw_pcs_sample;
END
GO

IF OBJECT_ID(N'dbo.plant_activity_sample', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.plant_activity_sample (
        IdCorrida                       UNIQUEIDENTIFIER NOT NULL,
        NumeroFilaOrigen                INT            NOT NULL,
        SerialFecha                     FLOAT          NULL,      -- puede faltar (F-31)
        MarcaTiempoMuestra              DATETIME       NULL,
        FactorOperacionalRaw            FLOAT          NULL,      -- col C
        EventoExcusadoRaw               FLOAT          NULL,      -- col D: espejo de Calc!BO, no pondera (F-37)
        SetpointPotenciaActivaKW        FLOAT          NULL,
        DroopSobrefrecuenciaHabilitado  FLOAT          NULL,
        DroopBajafrecuenciaHabilitado   FLOAT          NULL,
        PotenciaActivaPOIKW             FLOAT          NULL,
        PorcentajeSOC                   FLOAT          NULL
    );
    CREATE CLUSTERED COLUMNSTORE INDEX CCI_plant_activity_sample ON dbo.plant_activity_sample;
END
GO

IF OBJECT_ID(N'dbo.exclusion_matrix_carga', N'U') IS NULL
CREATE TABLE dbo.exclusion_matrix_carga (   -- D-17: una fila por entrega mensual (append-only; vigente = última)
    IdCarga                 BIGINT IDENTITY CONSTRAINT PK_exclusion_matrix_carga PRIMARY KEY,
    IdProyecto              INT            NOT NULL CONSTRAINT FK_emc_proyecto REFERENCES dbo.proyecto(IdProyecto),
    Anio                    INT            NOT NULL,
    Mes                     INT            NOT NULL CONSTRAINT CK_emc_mes CHECK (Mes BETWEEN 1 AND 12),
    ArchivoOrigen           NVARCHAR(500)  NOT NULL,
    Sha256Archivo           CHAR(64)       NOT NULL,
    CargadoEn               DATETIME       NOT NULL CONSTRAINT DF_emc_cargado DEFAULT GETDATE(),
    IdCorridaCierre         UNIQUEIDENTIFIER NULL     -- cierre_mensual con_exclusiones que la usó
);
GO

IF OBJECT_ID(N'dbo.exclusion_matrix_sample', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.exclusion_matrix_sample (  -- F-37: una fila por (fila, PCS) con valor ≠ 0
        IdCorrida               UNIQUEIDENTIFIER NOT NULL,
        NumeroFilaOrigen        INT            NOT NULL,
        NumeroPCS               INT            NOT NULL,
        SerialFecha             FLOAT          NULL,
        MarcaTiempoMuestra      DATETIME       NULL,
        ValorExclusion          TINYINT        NOT NULL CONSTRAINT CK_ems_valor CHECK (ValorExclusion IN (1, 2)),
        BateriasPrevias         FLOAT          NULL,      -- solo valor 2
        EventoExcusadoFila      FLOAT          NULL,      -- columna "Excused Event"
        Comentario              NVARCHAR(500)  NULL       -- columna "Comments" (causa del EE)
    );
    CREATE CLUSTERED COLUMNSTORE INDEX CCI_exclusion_matrix_sample ON dbo.exclusion_matrix_sample;
END
GO

IF OBJECT_ID(N'dbo.correccion_dato', N'U') IS NULL
CREATE TABLE dbo.correccion_dato (          -- append-only; celdas cambiadas (D-13, D-17)
    IdCorreccion            BIGINT IDENTITY CONSTRAINT PK_correccion_dato PRIMARY KEY,
    IdCorrida               UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_correccion_run REFERENCES dbo.etl_run(IdCorrida),
    TipoCorreccion          NVARCHAR(20)   NOT NULL CONSTRAINT CK_correccion_tipo CHECK (TipoCorreccion IN (N'reproceso_raw', N'exclusion_matrix', N'plant_activity')),
    Hoja                    NVARCHAR(40)   NOT NULL,   -- 'RawData-PCS' | 'Exclusion_Matrix' | 'PlantActivity'
    NumeroFilaOrigen        INT            NOT NULL,
    SerialFechaExcelOrigen  FLOAT          NOT NULL,
    MarcaTiempoLocalOrigen  DATETIME       NOT NULL,
    NumeroPCS               INT            NULL,       -- NULL en columnas por fila
    Campo                   NVARCHAR(40)   NOT NULL,   -- FAULT | STATUS | WARNING | MODULES | PCSnn | C | D | …
    ValorAnterior           NVARCHAR(255)  NULL,
    ValorNuevo              NVARCHAR(255)  NULL,
    ArchivoOrigen           NVARCHAR(260)  NOT NULL,
    Sha256Archivo           CHAR(64)       NOT NULL
);
GO

-- ============================================================ resultados
IF OBJECT_ID(N'dbo.availability_sample_result', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.availability_sample_result (   -- solo filas dentro del período KPI
        IdCorrida                        UNIQUEIDENTIFIER NOT NULL,
        NumeroFilaOrigen                 INT      NOT NULL,
        NumeroPCS                        INT      NOT NULL,
        SerialFecha                      FLOAT    NOT NULL,
        MarcaTiempoMuestra               DATETIME NOT NULL,
        ModulosDisponibles               FLOAT    NOT NULL,
        ModulosDisponiblesNulo           BIT      NOT NULL,
        BateriasIndisponibles            FLOAT    NOT NULL,
        ValorExclusion                   TINYINT  NOT NULL,   -- Exclusion_Matrix 0/1/2 (F-37)
        FactorOperacional                FLOAT    NOT NULL,
        BateriasIndisponiblesPonderadas  FLOAT    NOT NULL,
        ImpactoRackPonderado             FLOAT    NOT NULL
    );
    CREATE CLUSTERED COLUMNSTORE INDEX CCI_availability_sample_result ON dbo.availability_sample_result;
END
GO

IF OBJECT_ID(N'dbo.availability_run_result', N'U') IS NULL
CREATE TABLE dbo.availability_run_result (
    IdCorrida                     UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_availability_run_result PRIMARY KEY
                                  CONSTRAINT FK_arr_run REFERENCES dbo.etl_run(IdCorrida),
    BloquesMuestreo               INT    NOT NULL,   -- C12
    TotalRacks                    INT    NOT NULL,   -- C11
    BloquesRacksIndisponibles     FLOAT  NOT NULL,   -- C14
    DisponibilidadPeriodo         FLOAT  NULL,       -- C16
    DisponibilidadAnualAcumulada  FLOAT  NULL,       -- C19
    MinutosMuestreoDerivado       FLOAT  NULL,       -- C23
    HorasRackEventos              FLOAT  NOT NULL,   -- ListOfFaults!L10
    BloquesMuestreoCalendario     FLOAT  NULL        -- (C7-C5+1)*24*60/C23 (informativo)
);
GO

IF OBJECT_ID(N'dbo.fault_event', N'U') IS NULL
CREATE TABLE dbo.fault_event (
    IdCorrida                     UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_fault_event_run REFERENCES dbo.etl_run(IdCorrida),
    OrdenExcel                    INT            NOT NULL,     -- fila de ListOfFaults - 5
    NumeroPCS                     INT            NOT NULL,
    SerialInicio                  FLOAT          NOT NULL,
    SerialFin                     FLOAT          NOT NULL,
    MarcaTiempoInicio             DATETIME       NOT NULL,
    MarcaTiempoFin                DATETIME       NOT NULL,
    DuracionHoras                 FLOAT          NOT NULL,
    CodigoFalla                   NVARCHAR(300)  NOT NULL,     -- "F" & descripción puede ser largo
    DescripcionFalla              NVARCHAR(255)  NOT NULL,
    DescripcionFallaFallback      BIT            NOT NULL,
    NumeroBloques                 INT            NOT NULL,
    SumaBloques                   FLOAT          NOT NULL,
    PromedioBateriasInvolucradas  FLOAT          NOT NULL,
    HorasRackIndisponibles        FLOAT          NOT NULL,
    EventoArrastradoExcel         BIT            NOT NULL,
    ExcelHabriaFallado            BIT            NOT NULL,
    CerradoPorModulosNulo         BIT            NOT NULL,
    TieneExclusion                BIT            NOT NULL,     -- algún bloque con Exclusion_Matrix ≠ 0 (F-37)
    CONSTRAINT PK_fault_event PRIMARY KEY (IdCorrida, OrdenExcel)
);
GO

IF OBJECT_ID(N'dbo.fault_code_summary', N'U') IS NULL
CREATE TABLE dbo.fault_code_summary (      -- ListOfFaults!N:Q (166 filas: F230-F232 repetidos, F-35)
    IdCorrida              UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_fcs_run REFERENCES dbo.etl_run(IdCorrida),
    Ranking                INT            NOT NULL,
    CodigoFalla            NVARCHAR(10)   NOT NULL,
    DescripcionFallaPE     NVARCHAR(150)  NULL,
    HorasRackIndisponibles FLOAT          NOT NULL,
    Porcentaje             FLOAT          NULL,
    CONSTRAINT PK_fault_code_summary PRIMARY KEY (IdCorrida, Ranking)
);
GO

IF OBJECT_ID(N'dbo.daily_availability', N'U') IS NULL
CREATE TABLE dbo.daily_availability (
    IdCorrida                            UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_daily_run REFERENCES dbo.etl_run(IdCorrida),
    DiaN                                 INT      NOT NULL,
    Dia                                  DATETIME NOT NULL,
    BloquesRacksIndisponiblesDiarios     FLOAT    NOT NULL,
    BloquesRacksIndisponiblesAcumulados  FLOAT    NOT NULL,
    Disponibilidad                       FLOAT    NULL,
    Variacion                            FLOAT    NULL,
    CONSTRAINT PK_daily_availability PRIMARY KEY (IdCorrida, DiaN)
);
GO

IF OBJECT_ID(N'dbo.monthly_official_kpi', N'U') IS NULL
CREATE TABLE dbo.monthly_official_kpi (    -- append-only; vigente = último RegistradoEn
    IdKpiMensual               BIGINT IDENTITY CONSTRAINT PK_monthly_official_kpi PRIMARY KEY,
    IdProyecto                 INT            NOT NULL CONSTRAINT FK_mok_proyecto REFERENCES dbo.proyecto(IdProyecto),
    Anio                       INT            NOT NULL,
    Mes                        INT            NOT NULL CONSTRAINT CK_mok_mes CHECK (Mes BETWEEN 1 AND 12),
    DiasMes                    FLOAT          NOT NULL,
    BloquesMuestreo            FLOAT          NOT NULL,   -- C12 (intervalos existentes, D-07)
    BloquesRacksIndisponibles  FLOAT          NOT NULL,
    DisponibilidadContractual  FLOAT          NOT NULL CONSTRAINT DF_mok_contractual DEFAULT 0.98,
    Origen                     NVARCHAR(20)   NOT NULL CONSTRAINT CK_mok_origen CHECK (Origen IN (N'corrida', N'excel_manual')),
    IdCorrida                  UNIQUEIDENTIFIER NULL CONSTRAINT FK_mok_run REFERENCES dbo.etl_run(IdCorrida),
    RegistradoEn               DATETIME       NOT NULL CONSTRAINT DF_mok_registrado DEFAULT GETDATE()
);
GO

IF OBJECT_ID(N'dbo.annual_availability', N'U') IS NULL
CREATE TABLE dbo.annual_availability (     -- snapshot Annual_AVA calculado en cada corrida
    IdCorrida                       UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_annual_run REFERENCES dbo.etl_run(IdCorrida),
    Anio                            INT   NOT NULL,
    Mes                             INT   NOT NULL,
    DiasMes                         FLOAT NOT NULL,
    BloquesMuestreo                 FLOAT NOT NULL,
    BloquesRacksIndisponibles       FLOAT NULL,
    DisponibilidadMensual           FLOAT NULL,
    BloquesMuestreoAcumulados       FLOAT NOT NULL,
    BloquesIndisponiblesAcumulados  FLOAT NOT NULL,
    DisponibilidadAcumulada         FLOAT NULL,
    DisponibilidadContractual       FLOAT NULL,
    CONSTRAINT PK_annual_availability PRIMARY KEY (IdCorrida, Anio, Mes)
);
GO

-- ============================================================ detenciones (R12)
IF OBJECT_ID(N'dbo.detencion', N'U') IS NULL
CREATE TABLE dbo.detencion (               -- append-only por corrida
    IdDetencion                   BIGINT IDENTITY CONSTRAINT PK_detencion PRIMARY KEY,
    IdProyecto                    INT            NOT NULL CONSTRAINT FK_detencion_proyecto REFERENCES dbo.proyecto(IdProyecto),
    IdCorrida                     UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_detencion_run REFERENCES dbo.etl_run(IdCorrida),
    OrdenExcel                    INT            NOT NULL,
    NumeroPCS                     INT            NOT NULL,
    FechaInicio                   DATETIME       NOT NULL,
    FechaTermino                  DATETIME       NOT NULL,
    DuracionSegundos              INT            NOT NULL,   -- ROUND(DuracionHoras*3600)
    IdTipoDetencion               INT            NULL CONSTRAINT FK_detencion_tipo REFERENCES dbo.tipo_detencion(IdTipoDetencion),
    CodigoFalla                   NVARCHAR(300)  NOT NULL,
    DescripcionFalla              NVARCHAR(255)  NOT NULL,
    DescripcionFallaFallback      BIT            NOT NULL,
    PromedioBateriasInvolucradas  FLOAT          NOT NULL,
    HorasRackIndisponibles        FLOAT          NOT NULL,
    CerradoPorModulosNulo         BIT            NOT NULL,   -- terminó porque la fila siguiente estaba vacía
    EsExcusable                   BIT            NOT NULL,   -- algún bloque con Exclusion_Matrix ≠ 0 (F-37)
    EventoArrastradoExcel         BIT            NOT NULL
);
GO

IF OBJECT_ID(N'dbo.detencion_revision', N'U') IS NULL
CREATE TABLE dbo.detencion_revision (      -- workflow; historial append-only (F-27)
    IdRevision       BIGINT IDENTITY CONSTRAINT PK_detencion_revision PRIMARY KEY,
    IdProyecto       INT            NOT NULL CONSTRAINT FK_revision_proyecto REFERENCES dbo.proyecto(IdProyecto),
    NumeroPCS        INT            NOT NULL,
    FechaInicio      DATETIME       NOT NULL,
    EstadoRevision   NVARCHAR(20)   NOT NULL CONSTRAINT CK_revision_estado CHECK (EstadoRevision IN (N'pendiente', N'revisado', N'excluido')),
    Observacion      NVARCHAR(500)  NULL,
    RevisadoPor      NVARCHAR(100)  NOT NULL,
    RevisadoEn       DATETIME       NOT NULL CONSTRAINT DF_revision_en DEFAULT GETDATE()
);
GO

-- ============================================================ calidad, referencia Excel, reconciliación
IF OBJECT_ID(N'dbo.data_quality_issue', N'U') IS NULL
CREATE TABLE dbo.data_quality_issue (
    IdCorrida        UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_dqi_run REFERENCES dbo.etl_run(IdCorrida),
    Tipo             NVARCHAR(50)   NOT NULL,
    Severidad        NVARCHAR(12)   NOT NULL CONSTRAINT CK_dqi_severidad CHECK (Severidad IN (N'info', N'advertencia', N'error')),
    NumeroFilaOrigen INT            NULL,
    NumeroPCS        INT            NULL,
    SerialFecha      FLOAT          NULL,
    Detalle          NVARCHAR(1000) NULL
);
GO

IF OBJECT_ID(N'dbo.excel_reference_run', N'U') IS NULL
CREATE TABLE dbo.excel_reference_run (
    IdReferencia      UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_excel_reference_run PRIMARY KEY,
    Corte             NVARCHAR(50)   NOT NULL,
    ArchivoLibro      NVARCHAR(500)  NOT NULL,
    HashLibro         CHAR(64)       NOT NULL,
    ExtraidoEn        DATETIME       NOT NULL,
    C5 DATETIME NOT NULL, C7 DATETIME NOT NULL, C21 NVARCHAR(10) NULL, C31 NVARCHAR(10) NULL,
    L2 DATETIME NOT NULL, L4 DATETIME NOT NULL, L14 NVARCHAR(10) NULL, DailyD5 DATETIME NULL,
    C12 INT NULL, C14 FLOAT NULL, C16 FLOAT NULL, C19 FLOAT NULL, C23 FLOAT NULL, L10 FLOAT NULL
);
GO

IF OBJECT_ID(N'dbo.excel_reference_sample', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.excel_reference_sample (  -- Calc!E4:BO, solo celdas no vacías
        IdReferencia  UNIQUEIDENTIFIER NOT NULL,
        FilaResultado INT   NOT NULL,
        SerialFecha   FLOAT NOT NULL,
        NumeroPCS     INT   NOT NULL,
        Valor         FLOAT NOT NULL
    );
    CREATE CLUSTERED COLUMNSTORE INDEX CCI_excel_reference_sample ON dbo.excel_reference_sample;
END
GO

IF OBJECT_ID(N'dbo.excel_reference_fault_event', N'U') IS NULL
CREATE TABLE dbo.excel_reference_fault_event (
    IdReferencia UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_erfe_ref REFERENCES dbo.excel_reference_run(IdReferencia),
    OrdenExcel INT NOT NULL, B FLOAT NULL, C FLOAT NULL, D FLOAT NULL, E FLOAT NULL,
    F NVARCHAR(300) NULL, G NVARCHAR(255) NULL, H FLOAT NULL, I FLOAT NULL,
    CONSTRAINT PK_excel_reference_fault_event PRIMARY KEY (IdReferencia, OrdenExcel)
);
GO

IF OBJECT_ID(N'dbo.excel_reference_daily', N'U') IS NULL
CREATE TABLE dbo.excel_reference_daily (
    IdReferencia UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_erd_ref REFERENCES dbo.excel_reference_run(IdReferencia),
    DiaN INT NOT NULL, Dia DATETIME NULL, D FLOAT NULL, E FLOAT NULL, F FLOAT NULL, G FLOAT NULL,
    CONSTRAINT PK_excel_reference_daily PRIMARY KEY (IdReferencia, DiaN)
);
GO

IF OBJECT_ID(N'dbo.reconciliation_result', N'U') IS NULL
CREATE TABLE dbo.reconciliation_result (
    IdCorrida     UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_recon_run REFERENCES dbo.etl_run(IdCorrida),
    IdReferencia  UNIQUEIDENTIFIER NULL CONSTRAINT FK_recon_ref REFERENCES dbo.excel_reference_run(IdReferencia),
    Nivel         INT            NOT NULL,          -- 1..5; 9 = invariante
    Metrica       NVARCHAR(100)  NOT NULL,
    Clave         NVARCHAR(200)  NULL,
    ValorPython   NVARCHAR(100)  NULL,
    ValorExcel    NVARCHAR(100)  NULL,
    Delta         FLOAT          NULL,
    Tolerancia    FLOAT          NULL,
    Aprobado      BIT            NOT NULL
);
GO

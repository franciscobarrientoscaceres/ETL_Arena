-- 02_corrida.sql — esquema v2 (ADR-12): tablas de REGISTRO (solo inserción) y tablas de ESTADO VIGENTE
-- (una fila por dato, reemplazadas mes a mes al publicar). Idempotente. Requisitos R10–R15 (rev. 3).
-- Convención: fechas y marcas de tiempo en DATETIME (no DATE ni DATETIME2).
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

-- Correlativos sin saltos: por defecto SQL Server/Azure reserva IDENTITY en bloques de 1000 y los pierde
-- al reiniciarse (una base serverless se pausa y reanuda a diario). NumCorrida debe ser 1, 2, 3…
ALTER DATABASE SCOPED CONFIGURATION SET IDENTITY_CACHE = OFF;
GO

-- ============================================================ REGISTRO (append-only)
IF OBJECT_ID(N'dbo.etl_run', N'U') IS NULL
CREATE TABLE dbo.etl_run (                 -- una fila por ejecución
    IdCorrida                      UNIQUEIDENTIFIER NOT NULL CONSTRAINT PK_etl_run PRIMARY KEY,
    NumCorrida                     INT IDENTITY(1,1) NOT NULL CONSTRAINT UQ_etl_run_NumCorrida UNIQUE,  -- 1, 2, 3… para personas
    IdProyecto                     INT            NOT NULL CONSTRAINT FK_etl_run_proyecto REFERENCES dbo.proyecto(IdProyecto),
    TipoCorrida                    NVARCHAR(20)   NOT NULL CONSTRAINT CK_etl_run_tipo CHECK (TipoCorrida IN (N'semanal', N'cierre_mensual', N'reproceso', N'golden')),
    EstadoExclusiones              NVARCHAR(20)   NOT NULL CONSTRAINT CK_etl_run_exclusiones CHECK (EstadoExclusiones IN (N'sin_exclusiones', N'con_exclusiones')),
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
    Publicada                      BIT            NOT NULL CONSTRAINT DF_etl_run_publicada DEFAULT 0,  -- reemplazó el estado vigente (D-22)
    MesesPublicados                NVARCHAR(100)  NULL,       -- p. ej. '2026-09'
    MensajeError                   NVARCHAR(MAX)  NULL,
    ResumenCalidad                 NVARCHAR(MAX)  NULL CONSTRAINT CK_etl_run_resumen CHECK (ResumenCalidad IS NULL OR ISJSON(ResumenCalidad) = 1),
    VersionAlgoritmo               NVARCHAR(100)  NOT NULL
);
GO

IF OBJECT_ID(N'dbo.exclusion_matrix_carga', N'U') IS NULL
CREATE TABLE dbo.exclusion_matrix_carga (   -- D-17: una fila por entrega mensual de Alex
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

IF OBJECT_ID(N'dbo.correccion_dato', N'U') IS NULL
CREATE TABLE dbo.correccion_dato (          -- cada dato que cambió al recargar o al cargar la matriz (D-13, D-24)
    IdCorreccion            BIGINT IDENTITY CONSTRAINT PK_correccion_dato PRIMARY KEY,
    IdCorrida               UNIQUEIDENTIFIER NOT NULL CONSTRAINT FK_correccion_run REFERENCES dbo.etl_run(IdCorrida),
    TipoCorreccion          NVARCHAR(20)   NOT NULL CONSTRAINT CK_correccion_tipo CHECK (TipoCorreccion IN (N'recarga', N'reproceso_raw', N'exclusion_matrix', N'plant_activity')),
    Hoja                    NVARCHAR(40)   NOT NULL,   -- 'RawData-PCS' | 'Exclusion_Matrix' | 'PlantActivity'
    NumeroFilaOrigen        INT            NOT NULL,
    SerialFechaExcelOrigen  FLOAT          NOT NULL,
    MarcaTiempoLocalOrigen  DATETIME       NOT NULL,
    NumeroPCS               INT            NULL,       -- NULL en columnas por fila
    Campo                   NVARCHAR(40)   NOT NULL,   -- FAULT | STATUS | WARNING | MODULES | PCSnn | C | …
    ValorAnterior           NVARCHAR(255)  NULL,
    ValorNuevo              NVARCHAR(255)  NULL,
    ArchivoOrigen           NVARCHAR(260)  NOT NULL,
    Sha256Archivo           CHAR(64)       NOT NULL
);
GO

IF OBJECT_ID(N'dbo.excel_reference_run', N'U') IS NULL
CREATE TABLE dbo.excel_reference_run (      -- solo parámetros y KPI; el detalle queda en referencia_excel.json (D-24)
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

IF OBJECT_ID(N'dbo.reconciliation_result', N'U') IS NULL
CREATE TABLE dbo.reconciliation_result (    -- resumen por nivel + diferencias fuera de tolerancia
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

IF OBJECT_ID(N'dbo.detencion_revision', N'U') IS NULL
CREATE TABLE dbo.detencion_revision (      -- workflow humano; historial append-only (F-27)
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

-- ============================================================ ESTADO VIGENTE (se reemplaza por mes, ADR-12)
-- Anio/Mes = mes del período que publicó la fila (unidad de reemplazo, D-20). NumCorrida = última carga que la escribió.
IF OBJECT_ID(N'dbo.muestra_pcs', N'U') IS NULL
CREATE TABLE dbo.muestra_pcs (             -- 1 fila por bloque de 15 min y PCS: dato de SCADA + exclusión + aporte al KPI
    IdProyecto                       INT      NOT NULL CONSTRAINT FK_mpcs_proyecto REFERENCES dbo.proyecto(IdProyecto),
    SerialFecha                      FLOAT    NOT NULL,   -- RawData-PCS!A exacto (ADR-03)
    Ocurrencia                       TINYINT  NOT NULL,   -- 1; 2… si el timestamp se repite (cambio de hora, F-32)
    NumeroPCS                        INT      NOT NULL,
    Anio                             SMALLINT NOT NULL,
    Mes                              TINYINT  NOT NULL,
    MarcaTiempo                      DATETIME NULL,
    NumeroFilaOrigen                 INT      NOT NULL,
    FallaRaw                         NVARCHAR(255) NULL,
    FallaRawEsNumero                 BIT      NOT NULL,
    EstadoRaw                        NVARCHAR(255) NULL,
    AdvertenciaRaw                   NVARCHAR(255) NULL,
    ModulosRaw                       FLOAT    NULL,       -- NUMBER_OF_MODULES tal cual (NULL = vacío)
    ModulosDisponibles               FLOAT    NOT NULL,   -- = baterías por PCS si vacío
    ModulosDisponiblesNulo           BIT      NOT NULL,
    ValorExclusion                   TINYINT  NOT NULL,   -- Exclusion_Matrix 0/1/2 (F-37)
    BateriasPrevias                  FLOAT    NULL,       -- solo valor 2
    BateriasIndisponibles            FLOAT    NOT NULL,
    FactorOperacional                FLOAT    NOT NULL,   -- PlantActivity!C (solo si C21)
    BateriasIndisponiblesPonderadas  FLOAT    NOT NULL,   -- Calc!F:BN
    ImpactoRackPonderado             FLOAT    NOT NULL,   -- aporte a C14
    EsFilaNuevaDelExport             BIT      NOT NULL,
    NumCorrida                       INT      NOT NULL,
    CONSTRAINT PK_muestra_pcs PRIMARY KEY (IdProyecto, SerialFecha, Ocurrencia, NumeroPCS)
);
GO

IF OBJECT_ID(N'dbo.muestra_planta', N'U') IS NULL
CREATE TABLE dbo.muestra_planta (          -- 1 fila por bloque: PlantActivity + fila de Exclusion_Matrix
    IdProyecto                      INT      NOT NULL CONSTRAINT FK_mplanta_proyecto REFERENCES dbo.proyecto(IdProyecto),
    SerialFecha                     FLOAT    NOT NULL,
    Ocurrencia                      TINYINT  NOT NULL,
    Anio                            SMALLINT NOT NULL,
    Mes                             TINYINT  NOT NULL,
    MarcaTiempo                     DATETIME NULL,
    NumeroFilaOrigen                INT      NOT NULL,
    FactorOperacionalRaw            FLOAT    NULL,        -- PlantActivity!C
    EventoExcusadoRaw               FLOAT    NULL,        -- PlantActivity!D: espejo de Calc!BO, no pondera (F-37)
    SetpointPotenciaActivaKW        FLOAT    NULL,
    DroopSobrefrecuenciaHabilitado  FLOAT    NULL,
    DroopBajafrecuenciaHabilitado   FLOAT    NULL,
    PotenciaActivaPOIKW             FLOAT    NULL,
    PorcentajeSOC                   FLOAT    NULL,
    EventoExcusadoFila              FLOAT    NULL,        -- Exclusion_Matrix "Excused Event"
    ComentarioExclusion             NVARCHAR(500) NULL,   -- Exclusion_Matrix "Comments"
    NumCorrida                      INT      NOT NULL,
    CONSTRAINT PK_muestra_planta PRIMARY KEY (IdProyecto, SerialFecha, Ocurrencia)
);
GO

IF OBJECT_ID(N'dbo.detencion', N'U') IS NULL
CREATE TABLE dbo.detencion (               -- ListOfFaults + detención operacional; IdDetencion estable entre recargas (D-25)
    IdDetencion                   BIGINT IDENTITY CONSTRAINT PK_detencion PRIMARY KEY,
    IdProyecto                    INT            NOT NULL CONSTRAINT FK_detencion_proyecto REFERENCES dbo.proyecto(IdProyecto),
    Anio                          SMALLINT       NOT NULL,   -- período que la generó (FechaInicio puede ser del mes anterior: A − C23)
    Mes                           TINYINT        NOT NULL,
    NumeroPCS                     INT            NOT NULL,
    FechaInicio                   DATETIME       NOT NULL,
    Ocurrencia                    TINYINT        NOT NULL,
    FechaTermino                  DATETIME       NOT NULL,
    DuracionSegundos              INT            NOT NULL,   -- ROUND(DuracionHoras*3600)
    DuracionHoras                 FLOAT          NOT NULL,   -- ListOfFaults!E
    IdTipoDetencion               INT            NULL CONSTRAINT FK_detencion_tipo REFERENCES dbo.tipo_detencion(IdTipoDetencion),
    CodigoFalla                   NVARCHAR(300)  NOT NULL,
    DescripcionFalla              NVARCHAR(255)  NOT NULL,
    DescripcionFallaFallback      BIT            NOT NULL,
    SerialInicio                  FLOAT          NOT NULL,
    SerialFin                     FLOAT          NOT NULL,
    NumeroBloques                 INT            NOT NULL,
    SumaBloques                   FLOAT          NOT NULL,
    PromedioBateriasInvolucradas  FLOAT          NOT NULL,
    HorasRackIndisponibles        FLOAT          NOT NULL,
    EventoArrastradoExcel         BIT            NOT NULL,
    ExcelHabriaFallado            BIT            NOT NULL,
    CerradoPorModulosNulo         BIT            NOT NULL,   -- terminó porque la fila siguiente estaba vacía
    EsExcusable                   BIT            NOT NULL,   -- algún bloque con Exclusion_Matrix ≠ 0 (F-37)
    OrdenExcel                    INT            NOT NULL,   -- fila de ListOfFaults − 5 en la última carga
    NumCorrida                    INT            NOT NULL,
    CONSTRAINT UQ_detencion_negocio UNIQUE (IdProyecto, Anio, Mes, NumeroPCS, FechaInicio, Ocurrencia)
);
GO

IF OBJECT_ID(N'dbo.disponibilidad_diaria', N'U') IS NULL
CREATE TABLE dbo.disponibilidad_diaria (   -- hoja Daily
    IdProyecto                           INT      NOT NULL CONSTRAINT FK_ddiaria_proyecto REFERENCES dbo.proyecto(IdProyecto),
    Dia                                  DATETIME NOT NULL,
    Anio                                 SMALLINT NOT NULL,
    Mes                                  TINYINT  NOT NULL,
    DiaN                                 INT      NOT NULL,
    BloquesRacksIndisponiblesDiarios     FLOAT    NOT NULL,
    BloquesRacksIndisponiblesAcumulados  FLOAT    NOT NULL,
    Disponibilidad                       FLOAT    NULL,
    Variacion                            FLOAT    NULL,
    EstadoExclusiones                    NVARCHAR(20) NOT NULL,
    NumCorrida                           INT      NOT NULL,
    CONSTRAINT PK_disponibilidad_diaria PRIMARY KEY (IdProyecto, Dia)
);
GO

IF OBJECT_ID(N'dbo.disponibilidad_mensual', N'U') IS NULL
CREATE TABLE dbo.disponibilidad_mensual (  -- 1 fila por mes (C12, C14, C16, C19, L10) — lo que lee Power BI
    IdProyecto                    INT            NOT NULL CONSTRAINT FK_dmensual_proyecto REFERENCES dbo.proyecto(IdProyecto),
    Anio                          SMALLINT       NOT NULL,
    Mes                           TINYINT        NOT NULL CONSTRAINT CK_dmensual_mes CHECK (Mes BETWEEN 1 AND 12),
    InicioPeriodo                 DATETIME       NOT NULL,
    UltimoDato                    DATETIME       NULL,
    MesCompleto                   BIT            NOT NULL,
    DiasMes                       FLOAT          NOT NULL,
    BloquesMuestreo               FLOAT          NOT NULL,   -- C12 (intervalos existentes, D-07)
    BloquesMuestreoCalendario     FLOAT          NULL,
    TotalRacks                    INT            NOT NULL,
    BloquesRacksIndisponibles     FLOAT          NOT NULL,   -- C14
    DisponibilidadMensual         FLOAT          NULL,       -- C16
    DisponibilidadAnualAcumulada  FLOAT          NULL,       -- C19 (no es el acumulado del año: ver v_disponibilidad_anual)
    HorasRackEventos              FLOAT          NULL,       -- L10
    MinutosMuestreoDerivado       FLOAT          NULL,       -- C23
    DisponibilidadContractual     FLOAT          NOT NULL CONSTRAINT DF_dmensual_contractual DEFAULT 0.98,
    AcumulaEnAnual                BIT            NOT NULL,   -- D-06: desde jul-2026; los años siguientes, desde enero
    EstadoExclusiones             NVARCHAR(20)   NOT NULL CONSTRAINT CK_dmensual_exclusiones CHECK (EstadoExclusiones IN (N'sin_exclusiones', N'con_exclusiones')),
    TipoCorrida                   NVARCHAR(20)   NOT NULL,
    Origen                        NVARCHAR(20)   NOT NULL CONSTRAINT CK_dmensual_origen CHECK (Origen IN (N'corrida', N'excel_manual')),
    VersionAlgoritmo              NVARCHAR(100)  NULL,
    NumCorrida                    INT            NULL,       -- NULL si excel_manual
    ActualizadoEn                 DATETIME       NOT NULL CONSTRAINT DF_dmensual_actualizado DEFAULT GETDATE(),
    CONSTRAINT PK_disponibilidad_mensual PRIMARY KEY (IdProyecto, Anio, Mes)
);
GO

IF OBJECT_ID(N'dbo.calidad_dato', N'U') IS NULL
CREATE TABLE dbo.calidad_dato (            -- anomalías vigentes del mes
    IdCalidad        BIGINT IDENTITY CONSTRAINT PK_calidad_dato PRIMARY KEY,
    IdProyecto       INT            NOT NULL CONSTRAINT FK_calidad_proyecto REFERENCES dbo.proyecto(IdProyecto),
    Anio             SMALLINT       NOT NULL,
    Mes              TINYINT        NOT NULL,
    Tipo             NVARCHAR(50)   NOT NULL,
    Severidad        NVARCHAR(12)   NOT NULL CONSTRAINT CK_calidad_severidad CHECK (Severidad IN (N'info', N'advertencia', N'error')),
    NumeroFilaOrigen INT            NULL,
    NumeroPCS        INT            NULL,
    SerialFecha      FLOAT          NULL,
    Detalle          NVARCHAR(1000) NULL,
    NumCorrida       INT            NOT NULL
);
GO

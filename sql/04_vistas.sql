-- 04_vistas.sql — vistas para Power BI y auditoría (R12.6, R15.4, R18.4, R19.7). Idempotente.
-- bi_reader solo lee estas vistas (06_roles.sql); el encadenamiento de propiedad (dbo) evita dar
-- permisos sobre las tablas base.
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

-- Corrida vigente por (proyecto, año, mes de FinPeriodo): oficial y exitosa; gana "Con Exclusiones"
-- (cierre mensual con la Exclusion_Matrix de Alex) sobre "Sin Exclusiones" (semanales); a igualdad,
-- el período más largo y luego la más reciente. Las corridas golden nunca son vigentes.
CREATE OR ALTER VIEW dbo.v_corrida_oficial_vigente AS
SELECT IdCorrida, NumCorrida, IdProyecto, Anio, Mes, TipoCorrida, EstadoExclusiones, EtiquetaExclusiones,
       InicioPeriodo, FinPeriodo, IniciadoEn, FinalizadoEn, VersionAlgoritmo, ArchivoOrigen, HashArchivoOrigen
FROM (
    SELECT r.IdCorrida, r.NumCorrida, r.IdProyecto,
           YEAR(r.FinPeriodo) AS Anio, MONTH(r.FinPeriodo) AS Mes,
           r.TipoCorrida, r.EstadoExclusiones,
           CASE r.EstadoExclusiones WHEN N'con_exclusiones' THEN N'Con Exclusiones' ELSE N'Sin Exclusiones' END AS EtiquetaExclusiones,
           r.InicioPeriodo, r.FinPeriodo, r.IniciadoEn, r.FinalizadoEn, r.VersionAlgoritmo, r.ArchivoOrigen, r.HashArchivoOrigen,
           ROW_NUMBER() OVER (
               PARTITION BY r.IdProyecto, YEAR(r.FinPeriodo), MONTH(r.FinPeriodo)
               ORDER BY CASE r.EstadoExclusiones WHEN N'con_exclusiones' THEN 0 ELSE 1 END,
                        r.FinPeriodo DESC, r.IniciadoEn DESC) AS rn
    FROM dbo.etl_run AS r
    WHERE r.EsOficial = 1 AND r.Estado = N'success' AND r.TipoCorrida <> N'golden'
) AS x
WHERE rn = 1;
GO

CREATE OR ALTER VIEW dbo.v_kpi_vigente AS
SELECT v.IdProyecto, v.Anio, v.Mes, v.IdCorrida, v.NumCorrida, v.TipoCorrida, v.EtiquetaExclusiones,
       v.InicioPeriodo, v.FinPeriodo,
       a.BloquesMuestreo, a.TotalRacks, a.BloquesRacksIndisponibles,
       a.DisponibilidadPeriodo, a.DisponibilidadAnualAcumulada, a.MinutosMuestreoDerivado, a.HorasRackEventos
FROM dbo.v_corrida_oficial_vigente AS v
JOIN dbo.availability_run_result AS a ON a.IdCorrida = v.IdCorrida;
GO

CREATE OR ALTER VIEW dbo.v_daily_vigente AS
SELECT v.IdProyecto, v.Anio, v.Mes, v.IdCorrida, v.NumCorrida, v.EtiquetaExclusiones,
       d.DiaN, d.Dia, d.BloquesRacksIndisponiblesDiarios, d.BloquesRacksIndisponiblesAcumulados,
       d.Disponibilidad, d.Variacion
FROM dbo.v_corrida_oficial_vigente AS v
JOIN dbo.daily_availability AS d ON d.IdCorrida = v.IdCorrida;
GO

CREATE OR ALTER VIEW dbo.v_fault_event_vigente AS
SELECT v.IdProyecto, v.Anio, v.Mes, v.IdCorrida, v.NumCorrida, v.EtiquetaExclusiones, e.*
FROM dbo.v_corrida_oficial_vigente AS v
CROSS APPLY (SELECT f.OrdenExcel, f.NumeroPCS, f.MarcaTiempoInicio, f.MarcaTiempoFin, f.DuracionHoras,
                    f.CodigoFalla, f.DescripcionFalla, f.DescripcionFallaFallback, f.PromedioBateriasInvolucradas,
                    f.HorasRackIndisponibles, f.EventoArrastradoExcel, f.ExcelHabriaFallado, f.TieneExclusion
             FROM dbo.fault_event AS f WHERE f.IdCorrida = v.IdCorrida) AS e;
GO

CREATE OR ALTER VIEW dbo.v_fault_code_vigente AS
SELECT v.IdProyecto, v.Anio, v.Mes, v.IdCorrida, v.NumCorrida, v.EtiquetaExclusiones,
       s.Ranking, s.CodigoFalla, s.DescripcionFallaPE, s.HorasRackIndisponibles, s.Porcentaje
FROM dbo.v_corrida_oficial_vigente AS v
JOIN dbo.fault_code_summary AS s ON s.IdCorrida = v.IdCorrida;
GO

-- KPI mensual vigente por mes con la disponibilidad calculada (R10). Igual que las corridas (D-17,
-- Checkpoint C-2): gana el de una corrida "Con Exclusiones"; a igualdad, el más reciente.
CREATE OR ALTER VIEW dbo.v_monthly_kpi_vigente AS
SELECT k.IdProyecto, k.Anio, k.Mes, k.DiasMes, k.BloquesMuestreo, k.BloquesRacksIndisponibles,
       CASE WHEN p.TotalRacks > 0 AND k.BloquesMuestreo > 0
            THEN 1 - k.BloquesRacksIndisponibles / (p.TotalRacks * k.BloquesMuestreo) END AS DisponibilidadMensual,
       k.DisponibilidadContractual, k.Origen, k.IdCorrida, k.NumCorrida, k.RegistradoEn
FROM (
    SELECT m.*, r.NumCorrida, ROW_NUMBER() OVER (
               PARTITION BY m.IdProyecto, m.Anio, m.Mes
               ORDER BY CASE r.EstadoExclusiones WHEN N'con_exclusiones' THEN 0 ELSE 1 END,
                        m.RegistradoEn DESC, m.IdKpiMensual DESC) AS rn
    FROM dbo.monthly_official_kpi AS m
    LEFT JOIN dbo.etl_run AS r ON r.IdCorrida = m.IdCorrida
) AS k
JOIN dbo.proyecto AS p ON p.IdProyecto = k.IdProyecto
WHERE k.rn = 1;
GO

-- Annual_AVA del último mes vigente de cada año.
CREATE OR ALTER VIEW dbo.v_annual_vigente AS
SELECT x.IdProyecto, a.Anio, a.Mes, a.DiasMes, a.BloquesMuestreo, a.BloquesRacksIndisponibles, a.DisponibilidadMensual,
       a.BloquesMuestreoAcumulados, a.BloquesIndisponiblesAcumulados, a.DisponibilidadAcumulada,
       a.DisponibilidadContractual, x.IdCorrida, x.NumCorrida
FROM (
    SELECT v.IdProyecto, v.Anio, v.IdCorrida, v.NumCorrida,
           ROW_NUMBER() OVER (PARTITION BY v.IdProyecto, v.Anio ORDER BY v.Mes DESC) AS rn
    FROM dbo.v_corrida_oficial_vigente AS v
    WHERE EXISTS (SELECT 1 FROM dbo.annual_availability AS aa WHERE aa.IdCorrida = v.IdCorrida)
) AS x
JOIN dbo.annual_availability AS a ON a.IdCorrida = x.IdCorrida AND a.Anio = x.Anio
WHERE x.rn = 1;
GO

-- Detenciones de las corridas vigentes con su última revisión (la revisión sobrevive a los reprocesos, F-27).
CREATE OR ALTER VIEW dbo.v_detencion_vigente AS
SELECT d.IdDetencion, d.IdProyecto, v.Anio, v.Mes, d.IdCorrida, v.NumCorrida, v.EtiquetaExclusiones, d.OrdenExcel, d.NumeroPCS,
       d.FechaInicio, d.FechaTermino, d.DuracionSegundos, d.DuracionHoras, d.IdTipoDetencion, t.Significado, t.Operativo,
       d.CodigoFalla, d.DescripcionFalla, d.DescripcionFallaFallback, d.PromedioBateriasInvolucradas,
       d.HorasRackIndisponibles, d.CerradoPorModulosNulo, d.EsExcusable, d.EventoArrastradoExcel,
       COALESCE(rv.EstadoRevision, N'pendiente') AS EstadoRevision, rv.Observacion, rv.RevisadoPor, rv.RevisadoEn
FROM dbo.v_corrida_oficial_vigente AS v
JOIN dbo.detencion AS d ON d.IdCorrida = v.IdCorrida
LEFT JOIN dbo.tipo_detencion AS t ON t.IdTipoDetencion = d.IdTipoDetencion
OUTER APPLY (
    SELECT TOP (1) r.EstadoRevision, r.Observacion, r.RevisadoPor, r.RevisadoEn
    FROM dbo.detencion_revision AS r
    WHERE r.IdProyecto = d.IdProyecto AND r.NumeroPCS = d.NumeroPCS AND r.FechaInicio = d.FechaInicio
    ORDER BY r.RevisadoEn DESC, r.IdRevision DESC
) AS rv;
GO

CREATE OR ALTER VIEW dbo.v_calidad_corrida AS
SELECT r.IdCorrida, r.NumCorrida, r.IdProyecto, r.TipoCorrida, r.EsOficial, r.EstadoExclusiones, r.Estado,
       r.InicioPeriodo, r.FinPeriodo, r.IniciadoEn, q.Tipo, q.Severidad, q.Cantidad
FROM dbo.etl_run AS r
LEFT JOIN (
    SELECT IdCorrida, Tipo, Severidad, COUNT(*) AS Cantidad
    FROM dbo.data_quality_issue
    GROUP BY IdCorrida, Tipo, Severidad
) AS q ON q.IdCorrida = r.IdCorrida;
GO

-- Intervalos con NUMBER_OF_MODULES vacío en toda la historia vigente (R15.4): no se suprimen.
CREATE OR ALTER VIEW dbo.v_modulos_nulos_historico AS
SELECT v.IdProyecto, v.Anio, v.Mes, s.IdCorrida, v.NumCorrida, s.NumeroFilaOrigen, s.NumeroPCS,
       s.SerialFechaExcelOrigen, s.MarcaTiempoLocalOrigen, s.FallaRaw, s.EstadoRaw
FROM dbo.v_corrida_oficial_vigente AS v
JOIN dbo.raw_pcs_sample AS s ON s.IdCorrida = v.IdCorrida
WHERE s.ModulosDisponiblesNulo = 1;
GO

-- Cambios de datos (reproceso o cargas mensuales, D-13/D-17) con el KPI de la corrida que los aplicó.
CREATE OR ALTER VIEW dbo.v_correccion_dato AS
SELECT c.IdCorreccion, c.IdCorrida, r.NumCorrida, r.TipoCorrida, r.IniciadoEn AS FechaEjecucion, c.TipoCorreccion, c.Hoja,
       c.NumeroFilaOrigen, c.MarcaTiempoLocalOrigen, c.NumeroPCS, c.Campo, c.ValorAnterior, c.ValorNuevo,
       c.ArchivoOrigen, c.Sha256Archivo, a.BloquesRacksIndisponibles AS C14Corrida, a.DisponibilidadPeriodo AS C16Corrida
FROM dbo.correccion_dato AS c
JOIN dbo.etl_run AS r ON r.IdCorrida = c.IdCorrida
LEFT JOIN dbo.availability_run_result AS a ON a.IdCorrida = c.IdCorrida;
GO

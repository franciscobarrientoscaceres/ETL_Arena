-- 04_vistas.sql — vistas del esquema v2 (ADR-12). Power BI lee las tablas de estado y estas vistas;
-- ninguna expone corridas (R18.4 rev. 3). Idempotente.
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

-- Acumulado del año (hoja Annual_AVA, R10.3) desde disponibilidad_mensual. Solo acumulan los meses con
-- AcumulaEnAnual = 1 (D-06: 2026 desde julio; los años siguientes desde enero).
CREATE OR ALTER VIEW dbo.v_disponibilidad_anual AS
SELECT m.IdProyecto, m.Anio, m.Mes, m.DiasMes, m.BloquesMuestreo, m.TotalRacks, m.BloquesRacksIndisponibles,
       CASE WHEN m.TotalRacks > 0 AND m.BloquesMuestreo > 0
            THEN 1 - m.BloquesRacksIndisponibles / (m.TotalRacks * m.BloquesMuestreo) END AS DisponibilidadMensual,
       a.BloquesMuestreoAcumulados, a.BloquesIndisponiblesAcumulados,
       CASE WHEN m.TotalRacks > 0 AND a.BloquesMuestreoAcumulados > 0
            THEN 1 - a.BloquesIndisponiblesAcumulados / (m.TotalRacks * a.BloquesMuestreoAcumulados) END AS DisponibilidadAcumulada,
       m.DisponibilidadContractual, m.MesCompleto, m.EstadoExclusiones, m.Origen, m.NumCorrida
FROM dbo.disponibilidad_mensual AS m
CROSS APPLY (
    SELECT SUM(x.BloquesMuestreo) AS BloquesMuestreoAcumulados,
           SUM(x.BloquesRacksIndisponibles) AS BloquesIndisponiblesAcumulados
    FROM dbo.disponibilidad_mensual AS x
    WHERE x.IdProyecto = m.IdProyecto AND x.Anio = m.Anio AND x.Mes <= m.Mes AND x.AcumulaEnAnual = 1
) AS a
WHERE m.AcumulaEnAnual = 1;
GO

-- Horas-rack por código de falla y mes (antes ListOfFaults!N:Q / fault_code_summary).
CREATE OR ALTER VIEW dbo.v_resumen_codigo_mensual AS
SELECT d.IdProyecto, d.Anio, d.Mes, d.CodigoFalla, t.DescripcionFallaPE, t.Significado,
       COUNT(*) AS Detenciones,
       SUM(d.DuracionHoras) AS DuracionHoras,
       SUM(d.HorasRackIndisponibles) AS HorasRackIndisponibles,
       SUM(d.HorasRackIndisponibles)
         / NULLIF(SUM(SUM(d.HorasRackIndisponibles)) OVER (PARTITION BY d.IdProyecto, d.Anio, d.Mes), 0) AS Porcentaje
FROM dbo.detencion AS d
LEFT JOIN dbo.tipo_detencion AS t ON t.IdTipoDetencion = d.IdTipoDetencion
GROUP BY d.IdProyecto, d.Anio, d.Mes, d.CodigoFalla, t.DescripcionFallaPE, t.Significado;
GO

-- Detenciones vigentes con su última revisión (la revisión sobrevive a las recargas, F-27, D-25).
CREATE OR ALTER VIEW dbo.v_detencion AS
SELECT d.IdDetencion, d.IdProyecto, d.Anio, d.Mes, d.NumeroPCS, d.FechaInicio, d.FechaTermino,
       d.DuracionSegundos, d.DuracionHoras, d.IdTipoDetencion, t.Significado, t.Operativo,
       d.CodigoFalla, d.DescripcionFalla, d.DescripcionFallaFallback, d.PromedioBateriasInvolucradas,
       d.HorasRackIndisponibles, d.CerradoPorModulosNulo, d.EsExcusable, d.EventoArrastradoExcel, d.NumCorrida,
       COALESCE(rv.EstadoRevision, N'pendiente') AS EstadoRevision, rv.Observacion, rv.RevisadoPor, rv.RevisadoEn
FROM dbo.detencion AS d
LEFT JOIN dbo.tipo_detencion AS t ON t.IdTipoDetencion = d.IdTipoDetencion
OUTER APPLY (
    SELECT TOP (1) r.EstadoRevision, r.Observacion, r.RevisadoPor, r.RevisadoEn
    FROM dbo.detencion_revision AS r
    WHERE r.IdProyecto = d.IdProyecto AND r.NumeroPCS = d.NumeroPCS AND r.FechaInicio = d.FechaInicio
    ORDER BY r.RevisadoEn DESC, r.IdRevision DESC
) AS rv;
GO

-- Bloques con NUMBER_OF_MODULES vacío (R15.4): se cuentan como disponibles pero no se suprimen.
CREATE OR ALTER VIEW dbo.v_modulos_nulos AS
SELECT IdProyecto, Anio, Mes, SerialFecha, Ocurrencia, MarcaTiempo, NumeroFilaOrigen, NumeroPCS, FallaRaw, EstadoRaw, NumCorrida
FROM dbo.muestra_pcs
WHERE ModulosDisponiblesNulo = 1;
GO

-- Datos que cambiaron al recargar o al cargar la matriz (D-13, D-24), con la carga que los aplicó.
CREATE OR ALTER VIEW dbo.v_correccion_dato AS
SELECT c.IdCorreccion, r.NumCorrida, c.IdCorrida, r.TipoCorrida, r.IniciadoEn AS FechaEjecucion, c.TipoCorreccion,
       c.Hoja, c.NumeroFilaOrigen, c.MarcaTiempoLocalOrigen, c.NumeroPCS, c.Campo, c.ValorAnterior, c.ValorNuevo,
       c.ArchivoOrigen, c.Sha256Archivo
FROM dbo.correccion_dato AS c
JOIN dbo.etl_run AS r ON r.IdCorrida = c.IdCorrida;
GO

-- Registro de ejecuciones (una fila por carga), para auditoría.
CREATE OR ALTER VIEW dbo.v_ejecuciones AS
SELECT NumCorrida, IdCorrida, IdProyecto, TipoCorrida, EstadoExclusiones, InicioPeriodo, FinPeriodo, Estado,
       Publicada, MesesPublicados, IniciadoEn, FinalizadoEn, ArchivoOrigen, HashArchivoOrigen, VersionAlgoritmo,
       MensajeError
FROM dbo.etl_run;
GO

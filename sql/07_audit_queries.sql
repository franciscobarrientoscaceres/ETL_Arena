-- 07_audit_queries.sql — consultas de auditoría del esquema v2 (R15, ADR-12). No crea objetos: fijar el mes y ejecutar.
-- Trazabilidad: disponibilidad_mensual → muestra_pcs (aporte y dato crudo, fila de origen) → etl_run (archivo, sha256).
SET NOCOUNT ON;
DECLARE @IdProyecto INT = 1, @Anio SMALLINT = 2026, @Mes TINYINT = 9;
DECLARE @NumeroPCS INT = NULL;                -- opcional: filtrar un PCS

-- 1) KPI del mes y su recomposición desde las muestras: C14 = Σ ImpactoRackPonderado, C12 = bloques del mes.
SELECT m.Anio, m.Mes, m.NumCorrida, m.EstadoExclusiones, m.BloquesMuestreo AS C12, m.BloquesRacksIndisponibles AS C14,
       s.Bloques, s.SumaImpacto, m.BloquesRacksIndisponibles - s.SumaImpacto AS DeltaC14
FROM dbo.disponibilidad_mensual AS m
CROSS APPLY (
    SELECT COUNT(DISTINCT CONCAT(x.SerialFecha, '|', x.Ocurrencia)) AS Bloques, SUM(x.ImpactoRackPonderado) AS SumaImpacto
    FROM dbo.muestra_pcs AS x WHERE x.IdProyecto = m.IdProyecto AND x.Anio = m.Anio AND x.Mes = m.Mes
) AS s
WHERE m.IdProyecto = @IdProyecto AND m.Anio = @Anio AND m.Mes = @Mes;

-- 2) Bloques que aportan a C14 → dato crudo de SCADA, fila de origen y archivo de la carga vigente.
SELECT TOP (500)
       x.MarcaTiempo, x.NumeroFilaOrigen, x.NumeroPCS, x.ModulosRaw, x.ModulosDisponibles, x.ValorExclusion,
       x.BateriasIndisponiblesPonderadas, x.ImpactoRackPonderado, x.FallaRaw, r.NumCorrida, r.ArchivoOrigen, r.HashArchivoOrigen
FROM dbo.muestra_pcs AS x
JOIN dbo.etl_run AS r ON r.NumCorrida = x.NumCorrida
WHERE x.IdProyecto = @IdProyecto AND x.Anio = @Anio AND x.Mes = @Mes AND x.ImpactoRackPonderado <> 0
  AND (@NumeroPCS IS NULL OR x.NumeroPCS = @NumeroPCS)
ORDER BY x.SerialFecha, x.Ocurrencia, x.NumeroPCS;

-- 3) Detenciones con descripción tomada de la fila anterior (R8.8, F-04).
SELECT IdDetencion, NumeroPCS, FechaInicio, CodigoFalla, DescripcionFalla
FROM dbo.detencion
WHERE IdProyecto = @IdProyecto AND Anio = @Anio AND Mes = @Mes AND DescripcionFallaFallback = 1;

-- 4) Detenciones que reproducen defectos del VBA: arrastre entre PCS (F-06) y Type mismatch en la fila 2 (F-11).
SELECT IdDetencion, NumeroPCS, FechaInicio, EventoArrastradoExcel, ExcelHabriaFallado
FROM dbo.detencion
WHERE IdProyecto = @IdProyecto AND Anio = @Anio AND Mes = @Mes AND (EventoArrastradoExcel = 1 OR ExcelHabriaFallado = 1);

-- 5) Invariante C14 = 4·L10 (R13.8) y detenciones con exclusión (F-37).
SELECT m.BloquesRacksIndisponibles AS C14, m.HorasRackEventos AS L10,
       m.BloquesRacksIndisponibles - 4 * m.HorasRackEventos AS Delta,
       (SELECT COUNT(*) FROM dbo.detencion d WHERE d.IdProyecto = m.IdProyecto AND d.Anio = m.Anio AND d.Mes = m.Mes AND d.EsExcusable = 1) AS DetencionesConExclusion
FROM dbo.disponibilidad_mensual AS m
WHERE m.IdProyecto = @IdProyecto AND m.Anio = @Anio AND m.Mes = @Mes;

-- 6) Anomalías de calidad vigentes del mes por tipo (R15.2).
SELECT Tipo, Severidad, COUNT(*) AS Cantidad
FROM dbo.calidad_dato
WHERE IdProyecto = @IdProyecto AND Anio = @Anio AND Mes = @Mes
GROUP BY Tipo, Severidad
ORDER BY Cantidad DESC;

-- 7) Historia del mes: cada carga que lo tocó y los datos que cambiaron (D-24).
SELECT r.NumCorrida, r.TipoCorrida, r.Estado, r.Publicada, r.IniciadoEn, r.ArchivoOrigen,
       (SELECT COUNT(*) FROM dbo.correccion_dato c WHERE c.IdCorrida = r.IdCorrida) AS DatosCorregidos
FROM dbo.etl_run AS r
WHERE r.IdProyecto = @IdProyecto AND YEAR(r.InicioPeriodo) = @Anio AND MONTH(r.InicioPeriodo) = @Mes
ORDER BY r.NumCorrida;

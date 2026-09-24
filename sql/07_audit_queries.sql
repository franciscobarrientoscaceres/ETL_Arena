-- 07_audit_queries.sql — consultas de auditoría (R15.1). No crea objetos: fijar @IdCorrida y ejecutar.
-- Trazabilidad: availability_run_result → availability_sample_result → raw_pcs_sample → (archivo, hash, fila, columna).
SET NOCOUNT ON;
DECLARE @IdCorrida UNIQUEIDENTIFIER = NULL;   -- p. ej. (SELECT TOP 1 IdCorrida FROM dbo.v_corrida_oficial_vigente ORDER BY Anio DESC, Mes DESC)
DECLARE @NumeroPCS INT = NULL;                -- opcional: filtrar un PCS

-- 1) KPI de la corrida y su recomposición desde las muestras: C14 = Σ ImpactoRackPonderado, C12 = filas.
SELECT a.IdCorrida, a.BloquesMuestreo AS C12, a.BloquesRacksIndisponibles AS C14,
       m.FilasMuestra, m.SumaImpacto, a.BloquesRacksIndisponibles - m.SumaImpacto AS DeltaC14
FROM dbo.availability_run_result AS a
CROSS APPLY (
    SELECT COUNT(DISTINCT s.NumeroFilaOrigen) AS FilasMuestra, SUM(s.ImpactoRackPonderado) AS SumaImpacto
    FROM dbo.availability_sample_result AS s WHERE s.IdCorrida = a.IdCorrida
) AS m
WHERE a.IdCorrida = @IdCorrida;

-- 2) Muestra → dato crudo → columna y archivo de origen, para los bloques que aportan a C14.
SELECT TOP (500)
       s.NumeroFilaOrigen, s.NumeroPCS, s.MarcaTiempoMuestra, s.ModulosDisponibles, s.ValorExclusion,
       s.BateriasIndisponiblesPonderadas, s.ImpactoRackPonderado,
       r.ModulosRaw, r.FallaRaw, r.EstadoRaw,
       cm.NumeroColumna, cm.NombreColumna, e.ArchivoOrigen, e.HashArchivoOrigen
FROM dbo.availability_sample_result AS s
JOIN dbo.raw_pcs_sample AS r
  ON r.IdCorrida = s.IdCorrida AND r.NumeroFilaOrigen = s.NumeroFilaOrigen AND r.NumeroPCS = s.NumeroPCS
JOIN dbo.raw_pcs_column_map AS cm
  ON cm.IdCorrida = s.IdCorrida AND cm.NumeroPCS = s.NumeroPCS AND cm.Campo = N'modulos'
JOIN dbo.etl_run AS e ON e.IdCorrida = s.IdCorrida
WHERE s.IdCorrida = @IdCorrida AND s.ImpactoRackPonderado <> 0 AND (@NumeroPCS IS NULL OR s.NumeroPCS = @NumeroPCS)
ORDER BY s.NumeroFilaOrigen, s.NumeroPCS;

-- 3) Eventos con descripción tomada de la fila anterior (R8.8, F-04).
SELECT OrdenExcel, NumeroPCS, MarcaTiempoInicio, MarcaTiempoFin, CodigoFalla, DescripcionFalla
FROM dbo.fault_event
WHERE IdCorrida = @IdCorrida AND DescripcionFallaFallback = 1
ORDER BY OrdenExcel;

-- 4) Eventos que reproducen defectos del VBA: arrastre entre PCS (F-06) y Type mismatch en la fila 2 (F-11).
SELECT OrdenExcel, NumeroPCS, MarcaTiempoInicio, MarcaTiempoFin, NumeroBloques, SumaBloques,
       HorasRackIndisponibles, EventoArrastradoExcel, ExcelHabriaFallado
FROM dbo.fault_event
WHERE IdCorrida = @IdCorrida AND (EventoArrastradoExcel = 1 OR ExcelHabriaFallado = 1)
ORDER BY OrdenExcel;

-- 5) Invariante C14 = 4·L10 (R13.8) y eventos con exclusión (F-37).
SELECT a.BloquesRacksIndisponibles AS C14, a.HorasRackEventos AS L10,
       a.BloquesRacksIndisponibles - 4 * a.HorasRackEventos AS Delta,
       (SELECT COUNT(*) FROM dbo.fault_event f WHERE f.IdCorrida = a.IdCorrida AND f.TieneExclusion = 1) AS EventosConExclusion
FROM dbo.availability_run_result AS a
WHERE a.IdCorrida = @IdCorrida;

-- 6) Anomalías de calidad de la corrida por tipo (R15.2).
SELECT Tipo, Severidad, COUNT(*) AS Cantidad, MIN(NumeroFilaOrigen) AS PrimeraFila
FROM dbo.data_quality_issue
WHERE IdCorrida = @IdCorrida
GROUP BY Tipo, Severidad
ORDER BY CASE Severidad WHEN N'error' THEN 0 WHEN N'advertencia' THEN 1 ELSE 2 END, Cantidad DESC;

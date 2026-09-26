-- 06_seed_annual_manual.sql — KPI mensual histórico de Annual_AVA que no se puede reproducir
-- (jul/ago 2026, R10.4, F-20, D-06) en disponibilidad_mensual con Origen = excel_manual (esquema v2).
-- Idempotente: solo inserta si el mes no existe. Valores exactos de Annual_AVA!D/E/F del libro de septiembre 2026.
-- Agosto: D-11 abierta. Una carga de estos meses solo los reemplaza con --reemplazar-manual (D-23).
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

INSERT INTO dbo.disponibilidad_mensual (
    IdProyecto, Anio, Mes, InicioPeriodo, UltimoDato, MesCompleto, DiasMes, BloquesMuestreo, BloquesMuestreoCalendario,
    TotalRacks, BloquesRacksIndisponibles, DisponibilidadMensual, DisponibilidadAnualAcumulada, HorasRackEventos,
    MinutosMuestreoDerivado, DisponibilidadContractual, AcumulaEnAnual, EstadoExclusiones, TipoCorrida, Origen,
    VersionAlgoritmo, NumCorrida)
SELECT s.IdProyecto, s.Anio, s.Mes, DATEFROMPARTS(s.Anio, s.Mes, 1), NULL, 1, s.DiasMes, s.BloquesMuestreo, s.BloquesMuestreo,
       p.TotalRacks, s.BloquesRacksIndisponibles,
       1 - s.BloquesRacksIndisponibles / (p.TotalRacks * s.BloquesMuestreo), NULL, NULL,
       p.MinutosMuestreo, 0.98, 1, N'sin_exclusiones', N'cierre_mensual', N'excel_manual', NULL, NULL
FROM (VALUES
    (1, 2026, 7, 31.0, 2976.0, 350972.29999999946),
    (1, 2026, 8, 31.0, 2976.0, 305182.96799999941)
) AS s (IdProyecto, Anio, Mes, DiasMes, BloquesMuestreo, BloquesRacksIndisponibles)
JOIN dbo.proyecto AS p ON p.IdProyecto = s.IdProyecto
WHERE NOT EXISTS (
    SELECT 1 FROM dbo.disponibilidad_mensual AS k WHERE k.IdProyecto = s.IdProyecto AND k.Anio = s.Anio AND k.Mes = s.Mes
);
GO

-- 06_seed_annual_manual.sql — KPI mensual histórico de Annual_AVA que no se puede reproducir
-- (jul/ago 2026, R10.4, F-20, D-06). Origen = excel_manual. Idempotente: solo inserta si no existe.
-- Valores exactos de Annual_AVA!D/E/F del libro de septiembre 2026. Agosto: D-11 abierta (no se
-- reproduce con ninguna regla conocida; consultado a Alex).
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

INSERT INTO dbo.monthly_official_kpi (IdProyecto, Anio, Mes, DiasMes, BloquesMuestreo, BloquesRacksIndisponibles, DisponibilidadContractual, Origen, IdCorrida)
SELECT s.IdProyecto, s.Anio, s.Mes, s.DiasMes, s.BloquesMuestreo, s.BloquesRacksIndisponibles, 0.98, N'excel_manual', NULL
FROM (VALUES
    (1, 2026, 7, 31.0, 2976.0, 350972.29999999946),
    (1, 2026, 8, 31.0, 2976.0, 305182.96799999941)
) AS s (IdProyecto, Anio, Mes, DiasMes, BloquesMuestreo, BloquesRacksIndisponibles)
WHERE NOT EXISTS (
    SELECT 1 FROM dbo.monthly_official_kpi AS k
    WHERE k.IdProyecto = s.IdProyecto AND k.Anio = s.Anio AND k.Mes = s.Mes AND k.Origen = N'excel_manual'
);
GO

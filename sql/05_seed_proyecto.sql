-- 05_seed_proyecto.sql — proyectos (R12.1, R12.2). Idempotente (MERGE).
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

MERGE dbo.proyecto AS t
USING (VALUES
    (1, N'Arena BESS',    N'en_ejecucion',    CAST('2026-04-08T00:00:00' AS DATETIME), 61,   4,    12,   15,   N'America/Santiago', N'61 PCS × 4 BAC × 12 racks = 2928 racks'),
    (2, N'Copiapó A',     N'por_implementar', NULL,                                    NULL, NULL, NULL, NULL, N'America/Santiago', NULL),
    (3, N'Luz del Norte', N'por_implementar', NULL,                                    NULL, NULL, NULL, NULL, N'America/Santiago', NULL),
    (4, N'María Elena',   N'por_implementar', NULL,                                    NULL, NULL, NULL, NULL, N'America/Santiago', NULL)
) AS s (IdProyecto, Nombre, Estado, FechaInicio, NumPCS, NumBateriasPorPCS, NumRacksPorBAC, MinutosMuestreo, ZonaHoraria, Descripcion)
ON t.IdProyecto = s.IdProyecto
WHEN MATCHED THEN UPDATE SET
    Nombre = s.Nombre, Estado = s.Estado, FechaInicio = s.FechaInicio, NumPCS = s.NumPCS,
    NumBateriasPorPCS = s.NumBateriasPorPCS, NumRacksPorBAC = s.NumRacksPorBAC, MinutosMuestreo = s.MinutosMuestreo,
    ZonaHoraria = s.ZonaHoraria, Descripcion = s.Descripcion
WHEN NOT MATCHED THEN INSERT (IdProyecto, Nombre, Estado, FechaInicio, NumPCS, NumBateriasPorPCS, NumRacksPorBAC, MinutosMuestreo, ZonaHoraria, Descripcion)
    VALUES (s.IdProyecto, s.Nombre, s.Estado, s.FechaInicio, s.NumPCS, s.NumBateriasPorPCS, s.NumRacksPorBAC, s.MinutosMuestreo, s.ZonaHoraria, s.Descripcion);
GO

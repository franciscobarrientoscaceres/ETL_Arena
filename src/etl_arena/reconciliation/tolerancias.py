"""Tabla única de tolerancias de paridad (design.md §Tolerancias, R13.7).

``golden_index.json → _meta`` se mantiene sincronizado con estos valores (test dedicado).
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Tolerancias:
    ponderada_celda: float = 1e-12  # nivel 2: Calc!F:BN por celda
    c14: float = 1e-6
    kpi_disponibilidad: float = 1e-9  # C16, C19, Daily F/G
    daily_bloques: float = 1e-6  # Daily D/E
    evento_serial_dias: float = 1e-9  # ListOfFaults C/D
    evento_duracion_promedio: float = 1e-9  # ListOfFaults E/H
    evento_horas_rack: float = 1e-6  # ListOfFaults I
    l10: float = 1e-6


TOLERANCIAS = Tolerancias()

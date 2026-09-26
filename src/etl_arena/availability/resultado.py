"""Resultado de ``MotorDisponibilidad`` (espejo de ``Calculation-Availability``)."""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass

import numpy as np

from etl_arena.model import MuestraDisponibilidad


@dataclass
class ResultadoDisponibilidad:
    bloques_muestreo: int  # C12
    bloques_racks_indisponibles: float  # C14
    disponibilidad_periodo: float | None  # C16 (None = "N/A")
    disponibilidad_anual_acumulada: float | None  # C19
    minutos_muestreo_derivado: float | None  # C23
    filas_procesadas: np.ndarray  # (k,) índices i de la matriz, orden origen
    serial: np.ndarray  # (k,) Calc!E
    numero_fila: np.ndarray  # (k,) fila Excel de RawData-PCS
    pcs: list[int]  # PCS recorridos (C2)
    baterias: np.ndarray  # (k,p) C3 - módulos; NaN si la celda no se escribe
    ponderadas: np.ndarray  # (k,p) Calc!F:BN; NaN = celda vacía
    impacto: np.ndarray  # (k,p) aporte a C14; 0 si no aporta
    bo_evento_excusado_pa: np.ndarray  # (k,) Calc!BO (= PlantActivity!D, espejo; no pondera — F-37)
    exclusion: np.ndarray  # (k,p) valor de Exclusion_Matrix (0, 1, 2)
    factor_operacional: np.ndarray  # (k,) PlantActivity!C

    def ponderadas_o_cero(self) -> np.ndarray:
        """Tabla con celdas vacías = 0 (como la lee ``mcoDailyAvailability``)."""
        return np.nan_to_num(self.ponderadas, nan=0.0)

    def muestras(self) -> Iterator[MuestraDisponibilidad]:
        """Filas de ``muestra_pcs`` (una por fila en rango × PCS)."""
        for k in range(len(self.filas_procesadas)):
            for j, pcs in enumerate(self.pcs):
                b = self.baterias[k, j]
                escrita = not np.isnan(b)
                yield MuestraDisponibilidad(
                    numero_fila=int(self.numero_fila[k]),
                    serial=float(self.serial[k]),
                    numero_pcs=pcs,
                    baterias_indisponibles=float(b) if escrita else 0.0,
                    valor_exclusion=float(self.exclusion[k, j]),
                    factor_operacional=float(self.factor_operacional[k]),
                    baterias_ponderadas=float(self.ponderadas[k, j]) if escrita else 0.0,
                    impacto_rack_ponderado=float(self.impacto[k, j]),
                )

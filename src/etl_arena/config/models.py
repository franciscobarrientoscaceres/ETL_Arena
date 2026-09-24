"""Configuración de una corrida (R1). Ningún motor usa constantes de negocio fuera de aquí."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Literal

ModoHuecos = Literal["excel", "continuar"]
TipoCorrida = Literal["semanal", "cierre_mensual", "reproceso", "golden"]

MODOS_HUECOS: tuple[str, ...] = ("excel", "continuar")
TIPOS_CORRIDA: tuple[str, ...] = ("semanal", "cierre_mensual", "reproceso", "golden")
MAX_DIAS_DIARIO = 31  # Daily!C9:C39


class ErrorConfiguracion(ValueError):
    """Parámetros de corrida inválidos."""


@dataclass(frozen=True)
class ConfiguracionCalculo:
    id_corrida: str
    id_proyecto: int  # 1 = Arena
    version_algoritmo: str  # "availability-v1-excel-parity"
    nombre_proyecto: str
    fecha_inicio_proyecto: date
    total_pcs: int  # C2
    baterias_por_pcs: int  # C3 (y umbral literal "< 4" del VBA)
    racks_por_pcs: int  # literal 12 del VBA (racks por batería BAC)
    minutos_muestreo: int  # esperado; C23 se deriva de los datos (F-13)
    # parámetros KPI (cmdCalcAvailability)
    inicio_periodo: date  # C5
    fin_periodo: date  # C7 (inclusivo a nivel día)
    solo_tiempo_operacional: bool  # C21 <> "No" → pondera por PlantActivity!C
    aplicar_evento_excusable: bool  # C31 = "Yes" → aplica Exclusion_Matrix (F-37)
    # parámetros de eventos (mcoCreateList) — F-05
    inicio_periodo_eventos: date  # ListOfFaults!L2
    fin_periodo_eventos: date  # ListOfFaults!L4
    aplicar_evento_excusable_eventos: bool  # ListOfFaults!L14 = "Yes" → aplica Exclusion_Matrix (F-37)
    # Daily
    fin_diario: date  # última fecha de Daily!C9:C39 (F-09, F-33)
    modo_huecos: ModoHuecos
    tipo_corrida: TipoCorrida
    es_oficial: bool
    archivo_origen: str
    zona_horaria: str = "America/Santiago"  # documental (ADR-07)

    @property
    def total_racks(self) -> int:
        """C11 = 12 * C2 * C3."""
        return self.racks_por_pcs * self.total_pcs * self.baterias_por_pcs

    @property
    def dias_diario(self) -> int:
        return (self.fin_diario - self.inicio_periodo).days + 1

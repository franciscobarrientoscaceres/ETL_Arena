"""Anomalías de datos detectadas en ingesta, enriquecimiento y motores (R15, R16)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

Severidad = Literal["info", "advertencia", "error"]


@dataclass(frozen=True)
class Anomalia:
    tipo: str  # p. ej. "hueco", "dst_salto", "modulos_texto", "pa_sin_timestamp"
    severidad: Severidad
    numero_fila: int | None = None  # fila Excel (1-based)
    numero_pcs: int | None = None
    serial: float | None = None
    detalle: str = ""


class ErrorParidad(RuntimeError):
    """El libro contiene algo que haría fallar la macro VBA (modo paridad): la corrida se aborta."""

    def __init__(self, mensaje: str, anomalias: list[Anomalia] | None = None):
        super().__init__(mensaje)
        self.anomalias = anomalias or []

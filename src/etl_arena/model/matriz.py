"""Datos de entrada de los motores indexados por fila origen (ADR-02)."""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass
class MatrizPCS:
    """``RawData-PCS`` como matriz ``n filas × p PCS`` en el orden de la hoja.

    ``i = 0`` es la fila Excel 2 (la anterior es el encabezado, D-08).
    """

    numero_fila: np.ndarray  # (n,) int — fila Excel (2..)
    serial: np.ndarray  # (n,) float64 — RawData-PCS!A exacto
    marca_tiempo: np.ndarray  # (n,) datetime64[s] naive local
    modulos: np.ndarray  # (n,p) float64; NaN si vacío
    modulos_nulo: np.ndarray  # (n,p) bool
    falla: np.ndarray  # (n,p) object: str | float | None (tipo crudo)
    estado: np.ndarray  # (n,p) object
    advertencia: np.ndarray  # (n,p) object
    pcs: list[int]  # números de PCS en orden de encabezado
    encabezado_falla: list[str]  # texto de la fila 1 por PCS (columna FAULT)
    encabezado_modulos: list[str]  # texto de la fila 1 por PCS (columna NUMBER OF MODULES)
    # Fila posterior a la última leída (la A vacía donde corta el VBA): mcoCreateList la
    # consulta para cerrar el último evento. Por defecto, totalmente vacía.
    siguiente_modulos: np.ndarray | None = None  # (p,) float64; NaN si vacío
    siguiente_nulo: np.ndarray | None = None  # (p,) bool

    def __post_init__(self) -> None:
        n, p = len(self.serial), len(self.pcs)
        if self.siguiente_modulos is None:
            self.siguiente_modulos = np.full(p, np.nan)
        if self.siguiente_nulo is None:
            self.siguiente_nulo = np.ones(p, dtype=bool)
        for nombre in ("numero_fila", "marca_tiempo"):
            if getattr(self, nombre).shape != (n,):
                raise ValueError(f"{nombre}: forma {getattr(self, nombre).shape} != ({n},)")
        for nombre in ("modulos", "modulos_nulo", "falla", "estado", "advertencia"):
            if getattr(self, nombre).shape != (n, p):
                raise ValueError(f"{nombre}: forma {getattr(self, nombre).shape} != ({n}, {p})")
        if len(self.encabezado_falla) != p or len(self.encabezado_modulos) != p:
            raise ValueError("encabezados por PCS incompletos")

    @property
    def n(self) -> int:
        return len(self.serial)

    @property
    def p(self) -> int:
        return len(self.pcs)

    def modulos_disponibles(self, baterias_por_pcs: int) -> np.ndarray:
        """``ModulosDisponibles`` para persistencia: vacío = disponible (R5.3)."""
        return np.where(self.modulos_nulo, float(baterias_por_pcs), self.modulos)


@dataclass
class DatosActividad:
    """``PlantActivity`` alineada **por fila** con ``MatrizPCS`` (F-01): índice ``i`` ↔ misma fila Excel.

    Desde F-37 solo la columna C pondera (``solo_tiempo_operacional``, C21). La columna D ya no
    excusa: la exclusión viene de ``Exclusion_Matrix`` (``DatosExclusion``); D se conserva solo
    como espejo de ``Calculation-Availability!BO``, que la macro de septiembre sigue copiando.
    """

    numero_fila: np.ndarray  # (n,) int
    serial_pa: np.ndarray  # (n,) float64 — PlantActivity!B; NaN si vacío (F-31)
    factor_operacional: np.ndarray  # (n,) float64 — PlantActivity!C; vacío = 0
    evento_excusado_pa: np.ndarray  # (n,) float64 — PlantActivity!D (espejo de Calc!BO); vacío = 0
    porcentaje_soc: np.ndarray  # (n,) float64 — PlantActivity!I; NaN si vacío

    def __post_init__(self) -> None:
        n = len(self.numero_fila)
        for nombre in ("serial_pa", "factor_operacional", "evento_excusado_pa", "porcentaje_soc"):
            if getattr(self, nombre).shape != (n,):
                raise ValueError(f"{nombre}: forma {getattr(self, nombre).shape} != ({n},)")


VALORES_EXCLUSION = (0.0, 1.0, 2.0)


@dataclass
class DatosExclusion:
    """``Exclusion_Matrix`` alineada **por fila** con ``MatrizPCS`` (F-37).

    Por celda (fila, PCS): ``0``/vacío = sin evento de exclusión (EE); ``1`` = EE, el PCS se
    considera con todos sus módulos; ``2`` = EE, se consideran los módulos que tenía antes del
    inicio del evento. Solo pondera cuando ``aplicar_evento_excusable`` (C31 / L14 = "Yes").
    """

    numero_fila: np.ndarray  # (n,) int
    serial_em: np.ndarray  # (n,) float64 — Exclusion_Matrix!A; NaN si vacío
    valor: np.ndarray  # (n,p) float64 en {0, 1, 2}; vacío = 0
    baterias_previas: (
        np.ndarray
    )  # (n,p) float64 — baterías indisponibles consideradas antes del EE (valor 2); NaN si no aplica
    evento_excusado: np.ndarray  # (n,) float64 — columna "Excused Event" (resumen por fila, informativa); NaN si vacío
    comentario: np.ndarray  # (n,) object — columna "Comments" (causa del EE)

    def __post_init__(self) -> None:
        n = len(self.numero_fila)
        for nombre in ("serial_em", "evento_excusado", "comentario"):
            if getattr(self, nombre).shape != (n,):
                raise ValueError(f"{nombre}: forma {getattr(self, nombre).shape} != ({n},)")
        if self.valor.shape[0] != n or self.baterias_previas.shape != self.valor.shape:
            raise ValueError("valor/baterias_previas con forma inconsistente")

    @classmethod
    def vacia(cls, numero_fila: np.ndarray, p: int) -> DatosExclusion:
        """Sin ``Exclusion_Matrix`` (libros anteriores a agosto 2026): todo 0."""
        n = len(numero_fila)
        return cls(
            numero_fila.copy(),
            np.full(n, np.nan),
            np.zeros((n, p)),
            np.full((n, p), np.nan),
            np.full(n, np.nan),
            np.full(n, None, dtype=object),
        )

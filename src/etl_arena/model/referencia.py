"""Referencia Excel extraída tras las macros (R3.6) y filas de reconciliación (R13.9).

Solo estructuras: la extracción es ``workbook.referencia`` (Fase 4) y la comparación
``reconciliation`` (Fase 3); la persistencia las guarda en ``excel_reference_*`` y
``reconciliation_result``.
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from datetime import date, datetime


@dataclass
class EventoReferencia:
    """Fila ``B:I`` de ``ListOfFaults`` tal como la dejó la macro."""

    orden_excel: int
    b: float | None
    c: float | None
    d: float | None
    e: float | None
    f: str | None
    g: str | None
    h: float | None
    i: float | None


@dataclass
class DiaReferencia:
    """Fila ``B:G`` de ``Daily``."""

    dia_n: int
    dia: date | None
    d: float | None
    e: float | None
    f: float | None
    g: float | None


@dataclass
class ReferenciaExcel:
    corte: str
    archivo_libro: str
    hash_libro: str
    extraido_en: datetime
    c5: date
    c7: date
    l2: date
    l4: date
    c21: str | None = None
    c31: str | None = None
    l14: str | None = None
    daily_d5: date | None = None
    c12: int | None = None
    c14: float | None = None
    c16: float | None = None
    c19: float | None = None
    c23: float | None = None
    l10: float | None = None
    tabla: list[tuple[int, float, int, float]] = field(default_factory=list)  # (fila, serial E, PCS, valor F:BN)
    seriales_tabla: list[float] = field(default_factory=list)  # Calc!E4.. (una por fila de resultado)
    bo: list[float | None] = field(default_factory=list)  # Calc!BO por fila de resultado
    eventos: list[EventoReferencia] = field(default_factory=list)
    diario: list[DiaReferencia] = field(default_factory=list)
    resumen_codigos: list[tuple[str, float | None]] = field(default_factory=list)  # ListOfFaults!N:P
    id_referencia: str = field(default_factory=lambda: str(uuid.uuid4()))


@dataclass(frozen=True)
class FilaReconciliacion:
    nivel: int  # 1..5; 9 = invariante
    metrica: str
    clave: str | None
    valor_python: str | None
    valor_excel: str | None
    delta: float | None
    tolerancia: float | None
    aprobado: bool

"""Seriales de fecha de Excel (ADR-03, ADR-07): días desde 1899-12-30, hora local naive."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from fractions import Fraction

EPOCH_EXCEL = datetime(1899, 12, 30)
_DIA = timedelta(days=1)


def serial_a_datetime(s: float, redondear_a_segundo: bool = True) -> datetime:
    """Serial → ``datetime`` naive. Por defecto al segundo más cercano, porque los seriales
    de 15 min no son exactos en binario (``46120.010416666664`` = 2026-04-08 00:15:00).
    Solo para presentación y persistencia: los motores operan sobre el serial."""
    dt = EPOCH_EXCEL + timedelta(days=s)
    if redondear_a_segundo:
        dt = (dt + timedelta(microseconds=500_000)).replace(microsecond=0)
    return dt


def datetime_a_serial(v: date | datetime) -> float:
    """``date``/``datetime`` naive → serial Excel. Una ``date`` es la medianoche (entero exacto)."""
    if not isinstance(v, datetime):
        return float((v - EPOCH_EXCEL.date()).days)
    delta = v - EPOCH_EXCEL
    micros = (delta.days * 86400 + delta.seconds) * 1_000_000 + delta.microseconds
    return float(Fraction(micros, 86_400_000_000))  # double más cercano al valor exacto, como Excel


def fecha_vba_a_celda(s: float) -> float:
    """Serial que queda en una celda al escribir un valor VBA de tipo ``Date``.

    ``Sheet2.Cells(r, 1)`` es un ``Date`` (celda con formato fecha), así que
    ``A - C23/(24*60)`` también lo es; al asignarlo a una celda Excel lo convierte con
    resolución de un segundo. Evidencia: ``ListOfFaults!C`` del libro de septiembre guarda
    ``46266.010416666664`` (serial exacto de 00:15:00), no ``46266.01041666667`` (resultado
    crudo de la resta), y ``E = 24*(D-C) = 5.5000000001164153`` solo cuadra con el primero.
    Con datos cada 15 min los resultados caen en minutos exactos: redondear al segundo o al
    milisegundo da lo mismo.
    """
    return datetime_a_serial(serial_a_datetime(s, redondear_a_segundo=True))

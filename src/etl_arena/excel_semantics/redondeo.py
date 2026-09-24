"""``ROUND`` de Excel: mitad lejos de cero, sobre la representación de 15 dígitos."""

from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal


def redondear_excel(x: float, digitos: int) -> float:
    """Emula ``ROUND(x, digitos)``.

    Excel redondea mitad lejos de cero (no el redondeo bancario de ``round``) y trabaja
    sobre el valor con 15 dígitos significativos, por eso ``ROUND(2.675, 2) = 2.68``
    aunque el ``double`` sea 2.67499999… Solo se usa donde el libro usa ``ROUND`` (C23).
    """
    d = Decimal(f"{x:.15g}")
    return float(d.quantize(Decimal(1).scaleb(-digitos), rounding=ROUND_HALF_UP))

"""Resultados de la reconciliación Python vs Excel (R13.9)."""

from __future__ import annotations

import math
from dataclasses import dataclass, field

MAX_DISCREPANCIAS_POR_NIVEL = 500  # se cuentan todas; se guardan las primeras


@dataclass(frozen=True)
class Discrepancia:
    nivel: int
    metrica: str
    clave: str
    python: object
    excel: object
    delta: float | None
    tolerancia: float | None


@dataclass
class ResultadoNivel:
    nivel: int  # 1..5; 9 = invariantes
    nombre: str
    comparaciones: int = 0
    fallidas: int = 0
    max_delta: float | None = None
    discrepancias: list[Discrepancia] = field(default_factory=list)
    no_aplica: list[str] = field(default_factory=list)  # invariantes omitidos y por qué

    @property
    def aprobado(self) -> bool:
        return self.fallidas == 0

    def comparar(self, metrica: str, clave: str, python: object, excel: object, tolerancia: float | None) -> bool:
        """Registra una comparación. ``tolerancia`` None o 0 = igualdad exacta."""
        self.comparaciones += 1
        delta = None
        if isinstance(python, (int, float)) and isinstance(excel, (int, float)) and not isinstance(python, bool):
            python_f, excel_f = float(python), float(excel)
            if math.isnan(python_f) or math.isnan(excel_f):
                ok = math.isnan(python_f) and math.isnan(excel_f)
            else:
                delta = abs(python_f - excel_f)
                ok = delta <= (tolerancia or 0.0)
                self.max_delta = delta if self.max_delta is None else max(self.max_delta, delta)
        else:
            ok = python == excel
        if not ok:
            self.fallidas += 1
            if len(self.discrepancias) < MAX_DISCREPANCIAS_POR_NIVEL:
                self.discrepancias.append(Discrepancia(self.nivel, metrica, clave, python, excel, delta, tolerancia))
        return ok


@dataclass
class ReporteReconciliacion:
    id_corrida: str
    id_referencia: str | None
    niveles: list[ResultadoNivel]
    invariantes: ResultadoNivel | None
    sin_referencia: bool = False

    @property
    def aprobado_global(self) -> bool:
        """R13.9: la paridad falla si falla algún nivel 3–5."""
        return self.sin_referencia or all(n.aprobado for n in self.niveles if n.nivel in (3, 4, 5))

    @property
    def estado_corrida(self) -> str:
        return "success" if self.aprobado_global else "parity_failed"

    def nivel(self, n: int) -> ResultadoNivel:
        return next(x for x in [*self.niveles, *([self.invariantes] if self.invariantes else [])] if x.nivel == n)

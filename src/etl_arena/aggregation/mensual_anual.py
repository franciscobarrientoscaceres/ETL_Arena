"""KPI mensual oficial y acumulado anual (R10, F-20, D-06, D-07).

``Annual_AVA`` (filas 7..18 = meses 1..12)::

    E = D*24*60/15                         bloques de calendario (15 literal)
    G = IF(F<>"", 1-F/(Total_Racks*E), "")
    H = IF(F<>"", E+H_anterior, 0)         un mes sin F reinicia los bloques acumulados
    I = F + I_anterior
    J = IFERROR(1-I/(Total_Racks*H), "N/A")

D-07 (2026-09-24): el KPI mensual **oficial** usa ``BloquesMuestreo = C12`` (intervalos
existentes), no los bloques de calendario de la hoja. ``bloques_calendario`` queda para
emular ``Annual_AVA`` y para las filas históricas ``excel_manual``.
"""

from __future__ import annotations

import calendar
from collections.abc import Iterable
from dataclasses import dataclass

from etl_arena.availability import ResultadoDisponibilidad
from etl_arena.config import ConfiguracionCalculo
from etl_arena.excel_semantics import datetime_a_serial
from etl_arena.model import KpiMensual

MINUTOS_ANNUAL_AVA = 15  # literal de la fórmula E de Annual_AVA


class ErrorMensual(ValueError):
    """El período de la corrida no corresponde a un único mes."""


def bloques_calendario(dias_mes: float) -> float:
    """``Annual_AVA!E = D*24*60/15``."""
    return dias_mes * 24 * 60 / MINUTOS_ANNUAL_AVA


def registrar_mes_oficial(res: ResultadoDisponibilidad, cfg: ConfiguracionCalculo) -> KpiMensual:
    """KPI del mes de la corrida (R10.2). El período debe empezar el día 1 y no salir del mes.

    ``DiasMes`` (informativo): días calendario si el período cubre el mes completo; si no,
    ``serial(última fila procesada) − serial(día 1)`` (sep-2026: 20.59375, como ``Annual_AVA!D15``).
    """
    ini, fin = cfg.inicio_periodo, cfg.fin_periodo
    if ini.day != 1 or (fin.year, fin.month) != (ini.year, ini.month):
        raise ErrorMensual(f"período {ini}..{fin} no es un mes (día 1 → mismo mes)")
    dias_calendario = calendar.monthrange(ini.year, ini.month)[1]
    if fin.day == dias_calendario:
        dias_mes = float(dias_calendario)
    elif res.bloques_muestreo:
        dias_mes = float(res.serial[-1]) - datetime_a_serial(ini)
    else:
        dias_mes = 0.0
    return KpiMensual(
        anio=ini.year,
        mes=ini.month,
        dias_mes=dias_mes,
        bloques_muestreo=float(res.bloques_muestreo),
        bloques_racks_indisponibles=res.bloques_racks_indisponibles,
        origen="corrida",
        id_corrida=cfg.id_corrida,
    )


@dataclass(frozen=True)
class FilaAnual:
    anio: int
    mes: int
    dias_mes: float  # D
    bloques_muestreo: float  # E
    bloques_racks_indisponibles: float | None  # F
    disponibilidad_mensual: float | None  # G
    bloques_muestreo_acumulados: float  # H
    bloques_indisponibles_acumulados: float  # I
    disponibilidad_acumulada: float | None  # J
    disponibilidad_contractual: float  # K


def calcular_anual(
    meses: Iterable[KpiMensual], total_racks: int, anio: int, mes_inicio_acumulado: int
) -> list[FilaAnual]:
    """Filas de ``Annual_AVA`` desde ``mes_inicio_acumulado`` (D-06: julio 2026) hasta el último mes con dato."""
    por_mes = {k.mes: k for k in meses if k.anio == anio}
    if not por_mes:
        return []
    salida: list[FilaAnual] = []
    h_ant, i_ant = 0.0, 0.0
    for mes in range(mes_inicio_acumulado, max(por_mes) + 1):
        k = por_mes.get(mes)
        f = k.bloques_racks_indisponibles if k else None
        e = k.bloques_muestreo if k else bloques_calendario(calendar.monthrange(anio, mes)[1])
        g = 1 - f / (total_racks * e) if f is not None else None
        h = e + h_ant if f is not None else 0.0
        i = (f if f is not None else 0.0) + i_ant
        try:
            j = 1 - i / (total_racks * h)
        except ZeroDivisionError:
            j = None
        salida.append(
            FilaAnual(
                anio,
                mes,
                k.dias_mes if k else float(calendar.monthrange(anio, mes)[1]),
                e,
                f,
                g,
                h,
                i,
                j,
                k.disponibilidad_contractual if k else 0.98,
            )
        )
        h_ant, i_ant = h, i
    return salida

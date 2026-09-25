"""Construcción y validación de ``ConfiguracionCalculo`` (R1.3, R1.4)."""

from __future__ import annotations

import uuid
from dataclasses import fields
from datetime import date, datetime
from typing import Any

from etl_arena.config.defaults import CONFIG_POR_DEFECTO
from etl_arena.config.models import (
    MAX_DIAS_DIARIO,
    MODOS_HUECOS,
    TIPOS_CORRIDA,
    ConfiguracionCalculo,
    ErrorConfiguracion,
)

_ENTEROS_POSITIVOS = ("id_proyecto", "total_pcs", "baterias_por_pcs", "racks_por_pcs", "minutos_muestreo")
_FECHAS = (
    "fecha_inicio_proyecto",
    "inicio_periodo",
    "fin_periodo",
    "inicio_periodo_eventos",
    "fin_periodo_eventos",
    "fin_diario",
    "inicio_acumulado_anual",
)
_BOOLEANOS = ("solo_tiempo_operacional", "aplicar_evento_excusable", "aplicar_evento_excusable_eventos", "es_oficial")


def _como_fecha(nombre: str, v: Any) -> date:
    if isinstance(v, datetime):
        if v.time() != datetime.min.time():
            raise ErrorConfiguracion(f"{nombre} debe ser una fecha sin hora: {v!r}")
        return v.date()
    if isinstance(v, date):
        return v
    raise ErrorConfiguracion(f"{nombre} debe ser date: {v!r}")


def construir_config(**overrides: Any) -> ConfiguracionCalculo:
    """Aplica ``CONFIG_POR_DEFECTO`` + overrides y valida.

    ``inicio_periodo`` y ``fin_periodo`` son obligatorios. Si no se informan, los parámetros
    de eventos toman los del KPI (R1.4) y ``fin_diario`` toma ``fin_periodo``.
    """
    valores: dict[str, Any] = dict(CONFIG_POR_DEFECTO)
    valores.update(overrides)
    nombres = {f.name for f in fields(ConfiguracionCalculo)}
    desconocidos = set(valores) - nombres
    if desconocidos:
        raise ErrorConfiguracion(f"parámetros desconocidos: {sorted(desconocidos)}")
    for obligatorio in ("inicio_periodo", "fin_periodo"):
        if valores.get(obligatorio) is None:
            raise ErrorConfiguracion(f"falta {obligatorio}")

    valores.setdefault("id_corrida", str(uuid.uuid4()))
    for clave, desde in (
        ("inicio_periodo_eventos", "inicio_periodo"),
        ("fin_periodo_eventos", "fin_periodo"),
        ("fin_diario", "fin_periodo"),
    ):
        if valores.get(clave) is None:
            valores[clave] = valores[desde]

    for nombre in _FECHAS:
        valores[nombre] = _como_fecha(nombre, valores[nombre])
    for nombre in _ENTEROS_POSITIVOS:
        v = valores[nombre]
        if isinstance(v, bool) or not isinstance(v, int) or v <= 0:
            raise ErrorConfiguracion(f"{nombre} debe ser entero positivo: {v!r}")
    for nombre in _BOOLEANOS:
        if not isinstance(valores[nombre], bool):
            raise ErrorConfiguracion(f"{nombre} debe ser bool: {valores[nombre]!r}")
    if valores["modo_huecos"] not in MODOS_HUECOS:
        raise ErrorConfiguracion(f"modo_huecos inválido: {valores['modo_huecos']!r}")
    if valores["tipo_corrida"] not in TIPOS_CORRIDA:
        raise ErrorConfiguracion(f"tipo_corrida inválido: {valores['tipo_corrida']!r}")

    if valores["fin_periodo"] < valores["inicio_periodo"]:
        raise ErrorConfiguracion("fin_periodo < inicio_periodo")
    if valores["fin_periodo_eventos"] < valores["inicio_periodo_eventos"]:
        raise ErrorConfiguracion("fin_periodo_eventos < inicio_periodo_eventos")
    dias = (valores["fin_diario"] - valores["inicio_periodo"]).days + 1
    if not 1 <= dias <= MAX_DIAS_DIARIO:
        raise ErrorConfiguracion(f"Daily admite 1..{MAX_DIAS_DIARIO} días desde inicio_periodo; se pidieron {dias}")

    return ConfiguracionCalculo(**valores)

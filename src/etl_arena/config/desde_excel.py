"""Parámetros de corrida a partir de las celdas del libro (R1.5, F-05, F-09, F-12, F-33).

Recibe los valores crudos ya leídos (``{"Hoja!Ref": valor}``); la lectura del ``.xlsm``
vive en ``etl_arena.ingestion`` para que ``config`` no dependa del formato de archivo.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from typing import Any

from etl_arena.config.models import MAX_DIAS_DIARIO, ErrorConfiguracion
from etl_arena.excel_semantics import flag_no, flag_si, serial_a_datetime

HOJA_CALC = "Calculation-Availability"
HOJA_EVENTOS = "ListOfFaults"
HOJA_DIARIO = "Daily"

CELDAS_PARAMETROS: tuple[str, ...] = (
    f"{HOJA_CALC}!C2",
    f"{HOJA_CALC}!C3",
    f"{HOJA_CALC}!C5",
    f"{HOJA_CALC}!C7",
    f"{HOJA_CALC}!C11",
    f"{HOJA_CALC}!C21",
    f"{HOJA_CALC}!C31",
    f"{HOJA_EVENTOS}!L2",
    f"{HOJA_EVENTOS}!L4",
    f"{HOJA_EVENTOS}!L14",
    f"{HOJA_DIARIO}!D5",
) + tuple(f"{HOJA_DIARIO}!C{fila}" for fila in range(9, 9 + MAX_DIAS_DIARIO))


@dataclass
class ParametrosExcel:
    valores: dict[str, Any]
    notas: list[str] = field(default_factory=list)


def _entero(celdas: Mapping[str, Any], ref: str) -> int:
    v = celdas.get(ref)
    if not isinstance(v, (int, float)) or isinstance(v, bool) or float(v) != int(v):
        raise ErrorConfiguracion(f"{ref} debe ser entero: {v!r}")
    return int(v)


def _fecha(celdas: Mapping[str, Any], ref: str):
    v = celdas.get(ref)
    if not isinstance(v, (int, float)) or isinstance(v, bool):
        raise ErrorConfiguracion(f"{ref} debe ser una fecha (serial): {v!r}")
    dt = serial_a_datetime(float(v))
    if dt.time() != dt.min.time():
        raise ErrorConfiguracion(f"{ref} tiene hora; se espera una fecha: {v!r}")
    return dt.date()


def parametros_desde_celdas(celdas: Mapping[str, Any]) -> ParametrosExcel:
    """Traduce las celdas de parámetros a argumentos de ``construir_config``.

    * Flags con la semántica asimétrica del VBA: ``C21 <> "No"``; ``C31``/``L14`` ``= "Yes"``.
    * ``racks_por_pcs`` = ``C11 / (C2·C3)`` (la hoja calcula ``C11 = 12*C2*C3``).
    * ``fin_diario`` = última fecha contigua de ``Daily!C9:C39``: ``mcoDailyAvailability``
      corta en la primera ``C`` vacía, no en ``D5`` (F-33). Si ``D5`` difiere se anota.
    """
    total_pcs = _entero(celdas, f"{HOJA_CALC}!C2")
    baterias = _entero(celdas, f"{HOJA_CALC}!C3")
    valores: dict[str, Any] = {
        "total_pcs": total_pcs,
        "baterias_por_pcs": baterias,
        "inicio_periodo": _fecha(celdas, f"{HOJA_CALC}!C5"),
        "fin_periodo": _fecha(celdas, f"{HOJA_CALC}!C7"),
        "solo_tiempo_operacional": not flag_no(celdas.get(f"{HOJA_CALC}!C21")),
        "aplicar_evento_excusable": flag_si(celdas.get(f"{HOJA_CALC}!C31")),
        "inicio_periodo_eventos": _fecha(celdas, f"{HOJA_EVENTOS}!L2"),
        "fin_periodo_eventos": _fecha(celdas, f"{HOJA_EVENTOS}!L4"),
        "aplicar_evento_excusable_eventos": flag_si(celdas.get(f"{HOJA_EVENTOS}!L14")),
    }
    notas: list[str] = []

    c11 = celdas.get(f"{HOJA_CALC}!C11")
    if isinstance(c11, (int, float)) and not isinstance(c11, bool):
        racks = c11 / (total_pcs * baterias)
        if racks != int(racks):
            raise ErrorConfiguracion(f"C11={c11} no es múltiplo de C2*C3={total_pcs * baterias}")
        valores["racks_por_pcs"] = int(racks)

    fechas_diario = []
    for fila in range(9, 9 + MAX_DIAS_DIARIO):
        v = celdas.get(f"{HOJA_DIARIO}!C{fila}")
        if v is None or v == "":
            break
        fechas_diario.append(_fecha(celdas, f"{HOJA_DIARIO}!C{fila}"))
    if fechas_diario:
        valores["fin_diario"] = fechas_diario[-1]
        d5 = celdas.get(f"{HOJA_DIARIO}!D5")
        if (
            isinstance(d5, (int, float))
            and not isinstance(d5, bool)
            and _fecha(celdas, f"{HOJA_DIARIO}!D5") != fechas_diario[-1]
        ):
            notas.append(
                f"Daily!D5 ({_fecha(celdas, f'{HOJA_DIARIO}!D5')}) difiere de la última fecha de "
                f"Daily!C ({fechas_diario[-1]}); la macro usa Daily!C (F-33)"
            )
    return ParametrosExcel(valores, notas)

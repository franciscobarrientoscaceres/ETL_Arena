"""Estructuras de datos compartidas por ingesta, motores y persistencia."""

from etl_arena.model.anomalias import Anomalia, ErrorParidad, Severidad
from etl_arena.model.matriz import VALORES_EXCLUSION, DatosActividad, DatosExclusion, MatrizPCS
from etl_arena.model.referencia import DiaReferencia, EventoReferencia, FilaReconciliacion, ReferenciaExcel
from etl_arena.model.registros import (
    DiaDisponibilidad,
    EventoFalla,
    KpiMensual,
    MuestraDisponibilidad,
    RegistroLista,
)

__all__ = [
    "Anomalia",
    "DatosActividad",
    "DatosExclusion",
    "DiaDisponibilidad",
    "DiaReferencia",
    "ErrorParidad",
    "EventoFalla",
    "EventoReferencia",
    "FilaReconciliacion",
    "KpiMensual",
    "MatrizPCS",
    "MuestraDisponibilidad",
    "ReferenciaExcel",
    "RegistroLista",
    "Severidad",
    "VALORES_EXCLUSION",
]

"""Única fuente de las reglas implícitas de Excel/VBA (ADR-04, R17)."""

from etl_arena.excel_semantics.celdas import (
    ExcelTypeMismatch,
    a_numero_vba,
    comparar_variants,
    cstr_vba,
    es_vacio,
    flag_no,
    flag_si,
    igual_numero,
    igual_texto,
    menor_que,
)
from etl_arena.excel_semantics.fechas import EPOCH_EXCEL, datetime_a_serial, fecha_vba_a_celda, serial_a_datetime
from etl_arena.excel_semantics.redondeo import redondear_excel
from etl_arena.excel_semantics.texto import codigo_falla_excel, texto_excel

__all__ = [
    "EPOCH_EXCEL",
    "ExcelTypeMismatch",
    "a_numero_vba",
    "codigo_falla_excel",
    "comparar_variants",
    "cstr_vba",
    "datetime_a_serial",
    "es_vacio",
    "fecha_vba_a_celda",
    "flag_no",
    "flag_si",
    "igual_numero",
    "igual_texto",
    "menor_que",
    "redondear_excel",
    "serial_a_datetime",
    "texto_excel",
]

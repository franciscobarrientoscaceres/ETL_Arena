"""Comparaciones de celdas con la semántica de VBA (R17, F-12, F-16, F-21).

Una celda leída con ``Cells(r, c)`` es un ``Variant``. En el XML del libro una celda
vacía llega como ``None`` (``Empty`` en VBA); un texto vacío llega como ``""``. Son
distintos para VBA: ``Empty < 4`` vale 0 < 4, mientras que ``"" < 4`` es *Type mismatch*.

Reglas emuladas (VBA Language Reference, "Comparison operators"):

* Variant vs literal numérico: numérico → comparación numérica; ``Empty`` → 0;
  ``Boolean`` → ``True = -1``; texto convertible a número → numérica; texto no
  convertible → ``ExcelTypeMismatch``.
* Variant vs literal ``String``: comparación de texto, binaria (el módulo no declara
  ``Option Compare Text``); el Variant se convierte con ``CStr``.
* Variant vs Variant: ambos numéricos → numérica; ``Empty`` vale 0 frente a un número y
  ``""`` frente a un texto; número vs texto → el número es **menor** (sin error).

Los operadores ``And``/``Or`` de VBA no cortocircuitan: quien emule una condición debe
evaluar todos los operandos (p. ej. la fila de encabezado en ``mcoCreateList``, F-11).
"""

from __future__ import annotations

import re

Numero = int | float

# Texto que VBA convierte a número (sin "inf", "nan" ni separadores "_" que acepta float()).
_TEXTO_NUMERICO = re.compile(r"^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?\s*$")


class ExcelTypeMismatch(TypeError):
    """Error 13 de VBA: el valor de la celda no se puede comparar como número."""


def es_vacio(v: object) -> bool:
    """``Cells(...) = ""`` — verdadero para una celda vacía (``None``) o texto vacío."""
    return v is None or v == ""


def _es_numero(v: object) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def a_numero_vba(v: object) -> float:
    """Valor numérico que VBA usa al comparar ``v`` con un literal numérico.

    La conversión de texto usa punto decimal (el libro no trae números como texto; el
    data contract rechaza texto en columnas numéricas, así que el separador de la
    configuración regional de Windows no interviene).
    """
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return -1.0 if v else 0.0
    if _es_numero(v):
        return float(v)
    if isinstance(v, str):
        if not _TEXTO_NUMERICO.match(v):
            raise ExcelTypeMismatch(f"no numérico: {v!r}")
        return float(v)
    raise ExcelTypeMismatch(f"tipo no soportado: {type(v).__name__}")


def igual_numero(v: object, n: Numero) -> bool:
    """``Cells(...) = n`` con ``n`` literal numérico."""
    return a_numero_vba(v) == n


def menor_que(v: object, n: Numero) -> bool:
    """``Cells(...) < n`` con ``n`` literal numérico."""
    return a_numero_vba(v) < n


def cstr_vba(v: object) -> str:
    """``CStr`` de un Variant, usado en comparaciones contra un literal ``String``."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "True" if v else "False"
    if _es_numero(v):
        x = float(v)
        if x.is_integer() and abs(x) < 1e15:
            return str(int(x))
        return f"{x:.15g}"
    return str(v)


def igual_texto(v: object, s: str) -> bool:
    """``Cells(...) = "literal"``: comparación de texto binaria (sensible a mayúsculas)."""
    return cstr_vba(v) == s


def flag_si(v: object) -> bool:
    """``Cells(...) = "Yes"`` (C31, L14): solo ``"Yes"`` exacto activa el flag."""
    return igual_texto(v, "Yes")


def flag_no(v: object) -> bool:
    """``Cells(...) = "No"`` (C21): el factor operacional se aplica si esto es falso."""
    return igual_texto(v, "No")


def comparar_variants(a: object, b: object) -> int:
    """Compara dos celdas (Variant vs Variant). Devuelve -1, 0 o 1."""
    a_txt = isinstance(a, str)
    b_txt = isinstance(b, str)
    if a is None and b is None:
        return 0
    if a is None:
        a = "" if b_txt else 0.0
        a_txt = b_txt
    if b is None:
        b = "" if a_txt else 0.0
        b_txt = a_txt
    if a_txt and b_txt:
        return (a > b) - (a < b)
    if a_txt:  # texto vs número: el número es menor
        return 1
    if b_txt:
        return -1
    x, y = a_numero_vba(a), a_numero_vba(b)
    return (x > y) - (x < y)

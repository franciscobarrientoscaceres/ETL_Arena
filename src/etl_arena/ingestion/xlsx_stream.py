"""Lectura streaming de valores cacheados de un ``.xlsx``/``.xlsm`` (ADR-03, ADR-05).

Recorre ``xl/worksheets/sheetN.xml`` con ``iterparse`` y entrega los valores crudos:
``float`` para números y fechas (serial exacto del ``<v>``), ``str``, ``bool``,
``ErrorCelda`` o ``None`` (celda vacía). No usa ``openpyxl``: su conversión a
``datetime`` pierde el serial exacto.
"""

from __future__ import annotations

import re
import zipfile
from collections.abc import Iterable, Iterator
from pathlib import Path
from xml.etree.ElementTree import iterparse

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_REF = re.compile(r"([A-Z]+)(\d+)")


class ErrorLibro(RuntimeError):
    """Archivo u hoja inexistente o ilegible (R4.3)."""


class ErrorCelda(str):
    """Valor de error de Excel (``#VALUE!``, ``#N/A``…)."""


def indice_columna(letras: str) -> int:
    """``"A"`` → 1, ``"E"`` → 5, ``"IK"`` → 245."""
    n = 0
    for ch in letras:
        n = n * 26 + (ord(ch) - 64)
    return n


def letras_columna(indice: int) -> str:
    s = ""
    while indice:
        indice, r = divmod(indice - 1, 26)
        s = chr(65 + r) + s
    return s


def separar_ref(ref: str) -> tuple[int, int]:
    """``"C12"`` → ``(12, 3)``."""
    m = _REF.fullmatch(ref)
    if not m:
        raise ValueError(f"referencia inválida: {ref!r}")
    return int(m.group(2)), indice_columna(m.group(1))


class LibroXlsx:
    """Libro abierto en modo lectura. Usar como context manager."""

    def __init__(self, ruta: str | Path):
        self.ruta = Path(ruta)
        try:
            self._zip = zipfile.ZipFile(self.ruta)
            self._compartidas = self._leer_compartidas()
            self._hojas = self._leer_mapa_hojas()
        except (OSError, zipfile.BadZipFile, KeyError) as exc:
            raise ErrorLibro(f"no se puede leer {self.ruta}: {exc}") from exc

    def __enter__(self) -> LibroXlsx:
        return self

    def __exit__(self, *exc) -> None:
        self.cerrar()

    def cerrar(self) -> None:
        self._zip.close()

    @property
    def hojas(self) -> list[str]:
        return list(self._hojas)

    def _leer_compartidas(self) -> list[str]:
        if "xl/sharedStrings.xml" not in self._zip.namelist():
            return []
        salida: list[str] = []
        with self._zip.open("xl/sharedStrings.xml") as fh:
            for _, el in iterparse(fh):
                if el.tag == NS + "si":
                    salida.append("".join(t.text or "" for t in el.iter(NS + "t")))
                    el.clear()
        return salida

    def _leer_mapa_hojas(self) -> dict[str, str]:
        rels: dict[str, str] = {}
        with self._zip.open("xl/_rels/workbook.xml.rels") as fh:
            for _, el in iterparse(fh):
                if el.tag.endswith("Relationship"):
                    rels[el.get("Id")] = el.get("Target")
        hojas: dict[str, str] = {}
        with self._zip.open("xl/workbook.xml") as fh:
            for _, el in iterparse(fh):
                if el.tag == NS + "sheet":
                    destino = rels[el.get(_REL_NS + "id")].lstrip("/")
                    hojas[el.get("name")] = destino if destino.startswith("xl/") else "xl/" + destino
        return hojas

    def _valor(self, c) -> object:
        t = c.get("t")
        if t == "inlineStr":
            return "".join(x.text or "" for x in c.iter(NS + "t"))
        v = c.find(NS + "v")
        if v is None or v.text is None:
            return None
        if t == "s":
            return self._compartidas[int(v.text)]
        if t == "str":
            return v.text
        if t == "e":
            return ErrorCelda(v.text)
        if t == "b":
            return v.text == "1"
        return float(v.text)

    def filas(
        self, hoja: str, min_fila: int = 1, max_fila: int | None = None, max_col: int | None = None
    ) -> Iterator[tuple[int, dict[int, object]]]:
        """Itera ``(fila, {col: valor})`` en orden de la hoja, solo celdas con valor."""
        if hoja not in self._hojas:
            raise ErrorLibro(f"hoja inexistente: {hoja!r} en {self.ruta.name}")
        with self._zip.open(self._hojas[hoja]) as fh:
            fila_actual, celdas = None, {}
            for _, el in iterparse(fh):
                tag = el.tag
                if tag == NS + "c":
                    fila, col = separar_ref(el.get("r"))
                    if (max_col is None or col <= max_col) and fila >= min_fila:
                        val = self._valor(el)
                        if val is not None:
                            fila_actual = fila
                            celdas[col] = val
                    el.clear()
                elif tag == NS + "row":
                    r = int(el.get("r"))
                    el.clear()
                    if max_fila is not None and r > max_fila:
                        break
                    if celdas:
                        yield fila_actual, celdas
                        celdas = {}

    def celdas(self, refs: Iterable[str]) -> dict[str, object]:
        """Lee celdas sueltas ``"Hoja!A1"``; las inexistentes quedan en ``None``."""
        refs = list(refs)
        por_hoja: dict[str, dict[tuple[int, int], str]] = {}
        for ref in refs:
            hoja, celda = ref.rsplit("!", 1)
            por_hoja.setdefault(hoja, {})[separar_ref(celda)] = ref
        salida: dict[str, object] = dict.fromkeys(refs)
        for hoja, buscadas in por_hoja.items():
            max_fila = max(f for f, _ in buscadas)
            max_col = max(c for _, c in buscadas)
            for fila, valores in self.filas(hoja, max_fila=max_fila, max_col=max_col):
                for col, val in valores.items():
                    ref = buscadas.get((fila, col))
                    if ref is not None:
                        salida[ref] = val
        return salida

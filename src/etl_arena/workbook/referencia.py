"""Referencia Excel: valores que dejaron las macros en el libro, leídos por XML (R3.6, ADR-05).

Sin COM: lee los valores cacheados del ``.xlsm`` ya guardado. Layout oficial de septiembre 2026
(el mismo del maestro v1.1, D-19); el libro de agosto (F-38) usa otras celdas y no se soporta.
"""

from __future__ import annotations

import hashlib
from datetime import datetime
from pathlib import Path

from etl_arena.excel_semantics import serial_a_datetime, texto_excel
from etl_arena.ingestion.xlsx_stream import LibroXlsx
from etl_arena.model import DiaReferencia, EventoReferencia, ReferenciaExcel

HOJA_CALC, HOJA_EVENTOS, HOJA_DIARIO = "Calculation-Availability", "ListOfFaults", "Daily"


def _num(v) -> float | None:
    return v if isinstance(v, float) else None


def _fecha(v):
    return serial_a_datetime(v).date() if isinstance(v, float) else None


def _txt(v) -> str | None:
    """Mismo formato de número → texto que los motores (``texto_excel``), para no inventar
    discrepancias en ``ListOfFaults!G`` con descripciones numéricas (Checkpoint C-3)."""
    return None if v is None else texto_excel(v)


def sha256_archivo(ruta: str | Path) -> str:
    h = hashlib.sha256()
    with open(ruta, "rb") as fh:
        for bloque in iter(lambda: fh.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def extraer_referencia(ruta: str | Path, corte: str = "") -> ReferenciaExcel:
    """Lee C5/C7/C12/C14/C16/C19/C21/C23/C31, L2/L4/L10/L14, ``Daily!D5``, la tabla E4:BO,
    ``ListOfFaults!B6:I`` completa y N:P, y ``Daily!B9:G39``."""
    ruta = Path(ruta)
    with LibroXlsx(ruta) as libro:
        p = libro.celdas(
            [f"{HOJA_CALC}!C{f}" for f in (2, 5, 7, 12, 14, 16, 19, 21, 23, 31)]
            + [f"{HOJA_EVENTOS}!L{f}" for f in (2, 4, 10, 14)]
            + [f"{HOJA_DIARIO}!D5"]
        )
        total_pcs = int(p[f"{HOJA_CALC}!C2"])
        col_bo = 6 + total_pcs

        tabla, seriales, bo = [], [], []
        for fila, c in libro.filas(HOJA_CALC, min_fila=4, max_col=col_bo):
            e = c.get(5)
            if not isinstance(e, float):
                if seriales:
                    break
                continue
            seriales.append(e)
            bo.append(_num(c.get(col_bo)))
            for j in range(1, total_pcs + 1):
                v = c.get(5 + j)
                if isinstance(v, float):
                    tabla.append((fila, e, j, v))

        eventos, resumen = [], []
        for fila, c in libro.filas(HOJA_EVENTOS, min_fila=6, max_col=17):
            if c.get(2) is not None:
                eventos.append(
                    EventoReferencia(
                        fila - 5,
                        _num(c.get(2)),
                        _num(c.get(3)),
                        _num(c.get(4)),
                        _num(c.get(5)),
                        _txt(c.get(6)),
                        _txt(c.get(7)),
                        _num(c.get(8)),
                        _num(c.get(9)),
                    )
                )
            if fila <= 172 and c.get(14) is not None:
                resumen.append((_txt(c.get(14)), _num(c.get(16))))

        diario = []
        for fila, c in libro.filas(HOJA_DIARIO, min_fila=9, max_fila=39, max_col=7):
            if isinstance(c.get(3), float):
                diario.append(
                    DiaReferencia(
                        int(c[2]) if isinstance(c.get(2), float) else fila - 8,
                        _fecha(c[3]),
                        _num(c.get(4)),
                        _num(c.get(5)),
                        _num(c.get(6)),
                        _num(c.get(7)),
                    )
                )

    c12 = p[f"{HOJA_CALC}!C12"]
    return ReferenciaExcel(
        corte=corte or ruta.stem[:50],
        archivo_libro=ruta.name,
        hash_libro=sha256_archivo(ruta),
        extraido_en=datetime.now().replace(microsecond=0),
        c5=_fecha(p[f"{HOJA_CALC}!C5"]),
        c7=_fecha(p[f"{HOJA_CALC}!C7"]),
        l2=_fecha(p[f"{HOJA_EVENTOS}!L2"]),
        l4=_fecha(p[f"{HOJA_EVENTOS}!L4"]),
        c21=_txt(p[f"{HOJA_CALC}!C21"]),
        c31=_txt(p[f"{HOJA_CALC}!C31"]),
        l14=_txt(p[f"{HOJA_EVENTOS}!L14"]),
        daily_d5=_fecha(p[f"{HOJA_DIARIO}!D5"]),
        c12=int(c12) if isinstance(c12, float) else None,
        c14=_num(p[f"{HOJA_CALC}!C14"]),
        c16=_num(p[f"{HOJA_CALC}!C16"]),
        c19=_num(p[f"{HOJA_CALC}!C19"]),
        c23=_num(p[f"{HOJA_CALC}!C23"]),
        l10=_num(p[f"{HOJA_EVENTOS}!L10"]),
        tabla=tabla,
        seriales_tabla=seriales,
        bo=bo,
        eventos=eventos,
        diario=diario,
        resumen_codigos=resumen,
    )

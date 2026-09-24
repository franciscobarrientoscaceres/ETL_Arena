"""Fábricas de datos de prueba: libros .xlsx sintéticos y matrices en memoria."""

from __future__ import annotations

import zipfile
from collections.abc import Sequence
from datetime import date, datetime, timedelta
from pathlib import Path
from xml.sax.saxutils import escape

import numpy as np

from etl_arena.config import construir_config
from etl_arena.excel_semantics import datetime_a_serial, serial_a_datetime
from etl_arena.ingestion import letras_columna
from etl_arena.model import DatosActividad, DatosExclusion, MatrizPCS

CAMPOS_TXT = (
    "GEN3 HEx CURRENT FAULT",
    "GEN3 HEx CURRENT STATUS",
    "GEN3 HEx CURRENT WARNING",
    "HEM-k NUMBER OF MODULES",
)
SERIAL_SEP_1 = 46266.0
BLOQUE = 15 / 1440


def serial_min(minutos: int, dia: date = date(2026, 9, 1)) -> float:
    """Serial exacto (como lo guarda Excel) de ``dia 00:00 + minutos``."""
    return datetime_a_serial(datetime.combine(dia, datetime.min.time()) + timedelta(minutes=minutos))


def encabezado_raw(total_pcs: int) -> list[str]:
    fila = ["Date/time"]
    for k in range(1, total_pcs + 1):
        fila += [f"Arena - PCS {k:02d} - POWERELECTRONICS {c}" for c in CAMPOS_TXT]
    return fila


def config_prueba(total_pcs: int = 2, **cambios):
    base = {"inicio_periodo": date(2026, 9, 1), "fin_periodo": date(2026, 9, 1), "total_pcs": total_pcs}
    base.update(cambios)
    return construir_config(**base)


def _xml_hoja(filas: dict[int, Sequence[object]]) -> str:
    """Hoja con celdas numéricas (``repr``: 17 dígitos, como Excel) y texto inline."""
    partes = []
    for r in sorted(filas):
        celdas = []
        for c, v in enumerate(filas[r], start=1):
            if v is None:
                continue
            ref = f"{letras_columna(c)}{r}"
            if isinstance(v, bool):
                celdas.append(f'<c r="{ref}" t="b"><v>{int(v)}</v></c>')
            elif isinstance(v, (int, float)):
                celdas.append(f'<c r="{ref}"><v>{float(v)!r}</v></c>')
            else:
                celdas.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(str(v))}</t></is></c>')
        if celdas:
            partes.append(f'<row r="{r}">{"".join(celdas)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(partes)}</sheetData></worksheet>"
    )


def escribir_xlsx(ruta: Path, hojas: dict[str, dict[int, Sequence[object]]]) -> Path:
    """Escritor mínimo de .xlsx (solo lo que lee ``xlsx_stream``)."""
    ns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    sheets = "".join(f'<sheet name="{escape(n)}" sheetId="{i}" r:id="rId{i}"/>' for i, n in enumerate(hojas, 1))
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="{ns_rel}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, len(hojas) + 1)
    )
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "xl/workbook.xml",
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'xmlns:r="{ns_rel}"><sheets>{sheets}</sheets></workbook>',
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{rels}</Relationships>",
        )
        for i, filas in enumerate(hojas.values(), 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", _xml_hoja(filas))
    return ruta


def crear_libro(
    ruta: Path,
    filas: Sequence[Sequence[object]],
    total_pcs: int = 2,
    actividad: dict[int, Sequence[object]] | None = None,
    encabezado: Sequence[object] | None = None,
    exclusion: dict[int, Sequence[object]] | None = None,
) -> Path:
    """``filas``: listas completas a partir de la columna A (fila Excel 2 en adelante);
    ``None`` deja la celda vacía y una fila ``[]`` queda vacía. ``actividad``: fila Excel →
    valores desde la columna A de PlantActivity (A vacía, B ts, C op, D exc…). ``exclusion``:
    fila Excel → valores desde la columna A de Exclusion_Matrix (A ts, PCS01…, Excused Event,
    Comments); si es ``None`` el libro no trae la hoja."""
    raw = {1: list(encabezado) if encabezado is not None else encabezado_raw(total_pcs)}
    raw.update({r: list(v) for r, v in enumerate(filas, start=2)})
    pa = {1: [None, "Date/Time", "Activo 1\n Inactivo 0", "Excused Event"]}
    pa.update({r: list(v) for r, v in (actividad or {}).items()})
    hojas = {"RawData-PCS": raw, "PlantActivity": pa}
    if exclusion is not None:
        em = {1: encabezado_exclusion(total_pcs)}
        em.update({r: list(v) for r, v in exclusion.items()})
        hojas["Exclusion_Matrix"] = em
    return escribir_xlsx(ruta, hojas)


def encabezado_exclusion(total_pcs: int) -> list[str]:
    return ["Date/time", *[f"PCS{k:02d}" for k in range(1, total_pcs + 1)], "Excused Event", "Comments"]


def fila_raw(serial: float | None, modulos: Sequence[object], fallas: Sequence[object] | None = None) -> list[object]:
    """Fila completa de RawData-PCS con estado/advertencia fijos."""
    valores: list[object] = [serial]
    for k, m in enumerate(modulos):
        f = "NO FAULTS" if fallas is None else fallas[k]
        valores += [f, "ON", "NO WARNINGS", m]
    return valores


def matriz(
    modulos: Sequence[Sequence[float | None]],
    fallas: Sequence[Sequence[object]] | None = None,
    seriales: Sequence[float] | None = None,
    primera_fila: int = 2,
    siguiente: Sequence[float | None] | None = None,
) -> MatrizPCS:
    """Matriz en memoria: ``modulos[i][j]`` (``None`` = vacío), seriales cada 15 min desde el 01-sep."""
    mod = np.array([[np.nan if v is None else float(v) for v in fila] for fila in modulos], dtype=float)
    n, p = mod.shape
    ser = np.array(seriales if seriales is not None else [serial_min(15 * (i + 1)) for i in range(n)])
    fal = np.empty((n, p), dtype=object)
    for i in range(n):
        for j in range(p):
            fal[i, j] = "NO FAULTS" if fallas is None else fallas[i][j]
    sig = None if siguiente is None else np.array([np.nan if v is None else float(v) for v in siguiente])
    return MatrizPCS(
        numero_fila=np.arange(primera_fila, primera_fila + n),
        serial=ser,
        marca_tiempo=np.array([np.datetime64(serial_a_datetime(s), "s") for s in ser]),
        modulos=mod,
        modulos_nulo=np.isnan(mod),
        falla=fal,
        estado=np.full((n, p), "ON", dtype=object),
        advertencia=np.full((n, p), "NO WARNINGS", dtype=object),
        pcs=list(range(1, p + 1)),
        encabezado_falla=[f"Arena - PCS {k:02d} - POWERELECTRONICS GEN3 HEx CURRENT FAULT" for k in range(1, p + 1)],
        encabezado_modulos=[f"Arena - PCS {k:02d} - POWERELECTRONICS HEM-k NUMBER OF MODULES" for k in range(1, p + 1)],
        siguiente_modulos=sig,
        siguiente_nulo=None if sig is None else np.isnan(sig),
    )


def actividad(n: int, operacional: Sequence[float] | float = 1.0, primera_fila: int = 2) -> DatosActividad:
    """PlantActivity en memoria: solo la columna C pondera (F-37); D (espejo de BO) = 1."""
    op = np.broadcast_to(np.asarray(operacional, dtype=float), (n,)).copy()
    return DatosActividad(
        np.arange(primera_fila, primera_fila + n), np.full(n, np.nan), op, np.ones(n), np.full(n, np.nan)
    )


def exclusion(m: MatrizPCS, valores: Sequence[Sequence[float | None]], baterias_por_pcs: int = 4) -> DatosExclusion:
    """Exclusion_Matrix en memoria (``valores[i][j]``, ``None`` = vacío) pasando por el
    enriquecimiento real, que calcula las baterías previas de los EE con valor 2."""
    from etl_arena.enrichment import asociar_exclusion

    filas = {1: dict(enumerate(encabezado_exclusion(m.p), start=1))}
    for i, fila in enumerate(valores):
        celdas: dict[int, object] = {1: float(m.serial[i])}
        celdas.update({j + 2: float(v) for j, v in enumerate(fila) if v is not None})
        filas[int(m.numero_fila[i])] = celdas
    datos, _ = asociar_exclusion(m, filas, config_prueba(total_pcs=m.p, baterias_por_pcs=baterias_por_pcs))
    return datos

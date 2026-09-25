"""Carga mensual de la ``Exclusion_Matrix`` en la copia de trabajo (tarea 4.12; R3.8, R19.6, D-13, D-17, F-37).

1. ``leer_entrega``: la entrega de Alex. Formato asumido hasta tener la muestra (0.8): un ``.xlsx``/
   ``.xlsm`` con una hoja ``Exclusion_Matrix`` (o una sola hoja) con la estructura del libro de
   agosto: ``Date/time`` (serial Excel), ``PCS01…PCSnn``, ``Excused Event``, ``Comments``. Se toman
   solo las filas del mes cargado. Si la muestra real difiere, solo cambia esta función.
2. ``planificar``: ubica cada timestamp en la fila de ``RawData-PCS`` con el mismo serial (el Excel
   une por fila, F-01), **rechaza** timestamps inexistentes o valores fuera de 0/1/2, y calcula el
   diff celda a celda contra la hoja actual (vacío ≡ 0 en las columnas PCS).
3. ``aplicar_plan``: escribe por COM las filas del mes (crea la hoja con su encabezado si el libro no
   la trae) y guarda. Solo sobre copias de trabajo (``.bak``). El libro base no se toca.

``PlantActivity`` no se carga: las exclusiones salen solo de la matriz hasta que Alex confirme otra
cosa (el spec lo dejaba opcional).
"""

from __future__ import annotations

import csv
import hashlib
import logging
from calendar import monthrange
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from etl_arena.excel_semantics import datetime_a_serial, serial_a_datetime
from etl_arena.ingestion import LibroXlsx
from etl_arena.model import VALORES_EXCLUSION

log = logging.getLogger("etl_arena.exclusion")

HOJA = "Exclusion_Matrix"
HOJA_RAW = "RawData-PCS"


class ErrorCargaExclusion(ValueError):
    """La entrega no se puede cargar tal cual (mensaje accionable para Alex/Francisco)."""


def encabezado(total_pcs: int) -> list[str]:
    return ["Date/time", *(f"PCS{k:02d}" for k in range(1, total_pcs + 1)), "Excused Event", "Comments"]


def _clave(serial: float) -> int:
    return round(serial * 86400)  # segundos: tolera el último bit del serial


def _col(n: int) -> str:
    letras = ""
    while n:
        n, r = divmod(n - 1, 26)
        letras = chr(65 + r) + letras
    return letras


@dataclass(frozen=True)
class FilaEntrega:
    serial: float
    valores: tuple[float | None, ...]  # por PCS: 0/1/2 o vacío
    excusado: object
    comentario: str | None


@dataclass
class EntregaExclusion:
    archivo: Path
    sha256: str
    anio: int
    mes: int
    filas: list[FilaEntrega]
    filas_fuera_del_mes: int = 0


@dataclass(frozen=True)
class CambioCelda:
    fila: int
    serial: float
    numero_pcs: int | None  # None en Excused Event / Comments
    campo: str
    anterior: object
    nuevo: object


@dataclass
class PlanCarga:
    entrega: EntregaExclusion
    hoja_existe: bool
    filas: dict[int, list[object]]  # fila Excel → A..(p+3) a escribir
    cambios: list[CambioCelda] = field(default_factory=list)


def _sha256(ruta: Path) -> str:
    h = hashlib.sha256()
    with open(ruta, "rb") as fh:
        for bloque in iter(lambda: fh.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def _valor_pcs(v: object, donde: str) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v) in VALORES_EXCLUSION:
        return float(v)
    raise ErrorCargaExclusion(f"{donde}: valor {v!r} (se admite 0, 1, 2 o vacío)")


def leer_entrega(ruta: str | Path, anio: int, mes: int, total_pcs: int) -> EntregaExclusion:
    ruta = Path(ruta)
    esperado = encabezado(total_pcs)
    desde = datetime_a_serial(date(anio, mes, 1))
    hasta = datetime_a_serial(date(anio, mes, monthrange(anio, mes)[1])) + 1
    filas: list[FilaEntrega] = []
    fuera = 0
    with LibroXlsx(ruta) as lx:
        hojas = lx.hojas
        hoja = HOJA if HOJA in hojas else (hojas[0] if len(hojas) == 1 else None)
        if hoja is None:
            raise ErrorCargaExclusion(f"{ruta.name}: no tiene la hoja {HOJA!r} (hojas: {', '.join(hojas)})")
        for n_fila, celdas in lx.filas(hoja, max_col=len(esperado)):
            if n_fila == 1:
                for col, texto in enumerate(esperado, start=1):
                    if celdas.get(col) != texto:
                        raise ErrorCargaExclusion(
                            f"{ruta.name}!{hoja}: encabezado de la columna {_col(col)} = {celdas.get(col)!r}, "
                            f"se esperaba {texto!r}"
                        )
                continue
            ts = celdas.get(1)
            if ts is None:
                if any(celdas.get(c) not in (None, "") for c in range(2, len(esperado) + 1)):
                    raise ErrorCargaExclusion(f"{ruta.name}!{hoja} fila {n_fila}: valores sin Date/time")
                continue
            if not isinstance(ts, float):
                raise ErrorCargaExclusion(
                    f"{ruta.name}!{hoja} fila {n_fila}: Date/time {ts!r} no es una fecha de Excel (serial); "
                    "no se aceptan fechas como texto (F-21)"
                )
            if not desde <= ts < hasta:
                fuera += 1
                continue
            valores = tuple(
                _valor_pcs(celdas.get(j + 2), f"{ruta.name}!{hoja} fila {n_fila}, PCS{j + 1:02d}")
                for j in range(total_pcs)
            )
            comentario = celdas.get(total_pcs + 3)
            filas.append(
                FilaEntrega(ts, valores, celdas.get(total_pcs + 2), None if comentario is None else str(comentario))
            )
    if not filas:
        raise ErrorCargaExclusion(f"{ruta.name}: no hay filas de {anio}-{mes:02d}")
    return EntregaExclusion(ruta, _sha256(ruta), anio, mes, filas, fuera)


def _mismo(a: object, b: object, es_pcs: bool) -> bool:
    if es_pcs:
        return float(a or 0.0) == float(b or 0.0)  # vacío ≡ 0 (sin evento de exclusión)
    return (a in (None, "")) and (b in (None, "")) or a == b


def planificar(libro: str | Path, entrega: EntregaExclusion, total_pcs: int) -> PlanCarga:
    """Filas destino y diff contra la hoja actual. Falla si algún timestamp no existe en RawData-PCS."""
    libro = Path(libro)
    ancho = total_pcs + 3
    with LibroXlsx(libro) as lx:
        fila_de: dict[int, int] = {}
        for n_fila, c in lx.filas(HOJA_RAW, min_fila=2, max_col=1):
            a = c.get(1)
            if not isinstance(a, float):
                break  # el VBA corta en la primera A vacía
            fila_de.setdefault(_clave(a), n_fila)
        hoja_existe = HOJA in lx.hojas
        actual = dict(lx.filas(HOJA, max_col=ancho)) if hoja_existe else {}

    inexistentes = [f.serial for f in entrega.filas if _clave(f.serial) not in fila_de]
    if inexistentes:
        ejemplos = ", ".join(f"{serial_a_datetime(s):%Y-%m-%d %H:%M}" for s in inexistentes[:5])
        raise ErrorCargaExclusion(
            f"{len(inexistentes)} timestamps de la entrega no existen en RawData-PCS de {libro.name} "
            f"(p. ej. {ejemplos}): la matriz se une por fila y no se puede ubicar (R3.8)"
        )

    plan = PlanCarga(entrega, hoja_existe, {})
    for f in entrega.filas:
        fila = fila_de[_clave(f.serial)]
        if fila in plan.filas:
            raise ErrorCargaExclusion(f"timestamp repetido en la entrega: {serial_a_datetime(f.serial)}")
        nueva = [f.serial, *f.valores, f.excusado, f.comentario]
        plan.filas[fila] = nueva
        previa = actual.get(fila, {})
        for col in range(2, ancho + 1):
            es_pcs = col <= total_pcs + 1
            antes, despues = previa.get(col), nueva[col - 1]
            if not _mismo(antes, despues, es_pcs):
                campo = encabezado(total_pcs)[col - 1]
                plan.cambios.append(CambioCelda(fila, f.serial, col - 1 if es_pcs else None, campo, antes, despues))
    return plan


def escribir_cambios_csv(plan: PlanCarga, ruta: str | Path) -> Path:
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with open(ruta, "w", newline="", encoding="utf-8-sig") as fh:  # BOM: Excel lo abre con tildes
        w = csv.writer(fh, delimiter=";")
        w.writerow(["Hoja", "Fila", "MarcaTiempo", "NumeroPCS", "Campo", "ValorAnterior", "ValorNuevo", "Archivo"])
        for c in plan.cambios:
            w.writerow(
                [
                    HOJA,
                    c.fila,
                    f"{serial_a_datetime(c.serial):%Y-%m-%d %H:%M:%S}",
                    c.numero_pcs or "",
                    c.campo,
                    "" if c.anterior is None else c.anterior,
                    "" if c.nuevo is None else c.nuevo,
                    plan.entrega.archivo.name,
                ]
            )
    return ruta


def _bloques(filas: list[int]) -> list[tuple[int, int]]:
    bloques, inicio = [], None
    for k, f in enumerate(filas):
        if inicio is None:
            inicio = f
        if k == len(filas) - 1 or filas[k + 1] != f + 1:
            bloques.append((inicio, f))
            inicio = None
    return bloques


def aplicar_plan(libro: str | Path, plan: PlanCarga, total_pcs: int, timeout_s: float = 900) -> None:
    """Escribe las filas del plan en la hoja ``Exclusion_Matrix`` de la copia de trabajo, vía COM."""
    from etl_arena.workbook.com import SesionExcel

    libro = Path(libro)
    if not libro.with_suffix(libro.suffix + ".bak").exists():
        raise ValueError(f"{libro} no es una copia de trabajo (falta {libro.name}.bak): usar copiar_libro_trabajo")
    with SesionExcel(timeout_s=timeout_s) as sesion:
        _escribir(sesion, libro, plan, total_pcs)  # el proxy del libro muere antes de cerrar Excel


def _escribir(sesion, libro: Path, plan: PlanCarga, total_pcs: int) -> None:
    wb = sesion.abrir(libro)
    ultima_col = _col(total_pcs + 3)
    if plan.hoja_existe:
        hoja = wb.Worksheets(HOJA)
    else:
        hoja = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        hoja.Name = HOJA
        hoja.Range(f"A1:{ultima_col}1").Value2 = (tuple(encabezado(total_pcs)),)
        log.info("hoja %s creada en %s", HOJA, libro.name)
    filas = sorted(plan.filas)
    for f0, f1 in _bloques(filas):
        hoja.Range(f"A{f0}:{ultima_col}{f1}").Value2 = tuple(tuple(plan.filas[f]) for f in range(f0, f1 + 1))
    hoja.Range(f"A2:A{max(filas)}").NumberFormat = "dd-mm-yyyy hh:mm:ss"  # solo visual (F-22)
    wb.Save()
    log.info("Exclusion_Matrix: %d filas escritas, %d celdas cambiadas", len(filas), len(plan.cambios))


def exclusiones_en_periodo(libro: str | Path, desde: date, hasta: date, total_pcs: int) -> bool:
    """¿La hoja marca algún 1/2 en filas cuyo ``Date/time`` cae en [desde, hasta + 1)?

    Sin la hoja: ``False``. Filas marcadas sin ``Date/time`` cuentan como dentro (conservador).
    """
    a, b = datetime_a_serial(desde), datetime_a_serial(hasta) + 1
    with LibroXlsx(libro) as lx:
        if HOJA not in lx.hojas:
            return False
        for _fila, c in lx.filas(HOJA, min_fila=2, max_col=total_pcs + 1):
            if not any(c.get(j) not in (None, 0, 0.0, "") for j in range(2, total_pcs + 2)):
                continue
            ts = c.get(1)
            if not isinstance(ts, float) or a <= ts < b:
                return True
    return False

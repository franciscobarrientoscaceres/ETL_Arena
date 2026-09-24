"""Lectura de ``RawData-PCS``, ``PlantActivity`` y ``Exclusion_Matrix`` del libro de trabajo (R4, F-37)."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from etl_arena.config import CELDAS_PARAMETROS, ConfiguracionCalculo
from etl_arena.excel_semantics import es_vacio
from etl_arena.ingestion.anomalias_timestamp import detectar_anomalias_timestamp
from etl_arena.ingestion.xlsx_stream import LibroXlsx
from etl_arena.model import Anomalia, ErrorParidad

HOJA_RAW = "RawData-PCS"
HOJA_ACTIVIDAD = "PlantActivity"
HOJA_EXCLUSION = "Exclusion_Matrix"
COLUMNAS_ACTIVIDAD = 9  # A..I (A vacía, B ts, C op, D exc, E..I informativas)


@dataclass
class FilaCruda:
    numero_fila: int
    serial: float  # A; 0.0 si A2 está vacía (Empty = 0 en VBA)
    celdas: dict[int, object]  # {col: valor} de la fila completa


@dataclass
class LibroCrudo:
    ruta: Path
    encabezado: dict[int, object]  # fila 1 de RawData-PCS
    filas: list[FilaCruda]  # filas de datos en orden de hoja
    fila_siguiente: dict[int, object]  # fila posterior a la última leída (la que corta el VBA)
    actividad: dict[int, dict[int, object]]  # PlantActivity: fila → {col: valor}
    filas_descartadas: int = 0
    exclusion: dict[int, dict[int, object]] | None = None  # Exclusion_Matrix: fila → {col: valor}; None si no existe
    anomalias: list[Anomalia] = field(default_factory=list)


def _serial_a(v: object, fila: int) -> float:
    if isinstance(v, bool) or not isinstance(v, float):
        raise ErrorParidad(
            f"RawData-PCS!A{fila} no es fecha/serial: {v!r} (F-21)",
            [Anomalia("a_no_numerico", "error", numero_fila=fila, detalle=repr(v))],
        )
    return v


def leer_libro(ruta: str | Path, cfg: ConfiguracionCalculo) -> LibroCrudo:
    """Lee el libro según ``cfg.modo_huecos`` (R4.5).

    ``excel``: como el VBA, procesa la fila 2 aunque su A esté vacía y termina en la primera
    A vacía desde la fila 3; las filas posteriores se descartan y se reportan.
    ``continuar``: omite solo las filas con A vacía.
    """
    ruta = Path(ruta)
    anomalias: list[Anomalia] = []
    filas: list[FilaCruda] = []
    fila_siguiente: dict[int, object] = {}
    descartadas = 0
    encabezado: dict[int, object] = {}
    cortado = False

    with LibroXlsx(ruta) as libro:
        esperada = 2
        for numero, celdas in libro.filas(HOJA_RAW):
            if numero == 1:
                encabezado = celdas
                continue
            if cortado:
                descartadas += 1
                continue
            if numero > esperada:  # filas sin ninguna celda entre medio
                if esperada == 2:
                    anomalias.append(
                        Anomalia(
                            "celda_a_vacia",
                            "advertencia",
                            numero_fila=2,
                            detalle="fila 2 vacía: la macro la procesa como 0 y sigue",
                        )
                    )
                    filas.append(FilaCruda(2, 0.0, {}))
                    esperada = 3
                if numero > esperada and cfg.modo_huecos == "excel":
                    cortado = True
                    anomalias.append(
                        Anomalia(
                            "celda_a_vacia",
                            "advertencia",
                            numero_fila=esperada,
                            detalle="primera A vacía: el VBA termina aquí",
                        )
                    )
                    descartadas += 1
                    continue
            esperada = numero + 1
            a = celdas.get(1)
            if es_vacio(a):
                if numero == 2:
                    anomalias.append(
                        Anomalia(
                            "celda_a_vacia",
                            "advertencia",
                            numero_fila=2,
                            detalle="A2 vacía: la macro la procesa como 0 y sigue",
                        )
                    )
                    filas.append(FilaCruda(2, 0.0, celdas))
                    continue
                anomalias.append(
                    Anomalia(
                        "celda_a_vacia",
                        "advertencia",
                        numero_fila=numero,
                        detalle="A vacía con datos en otras columnas",
                    )
                )
                if cfg.modo_huecos == "excel":
                    cortado = True
                    fila_siguiente = celdas
                continue
            filas.append(FilaCruda(numero, _serial_a(a, numero), celdas))

        if descartadas:
            anomalias.append(
                Anomalia(
                    "filas_truncadas",
                    "advertencia",
                    detalle=f"{descartadas} filas con datos después de la primera A vacía (D-01)",
                )
            )
        ultima = filas[-1].numero_fila if filas else 1
        actividad: dict[int, dict[int, object]] = {}
        for numero, celdas in libro.filas(HOJA_ACTIVIDAD, min_fila=2, max_fila=ultima, max_col=COLUMNAS_ACTIVIDAD):
            actividad[numero] = celdas
        exclusion: dict[int, dict[int, object]] | None = None
        if HOJA_EXCLUSION in libro.hojas:
            # A (ts) + una columna por PCS + "Excused Event" + "Comments"
            exclusion = dict(libro.filas(HOJA_EXCLUSION, max_fila=ultima, max_col=cfg.total_pcs + 3))
        else:
            anomalias.append(
                Anomalia(
                    "exclusion_matrix_ausente",
                    "info",
                    detalle="el libro no trae Exclusion_Matrix: sin eventos de exclusión (F-37)",
                )
            )

    con_serial = [f for f in filas if not (f.numero_fila == 2 and f.serial == 0.0)]
    anomalias.extend(
        detectar_anomalias_timestamp(
            [f.serial for f in con_serial], [f.numero_fila for f in con_serial], cfg.minutos_muestreo, cfg.zona_horaria
        )
    )
    return LibroCrudo(
        ruta,
        encabezado,
        filas,
        fila_siguiente,
        actividad,
        filas_descartadas=descartadas,
        exclusion=exclusion,
        anomalias=anomalias,
    )


def leer_celdas_parametros(ruta: str | Path) -> dict[str, object]:
    """Valores crudos de las celdas de parámetros para ``config.parametros_desde_celdas``."""
    with LibroXlsx(ruta) as libro:
        return libro.celdas(CELDAS_PARAMETROS)

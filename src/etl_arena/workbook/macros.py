"""Ejecución de las macros del libro de trabajo vía COM (tarea 4.2; R3.5, F-05, F-21, F-23, F-33).

Escribe los parámetros de la configuración en el libro y ejecuta en orden
``cmdCalcAvailability → mcoCreateList → mcoDailyAvailability → Graphupdate``; luego guarda.
La referencia se lee después por XML (``workbook.referencia``). Solo sobre copias de trabajo.
"""

from __future__ import annotations

import logging
from pathlib import Path

from etl_arena.config import ConfiguracionCalculo
from etl_arena.config.models import MAX_DIAS_DIARIO
from etl_arena.excel_semantics import datetime_a_serial
from etl_arena.workbook.com import ErrorVBA, SesionExcel
from etl_arena.workbook.vba import macros_aplican_matriz

log = logging.getLogger("etl_arena.excel")

MACROS = ("cmdCalcAvailability", "mcoCreateList", "mcoDailyAvailability", "Graphupdate")
# Excel 2016 MSI (build < 10000) no tiene Sort.SortFields.Add2 y esas macros fallan con el error 438:
# * mcoCreateList: mcoOrder es su última línea; la lista de eventos ya está completa y solo queda sin
#   ordenar el resumen N:Q, que la reconciliación compara como mapa.
# * Graphupdate: capa de presentación (hoja Graph), fuera del motor y de la reconciliación.
# En ambos casos se tolera con advertencia; en Excel moderno el 438 es un error real.
BUILD_MINIMO_ADD2 = 10000
EFECTO_SIN_ADD2 = {
    "mcoCreateList": "mcoOrder no ordenó ListOfFaults!N:Q (sin efecto en la reconciliación)",
    "Graphupdate": "la hoja Graph quedó sin actualizar (presentación; sin efecto en la reconciliación)",
}
FILA_DIA_1 = 9  # Daily!C9 = D3 (= C5)


def _si_no(valor: bool) -> str:
    return "Yes" if valor else "No"


def escribir_parametros(wb, cfg: ConfiguracionCalculo, excel_aplica_matriz: bool = False) -> None:
    """C5/C7/C21/C31, L2/L4/L14, Daily!D5 y las filas Daily!B/C/E/F/G hasta ``fin_diario``.

    Fechas como serial (``Value2``), nunca texto (F-21). C21 se escribe "No" / "Yes" (el VBA
    aplica el factor operacional si C21 <> "No"); C31/L14 "Yes" / "No" (F-12).

    C31/L14 = "Yes" solo si las macros del libro aplican la ``Exclusion_Matrix`` (maestro v1.1,
    ``workbook.vba``): las de septiembre excusarían con ``PlantActivity!D``, que no se usa para
    exclusiones (F-44). Con "No" la referencia vale para períodos sin exclusiones en la matriz.
    """
    calc = wb.Worksheets("Calculation-Availability")
    calc.Range("C5").Value2 = datetime_a_serial(cfg.inicio_periodo)
    calc.Range("C7").Value2 = datetime_a_serial(cfg.fin_periodo)
    calc.Range("C21").Value2 = _si_no(cfg.solo_tiempo_operacional)
    calc.Range("C31").Value2 = _si_no(cfg.aplicar_evento_excusable and excel_aplica_matriz)

    lof = wb.Worksheets("ListOfFaults")
    lof.Range("L2").Value2 = datetime_a_serial(cfg.inicio_periodo_eventos)
    lof.Range("L4").Value2 = datetime_a_serial(cfg.fin_periodo_eventos)
    lof.Range("L14").Value2 = _si_no(cfg.aplicar_evento_excusable_eventos and excel_aplica_matriz)

    # Daily: la macro corta en la primera C vacía (F-33) y solo escribe D; B (n° de día), C, E, F
    # y G son de la plantilla, que se recorta a mano cada mes (septiembre trae solo 21 filas). Se
    # escriben las n filas pedidas con las fórmulas de la plantilla y se limpian las demás hasta la
    # fila 39: una F sobrante con E vacía mostraría disponibilidad 1.
    daily = wb.Worksheets("Daily")
    daily.Range("D5").Value2 = datetime_a_serial(cfg.fin_diario)
    n = cfg.dias_diario
    ultima = FILA_DIA_1 + n - 1
    for fila in range(FILA_DIA_1, ultima + 1):
        primera = fila == FILA_DIA_1
        daily.Range(f"B{fila}").Value2 = fila - FILA_DIA_1 + 1
        daily.Range(f"C{fila}").Formula = "=D3" if primera else f"=C{fila - 1}+1"
        daily.Range(f"E{fila}").Formula = f"=D{fila}" if primera else f"=D{fila}+E{fila - 1}"
        daily.Range(
            f"F{fila}"
        ).Formula = f'=IFERROR(1-E{fila}/(Total_Racks*24*60*B{fila}/Frecuencia_de_muestreo__min),"")'
        if primera:
            daily.Range(f"G{fila}").Value2 = 0
        else:
            daily.Range(f"G{fila}").Formula = f'=IFERROR(F{fila}-F{fila - 1},"")'
    fin = FILA_DIA_1 + MAX_DIAS_DIARIO - 1
    if ultima < fin:
        for col in "CEFG":
            daily.Range(f"{col}{ultima + 1}:{col}{fin}").ClearContents()


def ejecutar_macros(
    libro: str | Path,
    cfg: ConfiguracionCalculo,
    timeout_s: float = 900,
    visible: bool = True,
    excel_aplica_matriz: bool | None = None,
) -> dict[str, float]:
    """Parámetros + 4 macros + guardar. Devuelve los segundos por macro.

    Solo sobre copias de trabajo de ``copiar_libro_trabajo`` (con su ``.bak`` al lado): las macros
    escriben y guardan el libro, así que nunca se corren sobre el maestro ni sobre ``data/processed``.
    """
    libro = Path(libro)
    if not libro.with_suffix(libro.suffix + ".bak").exists():
        raise ValueError(f"{libro} no es una copia de trabajo (falta {libro.name}.bak): usar copiar_libro_trabajo")
    if excel_aplica_matriz is None:
        excel_aplica_matriz = macros_aplican_matriz(libro)
    if not excel_aplica_matriz and (cfg.aplicar_evento_excusable or cfg.aplicar_evento_excusable_eventos):
        log.warning(
            "las macros de %s no aplican Exclusion_Matrix: C31/L14 se escriben 'No' (no excusar con PlantActivity)",
            libro.name,
            extra={"id_corrida": cfg.id_corrida},
        )
    with SesionExcel(visible=visible, timeout_s=timeout_s) as sesion:
        return _correr(sesion, libro, cfg, excel_aplica_matriz)  # el proxy del libro muere antes de cerrar Excel


def _correr(
    sesion: SesionExcel, libro: str | Path, cfg: ConfiguracionCalculo, excel_aplica_matriz: bool
) -> dict[str, float]:
    tiempos: dict[str, float] = {}
    wb = sesion.abrir(libro)
    escribir_parametros(wb, cfg, excel_aplica_matriz)
    for macro in MACROS:
        try:
            tiempos[macro] = sesion.ejecutar_macro(wb, macro)
        except ErrorVBA as exc:
            if not (macro in EFECTO_SIN_ADD2 and "438" in exc.mensaje and sesion.build < BUILD_MINIMO_ADD2):
                raise
            tiempos[macro] = exc.segundos
            log.warning(
                "Excel %s sin SortFields.Add2: %s",
                sesion.version,
                EFECTO_SIN_ADD2[macro],
                extra={"id_corrida": cfg.id_corrida},
            )
        log.info("macro %s: %.1f s", macro, tiempos[macro], extra={"id_corrida": cfg.id_corrida})
    wb.Save()
    return tiempos

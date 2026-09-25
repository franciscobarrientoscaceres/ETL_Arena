"""Niveles 1–5 de reconciliación (R13.1–R13.6; design.md §reconciliation).

| Nivel | Compara |
|---|---|
| 1 Input | filas procesadas, primer/último serial, C23, parámetros efectivos |
| 2 Muestra | ``ponderadas[k, j]`` vs ``Calc!F:BN`` (vacío = 0) y ``Calc!BO`` |
| 3 Acumulados | C12 (exacto), C14 |
| 4 KPI | C16, C19 y Daily por día (D, E, F, G) |
| 5 Eventos | ``ListOfFaults!B:I`` en orden de escritura, conteo, L10, resumen N:Q (como mapa) |
"""

from __future__ import annotations

import numpy as np

from etl_arena.excel_semantics import datetime_a_serial, texto_excel
from etl_arena.model import ReferenciaExcel
from etl_arena.pipeline import ResultadoCalculo
from etl_arena.reconciliation.modelo import ResultadoNivel
from etl_arena.reconciliation.tolerancias import Tolerancias


def _hay_exclusiones(r: ResultadoCalculo, desde, hasta) -> bool:
    """¿La ``Exclusion_Matrix`` marca algún 1/2 en las filas del rango [desde, hasta + 1)?"""
    s = r.matriz.serial
    filas = (s >= datetime_a_serial(desde)) & (s < datetime_a_serial(hasta) + 1)
    return bool(np.any(r.exclusion.valor[filas] != 0))


def nivel_1(r: ResultadoCalculo, ref: ReferenciaExcel, tol: Tolerancias) -> ResultadoNivel:
    n = ResultadoNivel(1, "input")
    cfg, d = r.cfg, r.disponibilidad
    n.comparar("filas_procesadas", "C12 vs filas de Calc!E", d.bloques_muestreo, len(ref.seriales_tabla), 0)
    if d.bloques_muestreo and ref.seriales_tabla:
        n.comparar("primer_serial", "Calc!E4", float(d.serial[0]), ref.seriales_tabla[0], 0)
        n.comparar("ultimo_serial", "Calc!E ultima", float(d.serial[-1]), ref.seriales_tabla[-1], 0)
    n.comparar("C23", "Calc!C23", d.minutos_muestreo_derivado, ref.c23, 0)
    n.comparar("parametro", "C5", cfg.inicio_periodo, ref.c5, 0)
    n.comparar("parametro", "C7", cfg.fin_periodo, ref.c7, 0)
    n.comparar("parametro", "C21 (<> No)", cfg.solo_tiempo_operacional, ref.c21 != "No", 0)
    # Sin exclusiones en la matriz para el período, "Yes" ≡ "No": con las macros de septiembre
    # se escribe "No" para no excusar con PlantActivity!D (F-44).
    if _hay_exclusiones(r, cfg.inicio_periodo, cfg.fin_periodo):
        n.comparar("parametro", "C31 (= Yes)", cfg.aplicar_evento_excusable, ref.c31 == "Yes", 0)
    else:
        n.comparar("parametro", "C31 (sin exclusiones en el período: Yes ≡ No)", True, ref.c31 in ("Yes", "No"), 0)
    n.comparar("parametro", "L2", cfg.inicio_periodo_eventos, ref.l2, 0)
    n.comparar("parametro", "L4", cfg.fin_periodo_eventos, ref.l4, 0)
    if _hay_exclusiones(r, cfg.inicio_periodo_eventos, cfg.fin_periodo_eventos):
        n.comparar("parametro", "L14 (= Yes)", cfg.aplicar_evento_excusable_eventos, ref.l14 == "Yes", 0)
    else:
        n.comparar("parametro", "L14 (sin exclusiones en el período: Yes ≡ No)", True, ref.l14 in ("Yes", "No"), 0)
    if ref.diario:
        n.comparar("parametro", "fin_diario vs ultima Daily!C (F-33)", cfg.fin_diario, ref.diario[-1].dia, 0)
    return n


def nivel_2(r: ResultadoCalculo, ref: ReferenciaExcel, tol: Tolerancias) -> ResultadoNivel:
    n = ResultadoNivel(2, "muestra")
    d = r.disponibilidad
    excel = {(fila - 4, pcs): v for fila, _, pcs, v in ref.tabla}
    pond = np.nan_to_num(d.ponderadas, nan=0.0)
    filas = min(d.bloques_muestreo, len(ref.seriales_tabla))
    for k in range(filas):
        for j, pcs in enumerate(d.pcs):
            n.comparar(
                "ponderada", f"fila {k + 4} PCS {pcs}", float(pond[k, j]), excel.get((k, pcs), 0.0), tol.ponderada_celda
            )
        if k < len(ref.bo):
            n.comparar("BO", f"fila {k + 4}", float(d.bo_evento_excusado_pa[k]), ref.bo[k] or 0.0, 0)
    return n


def nivel_3(r: ResultadoCalculo, ref: ReferenciaExcel, tol: Tolerancias) -> ResultadoNivel:
    n = ResultadoNivel(3, "acumulados")
    d = r.disponibilidad
    n.comparar("C12", "Calc!C12", d.bloques_muestreo, ref.c12, 0)
    n.comparar("C14", "Calc!C14", d.bloques_racks_indisponibles, ref.c14, tol.c14)
    return n


def nivel_4(r: ResultadoCalculo, ref: ReferenciaExcel, tol: Tolerancias) -> ResultadoNivel:
    n = ResultadoNivel(4, "kpi")
    d = r.disponibilidad
    n.comparar("C16", "Calc!C16", d.disponibilidad_periodo, ref.c16, tol.kpi_disponibilidad)
    n.comparar("C19", "Calc!C19", d.disponibilidad_anual_acumulada, ref.c19, tol.kpi_disponibilidad)
    n.comparar("dias", "Daily filas", len(r.diario), len(ref.diario), 0)
    for dia, x in zip(r.diario, ref.diario, strict=False):
        clave = f"Daily dia {dia.numero_dia}"
        n.comparar("fecha", clave, dia.dia, x.dia, 0)
        n.comparar("D", clave, dia.diario, x.d, tol.daily_bloques)
        n.comparar("E", clave, dia.acumulado, x.e, tol.daily_bloques)
        n.comparar("F", clave, dia.disponibilidad, x.f, tol.kpi_disponibilidad)
        n.comparar("G", clave, dia.variacion, x.g, tol.kpi_disponibilidad)
    return n


def nivel_5(r: ResultadoCalculo, ref: ReferenciaExcel, tol: Tolerancias) -> ResultadoNivel:
    n = ResultadoNivel(5, "eventos")
    ev = r.eventos
    registros = [*ev.cerrados, *([ev.incompleto] if ev.incompleto is not None else [])]
    n.comparar("conteo", "ListOfFaults filas", len(registros), len(ref.eventos), 0)
    for py, x in zip(registros, ref.eventos, strict=False):
        clave = f"ListOfFaults orden {py.orden_excel}"
        n.comparar("orden", clave, py.orden_excel, x.orden_excel, 0)
        n.comparar("B", clave, py.numero_pcs, x.b, 0)
        n.comparar("C", clave, py.serial_inicio, x.c, tol.evento_serial_dias)
        n.comparar("D", clave, py.serial_fin, x.d, tol.evento_serial_dias)
        n.comparar("E", clave, py.duracion_horas, x.e, tol.evento_duracion_promedio)
        n.comparar("F", clave, py.codigo_falla, x.f, 0)
        n.comparar("G", clave, texto_excel(py.descripcion_falla), x.g if x.g is not None else "", 0)
        n.comparar("H", clave, py.promedio_baterias, x.h, tol.evento_duracion_promedio)
        n.comparar("I", clave, py.horas_rack, x.i, tol.evento_horas_rack)
    n.comparar("L10", "ListOfFaults!L10", ev.horas_rack_totales, ref.l10, tol.l10)
    # N:Q se compara como mapa código → P: el orden de los empates depende del estado previo de la hoja
    python_p = {f.codigo: f.horas_rack for f in ev.resumen}
    excel_p = {c: p for c, p in ref.resumen_codigos}
    n.comparar("resumen_codigos", "conjunto de códigos N", sorted(python_p), sorted(excel_p), 0)
    for codigo in sorted(set(python_p) & set(excel_p)):
        n.comparar("P", f"resumen {codigo}", python_p[codigo], excel_p[codigo], tol.l10)
    return n

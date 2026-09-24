"""``MotorDisponibilidad``: emulación de ``cmdCalcAvailability`` (R7, F-07, F-12, F-13).

VBA de referencia (Module1)::

    If A(r) >= C5 And A(r) < (1 + C7) Then
        E(res) = A(r) : BO(res) = PlantActivity!D(r)
        For pcs = 1 To C2
            If M(r) <> "" And M(r) < 4 Then
                If C31 = "Yes" Then F(res) = (C3 - M(r)) * BO(res) Else F(res) = C3 - M(r)
                If C21 = "No" Then C14 = C14 + 12 * F(res) Else C14 = C14 + 12 * F(res) * PlantActivity!C(r)
        C12 = C12 + 1

Las sumas son secuenciales en orden fila → PCS, con la misma asociación de operaciones
que el VBA, para reproducir C14 bit a bit.
"""

from __future__ import annotations

import numpy as np

from etl_arena.availability.resultado import ResultadoDisponibilidad
from etl_arena.config import ConfiguracionCalculo
from etl_arena.excel_semantics import datetime_a_serial, redondear_excel
from etl_arena.model import DatosActividad, MatrizPCS

DIAS_ANIO_C19 = 365  # C19 = 1 - C14/(C11*(365*24*4)): año fijo de 365 días, literal de la hoja


def calcular(m: MatrizPCS, act: DatosActividad, cfg: ConfiguracionCalculo) -> ResultadoDisponibilidad:
    if m.p < cfg.total_pcs:
        raise ValueError(f"la matriz tiene {m.p} PCS y C2 = {cfg.total_pcs}")
    inicio = datetime_a_serial(cfg.inicio_periodo)
    fin_mas_uno = datetime_a_serial(cfg.fin_periodo) + 1
    c3 = float(cfg.baterias_por_pcs)
    racks = float(cfg.racks_por_pcs)
    p = cfg.total_pcs

    en_rango = np.flatnonzero((m.serial >= inicio) & (m.serial < fin_mas_uno))  # orden origen (R7.1)
    k = len(en_rango)
    baterias = np.full((k, p), np.nan)
    ponderadas = np.full((k, p), np.nan)
    impacto = np.zeros((k, p))
    fe = act.factor_excusable[en_rango].astype(float)
    fo = act.factor_operacional[en_rango].astype(float)
    fe_py, fo_py = fe.tolist(), fo.tolist()  # floats de Python: misma aritmética IEEE, sin np.float64
    modulos, nulos = m.modulos.tolist(), m.modulos_nulo.tolist()

    c14 = 0.0
    for fila_res, i in enumerate(en_rango.tolist()):
        excusable = fe_py[fila_res]
        operacional = fo_py[fila_res]
        fila_mod, fila_nula = modulos[i], nulos[i]
        for j in range(p):
            if fila_nula[j]:
                continue
            x = fila_mod[j]
            if not x < c3:
                continue
            bat = c3 - x
            pond = bat * excusable if cfg.aplicar_evento_excusable else bat
            aporte = racks * pond * operacional if cfg.solo_tiempo_operacional else racks * pond
            c14 = c14 + aporte  # orden fila → PCS (R7.7)
            baterias[fila_res, j] = bat
            ponderadas[fila_res, j] = pond
            impacto[fila_res, j] = aporte
    c12 = k  # una por fila, no por timestamp (R7.2, F-07)

    serial = m.serial[en_rango].copy()
    c23 = redondear_excel(float(serial[1] - serial[0]) * 24 * 60, 2) if k >= 2 else None  # F-13
    total_racks = cfg.total_racks
    c16 = 1 - c14 / (total_racks * c12) if c12 else None
    c19 = 1 - c14 / (total_racks * (DIAS_ANIO_C19 * 24 * 4))
    return ResultadoDisponibilidad(
        bloques_muestreo=c12,
        bloques_racks_indisponibles=c14,
        disponibilidad_periodo=c16,
        disponibilidad_anual_acumulada=c19,
        minutos_muestreo_derivado=c23,
        filas_procesadas=en_rango,
        serial=serial,
        numero_fila=m.numero_fila[en_rango].copy(),
        pcs=list(m.pcs[:p]),
        baterias=baterias,
        ponderadas=ponderadas,
        impacto=impacto,
        factor_excusable=fe,
        factor_operacional=fo,
    )

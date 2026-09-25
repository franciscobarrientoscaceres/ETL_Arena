"""Invariantes cruzados (R13.8, F-24): ``C14 = 4·L10`` y, por PCS, ``Σ racks·ponderadas / 4 = Σ I``.

Solo valen si ambos motores miran lo mismo: mismos períodos y flag excusable (C5/C7/C31 =
L2/L4/L14), sin factor operacional (C21 = "No") y sin eventos arrastrados (F-06). En otro caso
se reportan como ``no_aplica`` con el motivo.
"""

from __future__ import annotations

import numpy as np

from etl_arena.pipeline import ResultadoCalculo
from etl_arena.reconciliation.modelo import ResultadoNivel
from etl_arena.reconciliation.tolerancias import Tolerancias


def motivos_no_aplica(r: ResultadoCalculo) -> list[str]:
    cfg, ev = r.cfg, r.eventos
    motivos = []
    if (cfg.inicio_periodo, cfg.fin_periodo) != (cfg.inicio_periodo_eventos, cfg.fin_periodo_eventos):
        motivos.append("períodos KPI (C5/C7) ≠ eventos (L2/L4)")
    if cfg.aplicar_evento_excusable != cfg.aplicar_evento_excusable_eventos:
        motivos.append("C31 ≠ L14")
    if cfg.solo_tiempo_operacional:
        motivos.append("C21 ≠ No (el factor operacional no entra en los eventos)")
    if any(e.arrastrado_excel for e in ev.cerrados) or ev.incompleto is not None:
        motivos.append("eventos arrastrados o incompletos (F-06)")
    return motivos


def evaluar(r: ResultadoCalculo, tol: Tolerancias) -> ResultadoNivel:
    n = ResultadoNivel(9, "invariantes")
    motivos = motivos_no_aplica(r)
    if motivos:
        n.no_aplica.extend(motivos)
        return n
    d, ev, cfg = r.disponibilidad, r.eventos, r.cfg
    n.comparar("C14 = 4·L10", "global", d.bloques_racks_indisponibles, 4 * ev.horas_rack_totales, tol.c14)
    por_pcs_kpi = cfg.racks_por_pcs * np.nan_to_num(d.ponderadas, nan=0.0).sum(axis=0) / 4
    por_pcs_eventos: dict[int, float] = {}
    for e in ev.cerrados:
        por_pcs_eventos[e.numero_pcs] = por_pcs_eventos.get(e.numero_pcs, 0.0) + e.horas_rack
    for j, pcs in enumerate(d.pcs):
        n.comparar("Σ racks·pond/4 = Σ I", f"PCS {pcs}", float(por_pcs_kpi[j]), por_pcs_eventos.get(pcs, 0.0), tol.c14)
    return n

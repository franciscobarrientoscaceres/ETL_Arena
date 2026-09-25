"""Resumen de calidad de una corrida (R15.2, R15.3) → ``etl_run.ResumenCalidad``.

Las anomalías individuales van a ``data_quality_issue`` (``persistence.paquete``); aquí solo se
resume lo que un revisor necesita ver de un vistazo.
"""

from __future__ import annotations

from collections import Counter

import numpy as np

from etl_arena.model import Anomalia
from etl_arena.pipeline import ResultadoCalculo

TIPOS_TIMESTAMP = ("hueco", "duplicado", "fuera_de_orden", "frecuencia_distinta", "dst_salto", "dst_repeticion")
TIPOS_PLANT_ACTIVITY = ("pa_sin_timestamp", "pa_desalineado", "pa_vacio")
TIPOS_EXCLUSION = (
    "exclusion_matrix_ausente",
    "em_desalineado",
    "em_sin_timestamp",
    "ee2_sin_fila_previa",
    "ee2_difiere_macro_agosto",
)


def generar_resumen(r: ResultadoCalculo, anomalias_extra: list[Anomalia] | None = None) -> dict:
    anomalias = [*r.anomalias, *(anomalias_extra or [])]
    por_tipo = Counter(a.tipo for a in anomalias)
    por_severidad = Counter(a.severidad for a in anomalias)
    d, ev, m = r.disponibilidad, r.eventos, r.matriz
    en_periodo = d.filas_procesadas
    registros = [*ev.cerrados, *([ev.incompleto] if ev.incompleto is not None else [])]
    return {
        "anomalias": {"total": len(anomalias), "por_severidad": dict(por_severidad), "por_tipo": dict(por_tipo)},
        "modulos_nulos": {  # R15.4: se reportan, no se suprimen
            "celdas_en_periodo": int(m.modulos_nulo[en_periodo].sum()) if len(en_periodo) else 0,
            "celdas_en_libro": int(m.modulos_nulo.sum()),
        },
        "eventos": {
            "total": len(ev.cerrados),
            "descripcion_fallback": sum(1 for e in ev.cerrados if e.fallback),
            "arrastrados_excel": sum(1 for e in ev.cerrados if e.arrastrado_excel),
            "excel_habria_fallado": sum(1 for e in registros if e.excel_habria_fallado),
            "incompleto": ev.incompleto is not None,
            "descripcion_numerica": sum(1 for e in ev.cerrados if isinstance(e.descripcion_falla, float)),
            "con_exclusion": sum(1 for e in ev.cerrados if e.tiene_exclusion),
            "sin_catalogo": por_tipo.get("codigo_sin_catalogo", 0),
        },
        "timestamp": {t: por_tipo[t] for t in TIPOS_TIMESTAMP if por_tipo.get(t)},
        "filas_truncadas": r.libro.filas_descartadas,
        "plant_activity": {t: por_tipo[t] for t in TIPOS_PLANT_ACTIVITY if por_tipo.get(t)},
        "exclusion_matrix": {
            "presente": r.libro.exclusion is not None,
            "celdas_valor_1_en_periodo": int((r.exclusion.valor[en_periodo] == 1).sum()) if len(en_periodo) else 0,
            "celdas_valor_2_en_periodo": int((r.exclusion.valor[en_periodo] == 2).sum()) if len(en_periodo) else 0,
            **{t: por_tipo[t] for t in TIPOS_EXCLUSION if por_tipo.get(t)},
        },
        "esquema": {
            t: por_tipo[t]
            for t in ("columna_faltante", "columna_desplazada", "nombre_inesperado", "pcs_distinto_c2")
            if por_tipo.get(t)
        },
        "c23": {
            "derivado": d.minutos_muestreo_derivado,
            "configurado": r.cfg.minutos_muestreo,
            "distinto": bool(
                d.minutos_muestreo_derivado is not None and d.minutos_muestreo_derivado != r.cfg.minutos_muestreo
            ),
        },
        "kpi": {
            "C12": d.bloques_muestreo,
            "C14": d.bloques_racks_indisponibles,
            "C16": d.disponibilidad_periodo,
            "L10": ev.horas_rack_totales,
            "PCS_con_falla_en_periodo": int(np.any(~np.isnan(d.baterias), axis=0).sum()) if len(en_periodo) else 0,
        },
    }

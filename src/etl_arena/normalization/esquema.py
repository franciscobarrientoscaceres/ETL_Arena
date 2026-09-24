"""Validación del encabezado de ``RawData-PCS`` (R5.2, R5.5, F-17).

El VBA lee **por posición**: PCS k → FAULT en la columna ``4k-2`` y NUMBER OF MODULES en
``4k+1``. ``cmdCalcAvailability`` recorre ``C2`` PCS; ``mcoCreateList`` recorre bloques
mientras el encabezado FAULT de la fila 1 no esté vacío. Cualquier desvío hace que los
dos motores del Excel lean datos distintos, así que en modo paridad es un error.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from etl_arena.config import ConfiguracionCalculo
from etl_arena.excel_semantics import es_vacio
from etl_arena.model import Anomalia

PATRON = re.compile(r"^Arena - PCS (\d{2}) - POWERELECTRONICS (.+)$")
CAMPOS = {
    "GEN3 HEx CURRENT FAULT": "falla",
    "GEN3 HEx CURRENT STATUS": "estado",
    "GEN3 HEx CURRENT WARNING": "advertencia",
    "HEM-k NUMBER OF MODULES": "modulos",
}
ORDEN_CAMPOS = ("falla", "estado", "advertencia", "modulos")


def columna(pcs: int, campo: str) -> int:
    """Columna 1-based del campo del PCS según la posición que usa el VBA."""
    return 4 * pcs - 2 + ORDEN_CAMPOS.index(campo)


@dataclass(frozen=True)
class MapaColumnas:
    pcs: list[int]  # 1..p, en orden de encabezado
    pcs_encabezado_eventos: int  # bloques que recorre mcoCreateList (F-17)


def pcs_recorridos_por_eventos(encabezado: dict[int, object]) -> int:
    """``mcoCreateList``: procesa la columna 2 y sigue de 4 en 4 hasta un encabezado vacío."""
    k = 1
    while not es_vacio(encabezado.get(columna(k + 1, "falla"))):
        k += 1
    return k


def validar_esquema(encabezado: dict[int, object], cfg: ConfiguracionCalculo) -> tuple[MapaColumnas, list[Anomalia]]:
    anomalias: list[Anomalia] = []
    for pcs in range(1, cfg.total_pcs + 1):
        for texto_campo, campo in CAMPOS.items():
            col = columna(pcs, campo)
            v = encabezado.get(col)
            esperado = f"Arena - PCS {pcs:02d} - POWERELECTRONICS {texto_campo}"
            if es_vacio(v):
                anomalias.append(
                    Anomalia(
                        "columna_faltante",
                        "error",
                        numero_fila=1,
                        numero_pcs=pcs,
                        detalle=f"col {col}: se esperaba {esperado!r}",
                    )
                )
                continue
            if v == esperado:
                continue
            m = PATRON.match(str(v))
            tipo = "columna_desplazada" if m and CAMPOS.get(m.group(2)) is not None else "nombre_inesperado"
            anomalias.append(
                Anomalia(
                    tipo, "error", numero_fila=1, numero_pcs=pcs, detalle=f"col {col}: {v!r}, se esperaba {esperado!r}"
                )
            )

    p_eventos = pcs_recorridos_por_eventos(encabezado)
    if p_eventos != cfg.total_pcs:
        anomalias.append(
            Anomalia(
                "pcs_distinto_c2",
                "error",
                numero_fila=1,
                detalle=f"mcoCreateList recorrería {p_eventos} PCS y cmdCalcAvailability {cfg.total_pcs} (C2) — F-17",
            )
        )
    return MapaColumnas(list(range(1, cfg.total_pcs + 1)), p_eventos), anomalias

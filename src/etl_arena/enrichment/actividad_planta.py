"""Factores de ``PlantActivity`` asociados **por fila** a ``RawData-PCS`` (R6, F-01, F-31).

El VBA lee ``Sheet4.Cells(dblRec, 3|4)`` con la fila de ``RawData-PCS``: no hay join por
timestamp. Una celda vacía vale 0. El timestamp de ``PlantActivity!B`` solo se usa para
reportar desalineaciones.
"""

from __future__ import annotations

import numpy as np

from etl_arena.excel_semantics import serial_a_datetime
from etl_arena.model import Anomalia, DatosActividad, ErrorParidad, MatrizPCS

COL_TS, COL_OPERACIONAL, COL_EXCUSABLE, COL_SOC = 2, 3, 4, 9
TOLERANCIA_ALINEACION = 1e-6  # días (~0,09 s)


def _factor(v: object, fila: int, col: str) -> float | None:
    if v is None:
        return None
    if isinstance(v, float):
        return v
    raise ErrorParidad(
        f"PlantActivity!{col}{fila} no numérico: {v!r} (el VBA fallaría al multiplicar)",
        [Anomalia("pa_no_numerico", "error", numero_fila=fila, detalle=f"{col}: {v!r}")],
    )


def asociar_actividad(m: MatrizPCS, actividad: dict[int, dict[int, object]]) -> tuple[DatosActividad, list[Anomalia]]:
    n = m.n
    serial_pa = np.full(n, np.nan)
    operacional = np.zeros(n)
    excusable = np.zeros(n)
    soc = np.full(n, np.nan)
    anomalias: list[Anomalia] = []

    for i in range(n):
        fila = int(m.numero_fila[i])
        celdas = actividad.get(fila, {})
        ts = celdas.get(COL_TS)
        if isinstance(ts, float):
            serial_pa[i] = ts
            if m.serial[i] != 0.0 and abs(ts - m.serial[i]) > TOLERANCIA_ALINEACION:
                anomalias.append(
                    Anomalia(
                        "pa_desalineado",
                        "advertencia",
                        numero_fila=fila,
                        serial=float(m.serial[i]),
                        detalle=f"PlantActivity!B={serial_a_datetime(ts):%Y-%m-%d %H:%M} ≠ "
                        f"RawData-PCS!A={serial_a_datetime(float(m.serial[i])):%Y-%m-%d %H:%M}",
                    )
                )
        else:
            anomalias.append(
                Anomalia(
                    "pa_sin_timestamp",
                    "advertencia",
                    numero_fila=fila,
                    serial=float(m.serial[i]),
                    detalle="PlantActivity!B vacío" if ts is None else f"PlantActivity!B={ts!r}",
                )
            )
        for col, letra, destino in ((COL_OPERACIONAL, "C", operacional), (COL_EXCUSABLE, "D", excusable)):
            f = _factor(celdas.get(col), fila, letra)
            if f is None:
                anomalias.append(
                    Anomalia("pa_vacio", "advertencia", numero_fila=fila, detalle=f"PlantActivity!{letra} vacío → 0")
                )
            else:
                destino[i] = f
        v_soc = celdas.get(COL_SOC)
        if isinstance(v_soc, float):
            soc[i] = v_soc

    return DatosActividad(m.numero_fila.copy(), serial_pa, operacional, excusable, soc), anomalias

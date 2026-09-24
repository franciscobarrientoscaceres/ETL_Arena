"""``Exclusion_Matrix`` asociada **por fila** a ``RawData-PCS`` (F-37, D-15).

Estructura (libro de agosto 2026): fila 1 = ``Date/time``, ``PCS01`` … ``PCSnn``,
``Excused Event``, ``Comments``; filas 2.. alineadas con ``RawData-PCS`` (misma fila Excel).
Valor por celda: ``0``/vacío = sin evento de exclusión (EE); ``1`` = EE, se consideran todos
los módulos en operación; ``2`` = EE, se consideran los módulos en operación **antes** del
inicio del evento.

``baterias_previas`` (solo celdas con valor 2): baterías indisponibles *consideradas* en la
fila anterior al primer 2 del tramo contiguo del PCS (hoja completa, no solo el período):
``C3 − M`` si esa fila tenía ``M < C3`` y valor 0; ``0`` si tenía todos los módulos, estaba
vacía o tenía valor 1. Es la definición de negocio y coincide con la macro de agosto
(``F(r) = F(r-1)``, encadenado) en todo el libro de agosto; donde podría diferir se reporta.
"""

from __future__ import annotations

import numpy as np

from etl_arena.config import ConfiguracionCalculo
from etl_arena.excel_semantics import serial_a_datetime
from etl_arena.model import VALORES_EXCLUSION, Anomalia, DatosExclusion, ErrorParidad, MatrizPCS

TOLERANCIA_ALINEACION = 1e-6  # días


def _encabezado_esperado(p: int) -> dict[int, str]:
    return {1: "Date/time", **{k + 1: f"PCS{k:02d}" for k in range(1, p + 1)}}


def asociar_exclusion(
    m: MatrizPCS, exclusion: dict[int, dict[int, object]] | None, cfg: ConfiguracionCalculo
) -> tuple[DatosExclusion, list[Anomalia]]:
    p = cfg.total_pcs
    if exclusion is None:
        return DatosExclusion.vacia(m.numero_fila, p), []

    anomalias: list[Anomalia] = []
    encabezado = exclusion.get(1, {})
    for col, esperado in _encabezado_esperado(p).items():
        if encabezado.get(col) != esperado:
            raise ErrorParidad(
                f"Exclusion_Matrix: encabezado de la columna {col} = {encabezado.get(col)!r}, se esperaba {esperado!r}",
                [Anomalia("em_encabezado", "error", numero_fila=1, detalle=f"col {col}")],
            )

    n = m.n
    serial_em = np.full(n, np.nan)
    valor = np.zeros((n, p))
    evento_excusado = np.full(n, np.nan)
    comentario = np.full(n, None, dtype=object)
    for i in range(n):
        fila = int(m.numero_fila[i])
        celdas = exclusion.get(fila, {})
        ts = celdas.get(1)
        marcas = False
        for j in range(p):
            v = celdas.get(j + 2)
            if v is None:
                continue
            if not isinstance(v, float) or v not in VALORES_EXCLUSION:
                raise ErrorParidad(
                    f"Exclusion_Matrix fila {fila}, PCS {j + 1}: valor {v!r} (se admite 0, 1, 2 o vacío)",
                    [Anomalia("em_valor_invalido", "error", numero_fila=fila, numero_pcs=j + 1, detalle=repr(v))],
                )
            valor[i, j] = v
            marcas = marcas or v != 0.0
        if isinstance(ts, float):
            serial_em[i] = ts
            if m.serial[i] != 0.0 and abs(ts - m.serial[i]) > TOLERANCIA_ALINEACION:
                anomalias.append(
                    Anomalia(
                        "em_desalineado",
                        "advertencia",
                        numero_fila=fila,
                        serial=float(m.serial[i]),
                        detalle=f"Exclusion_Matrix!A={serial_a_datetime(ts):%Y-%m-%d %H:%M} ≠ "
                        f"RawData-PCS!A={serial_a_datetime(float(m.serial[i])):%Y-%m-%d %H:%M}",
                    )
                )
        elif marcas:
            anomalias.append(
                Anomalia(
                    "em_sin_timestamp",
                    "advertencia",
                    numero_fila=fila,
                    detalle="fila con eventos de exclusión y sin Exclusion_Matrix!A",
                )
            )
        ev = celdas.get(p + 2)
        if isinstance(ev, float):
            evento_excusado[i] = ev
        if celdas.get(p + 3) not in (None, ""):
            comentario[i] = str(celdas[p + 3])

    previas, anomalias_previas = _baterias_previas(m, valor, cfg)
    return DatosExclusion(m.numero_fila.copy(), serial_em, valor, previas, evento_excusado, comentario), [
        *anomalias,
        *anomalias_previas,
    ]


def _baterias_previas(m: MatrizPCS, valor: np.ndarray, cfg: ConfiguracionCalculo) -> tuple[np.ndarray, list[Anomalia]]:
    c3 = float(cfg.baterias_por_pcs)
    n, p = valor.shape
    previas = np.full((n, p), np.nan)
    anomalias: list[Anomalia] = []
    for j in range(p):
        i = 0
        while i < n:
            if valor[i, j] != 2.0:
                i += 1
                continue
            inicio = i
            while i < n and valor[i, j] == 2.0:
                i += 1
            fin = i  # tramo [inicio, fin)
            pcs = m.pcs[j]
            if inicio == 0:
                # sin fila previa (la anterior es el encabezado): se toma la primera fila del tramo
                ref = inicio
                anomalias.append(
                    Anomalia(
                        "ee2_sin_fila_previa",
                        "advertencia",
                        numero_fila=int(m.numero_fila[0]),
                        numero_pcs=pcs,
                        detalle="EE valor 2 desde la primera fila: se usan sus módulos",
                    )
                )
            else:
                ref = inicio - 1
            if m.modulos_nulo[ref, j] or not m.modulos[ref, j] < c3 or valor[ref, j] == 1.0 and ref != inicio:
                b = 0.0
            else:
                b = c3 - m.modulos[ref, j]
            previas[inicio:fin, j] = b
            # donde la macro de agosto (F(r) = F(r-1)) daría otro valor: una fila del tramo con
            # todos los módulos corta la copia y las siguientes quedan en 0
            sin_falla = [
                int(m.numero_fila[r]) for r in range(inicio, fin) if m.modulos_nulo[r, j] or not m.modulos[r, j] < c3
            ]
            if b != 0.0 and sin_falla and sin_falla[0] < int(m.numero_fila[fin - 1]):
                anomalias.append(
                    Anomalia(
                        "ee2_difiere_macro_agosto",
                        "info",
                        numero_fila=sin_falla[0],
                        numero_pcs=pcs,
                        detalle="fila sin falla dentro de un EE valor 2: la macro de agosto habría dejado 0 después",
                    )
                )
    return previas, anomalias

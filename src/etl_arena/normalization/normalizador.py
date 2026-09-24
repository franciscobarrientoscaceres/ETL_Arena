"""Formato ancho de ``RawData-PCS`` → ``MatrizPCS`` (motores) y formato largo (persistencia) — R5."""

from __future__ import annotations

import numpy as np
import pandas as pd

from etl_arena.config import ConfiguracionCalculo
from etl_arena.excel_semantics import serial_a_datetime
from etl_arena.ingestion import LibroCrudo
from etl_arena.model import Anomalia, ErrorParidad, MatrizPCS
from etl_arena.normalization.esquema import MapaColumnas, columna


def _modulo(v: object, fila: int, pcs: int, cfg: ConfiguracionCalculo) -> float:
    """Valor de NUMBER OF MODULES; NaN si vacío. Todo lo que haría fallar o desviar al VBA
    se rechaza (data contract §2.3, R4.6, F-16): texto (incluido ``""``: ``"" < 4`` es
    *Type mismatch*), booleanos y valores fuera de ``[0, baterias_por_pcs]``."""
    if v is None:
        return np.nan
    if isinstance(v, float) and 0.0 <= v <= cfg.baterias_por_pcs:
        return v
    tipo = "modulos_texto" if isinstance(v, str) else "modulos_fuera_de_rango"
    raise ErrorParidad(
        f"RawData-PCS fila {fila}, PCS {pcs}: NUMBER OF MODULES = {v!r} (el VBA fallaría o se desviaría)",
        [Anomalia(tipo, "error", numero_fila=fila, numero_pcs=pcs, detalle=repr(v))],
    )


def a_matriz(libro: LibroCrudo, mapa: MapaColumnas, cfg: ConfiguracionCalculo) -> MatrizPCS:
    """Construye la matriz por **posición** de columna y en el orden de fila de la hoja (R5.6)."""
    n, p = len(libro.filas), len(mapa.pcs)
    modulos = np.full((n, p), np.nan)
    falla = np.empty((n, p), dtype=object)
    estado = np.empty((n, p), dtype=object)
    advertencia = np.empty((n, p), dtype=object)
    serial = np.empty(n)
    numero_fila = np.empty(n, dtype=np.int64)
    marca = np.empty(n, dtype="datetime64[s]")
    cols = [
        (columna(k, "falla"), columna(k, "estado"), columna(k, "advertencia"), columna(k, "modulos")) for k in mapa.pcs
    ]

    for i, fila in enumerate(libro.filas):
        numero_fila[i] = fila.numero_fila
        serial[i] = fila.serial
        marca[i] = (
            np.datetime64("NaT")
            if fila.serial == 0.0 and fila.numero_fila == 2
            else np.datetime64(serial_a_datetime(fila.serial), "s")
        )
        c = fila.celdas
        for j, (cf, ce, ca, cm) in enumerate(cols):
            falla[i, j] = c.get(cf)
            estado[i, j] = c.get(ce)
            advertencia[i, j] = c.get(ca)
            modulos[i, j] = _modulo(c.get(cm), fila.numero_fila, mapa.pcs[j], cfg)

    fila_sig = libro.filas[-1].numero_fila + 1 if libro.filas else 2
    sig = np.array(
        [_modulo(libro.fila_siguiente.get(cm), fila_sig, mapa.pcs[j], cfg) for j, (_, _, _, cm) in enumerate(cols)],
        dtype=float,
    )
    return MatrizPCS(
        numero_fila=numero_fila,
        serial=serial,
        marca_tiempo=marca,
        modulos=modulos,
        modulos_nulo=np.isnan(modulos),
        falla=falla,
        estado=estado,
        advertencia=advertencia,
        pcs=list(mapa.pcs),
        encabezado_falla=[str(libro.encabezado.get(cols[j][0], "")) for j in range(p)],
        encabezado_modulos=[str(libro.encabezado.get(cols[j][3], "")) for j in range(p)],
        siguiente_modulos=sig,
        siguiente_nulo=np.isnan(sig),
    )


def a_formato_largo(m: MatrizPCS, baterias_por_pcs: int) -> pd.DataFrame:
    """Una fila por ``(NumeroFilaOrigen, NumeroPCS)`` para ``raw_pcs_sample`` (R5.1, R5.3, R5.4)."""
    n, p = m.n, m.p
    falla = m.falla.ravel()
    return pd.DataFrame(
        {
            "NumeroFilaOrigen": np.repeat(m.numero_fila, p),
            "NumeroPCS": np.tile(np.asarray(m.pcs), n),
            "SerialFechaExcelOrigen": np.repeat(m.serial, p),
            "MarcaTiempoLocalOrigen": np.repeat(m.marca_tiempo, p),
            "FallaRaw": falla,
            "FallaRawEsNumero": np.array([isinstance(v, float) for v in falla], dtype=bool),
            "EstadoRaw": m.estado.ravel(),
            "AdvertenciaRaw": m.advertencia.ravel(),
            "ModulosRaw": m.modulos.ravel(),
            "ModulosDisponibles": m.modulos_disponibles(baterias_por_pcs).ravel(),
            "ModulosDisponiblesNulo": m.modulos_nulo.ravel(),
        }
    )

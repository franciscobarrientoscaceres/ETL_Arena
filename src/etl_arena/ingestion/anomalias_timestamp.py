"""Anomalías de la columna A de ``RawData-PCS`` (R4.4, R16.3, F-07, F-32).

Solo reporta: en modo paridad nada se corrige, reordena ni deduplica.
"""

from __future__ import annotations

from collections.abc import Sequence
from datetime import UTC, datetime, timedelta
from zoneinfo import ZoneInfo

from etl_arena.excel_semantics import redondear_excel, serial_a_datetime
from etl_arena.model import Anomalia


def _hora_inexistente(dt: datetime, tz: ZoneInfo) -> bool:
    """Hora local que no existe (salto de primavera): la ida y vuelta por UTC la mueve."""
    vuelta = dt.replace(tzinfo=tz).astimezone(UTC).astimezone(tz).replace(tzinfo=None)
    return vuelta != dt


def _hora_ambigua(dt: datetime, tz: ZoneInfo) -> bool:
    """Hora local repetida (retroceso de otoño): ``fold=0`` y ``fold=1`` dan offsets distintos."""
    return dt.replace(tzinfo=tz, fold=0).utcoffset() != dt.replace(tzinfo=tz, fold=1).utcoffset()


def detectar_anomalias_timestamp(
    seriales: Sequence[float],
    numeros_fila: Sequence[int],
    minutos_muestreo: int,
    zona_horaria: str = "America/Santiago",
) -> list[Anomalia]:
    """Compara cada fila con la anterior.

    * ``Δ = ROUND(Δserial × 1440, 2)`` minutos.
    * ``Δ = 0``: ``dst_repeticion`` (info) si la hora local es ambigua; si no ``duplicado``.
    * ``Δ < 0``: ``fuera_de_orden``.
    * ``Δ`` múltiplo de ``minutos_muestreo`` y mayor: ``dst_salto`` (info) si todas las horas
      faltantes son inexistentes en la zona; si no ``hueco``.
    * otro ``Δ``: ``frecuencia_distinta``.
    """
    tz = ZoneInfo(zona_horaria)
    anomalias: list[Anomalia] = []
    for k in range(1, len(seriales)):
        anterior, actual = seriales[k - 1], seriales[k]
        delta = redondear_excel((actual - anterior) * 24 * 60, 2)
        if delta == minutos_muestreo:
            continue
        fila = numeros_fila[k]
        t_ant, t_act = serial_a_datetime(anterior), serial_a_datetime(actual)
        contexto = f"{t_ant:%Y-%m-%d %H:%M} → {t_act:%Y-%m-%d %H:%M} (Δ={delta:g} min)"
        if delta == 0:
            tipo, sev = ("dst_repeticion", "info") if _hora_ambigua(t_act, tz) else ("duplicado", "advertencia")
        elif delta < 0:
            tipo, sev = "fuera_de_orden", "advertencia"
        elif delta % minutos_muestreo == 0:
            faltantes = [t_ant + timedelta(minutes=m) for m in range(minutos_muestreo, int(delta), minutos_muestreo)]
            if all(_hora_inexistente(t, tz) for t in faltantes):
                tipo, sev = "dst_salto", "info"
            else:
                tipo, sev = "hueco", "advertencia"
            contexto += f"; faltan {len(faltantes)} filas"
        else:
            tipo, sev = "frecuencia_distinta", "advertencia"
        anomalias.append(Anomalia(tipo, sev, numero_fila=fila, serial=actual, detalle=contexto))
    return anomalias

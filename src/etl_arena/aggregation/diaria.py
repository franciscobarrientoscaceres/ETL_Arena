"""``mcoDailyAvailability`` + fórmulas de ``Daily`` (R9, F-08, F-09, F-33).

VBA de referencia (Module1)::

    i = 9 : k = 4 : Suma = 0
    Do
        If (Calc!E(k) >= Daily!C(i)) And (Calc!E(k) < Daily!C(i) + 1) Then
            For j = 1 To C2 : Suma = Suma + 12 * Calc!(k, 5 + j) : Next
            k = k + 1
        Else
            Daily!D(i) = Suma : i = i + 1 : Suma = 0
        End If
        If Daily!C(i) = "" Then Exit Do
    Loop

La tabla ``Calc!F:BN`` guarda baterías ponderadas por el factor excusable pero **sin** el
factor operacional (F-08). Fórmulas: ``E9 = D9``, ``E(n) = D(n) + E(n-1)``,
``F = IFERROR(1-E/(Total_Racks*24*60*B/C23),"")``, ``G9 = 0``, ``G(n) = IFERROR(F(n)-F(n-1),"")``.
"""

from __future__ import annotations

from datetime import timedelta

from etl_arena.availability import ResultadoDisponibilidad
from etl_arena.config import ConfiguracionCalculo
from etl_arena.excel_semantics import datetime_a_serial
from etl_arena.model import DiaDisponibilidad


def calcular_diaria(res: ResultadoDisponibilidad, cfg: ConfiguracionCalculo) -> list[DiaDisponibilidad]:
    n_dias = cfg.dias_diario
    inicio = datetime_a_serial(cfg.inicio_periodo)
    dias = [inicio + d for d in range(n_dias)]  # Daily!C9.. = C5, C9+1, …
    racks = float(cfg.racks_por_pcs)
    tabla = res.ponderadas_o_cero().tolist()  # celda vacía = 0
    serial_e = res.serial.tolist()
    filas = len(serial_e)

    diario = [0.0] * n_dias
    i, k, suma = 0, 0, 0.0
    while True:
        e = serial_e[k] if k < filas else 0.0  # E(k) vacía = 0
        if e >= dias[i] and e < dias[i] + 1:
            for valor in tabla[k]:
                suma = suma + racks * valor
            k += 1
        else:
            diario[i] = suma
            i += 1
            suma = 0.0
        if i >= n_dias:  # Daily!C(i) = ""
            break

    c23 = res.minutos_muestreo_derivado
    salida: list[DiaDisponibilidad] = []
    acumulado, disp_anterior = 0.0, None
    for n, d in enumerate(diario, start=1):
        acumulado = d if n == 1 else d + acumulado  # E9 = D9 ; E(n) = D(n) + E(n-1)
        try:
            disp = 1 - acumulado / (cfg.total_racks * 24 * 60 * n / c23)  # type: ignore[operator]
        except (TypeError, ZeroDivisionError):
            disp = None  # IFERROR(…, "")
        if n == 1:
            var = 0.0  # G9 = 0 (valor fijo)
        elif disp is None or disp_anterior is None:
            var = None
        else:
            var = disp - disp_anterior
        salida.append(DiaDisponibilidad(n, cfg.inicio_periodo + timedelta(days=n - 1), d, acumulado, disp, var))
        disp_anterior = disp
    return salida

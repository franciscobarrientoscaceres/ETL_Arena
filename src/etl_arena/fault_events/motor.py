"""``MotorEventosFalla``: emulación de ``mcoCreateList`` (R8, F-03, F-04, F-05, F-06, F-11, F-18, F-37).

VBA de referencia (Module1), para cada bloque de PCS (``intColRec`` = 2, 6, …) y cada fila::

    If A(r) >= L2 And A(r) < (L4 + 1) Then
      If M(r) <> "" And M(r) < 4 Then
        If L14 = "Yes" Then sumablocks = sumablocks + <ponderada por Exclusion_Matrix>  ' F-37 (antes: * PA!D)
        Else sumablocks = sumablocks + 4 - M(r)
        numBlock = numBlock + 1
        If (M(r-1) = 4) Or (M(r-1) = "") Or (A(r-1) < L2) Then          ' inicio
          B = PCS : C = A(r) - C23/(24*60) : G = FAULT(r)
          If G = "NO FAULTS" Then G = IIf(FAULT(r-1) <> "NO FAULTS", FAULT(r-1), "F13 NO MODULES")
          If G = "" Then G = "F1 Watchdog"
          F = IFERROR(MID(G,1,FIND(" ",G,1)-1),CONCATENATE("F",G))
        If (M(r+1) = 4) Or (M(r+1) = "") Or (A(r+1) > (L4 + 1)) Then     ' cierre
          D = A(r) : E = 24*(D-C) : H = sumablocks/numBlock : I = 12*E*H
          L10 = L10 + 12*E*H : sumablocks = 0 : numBlock = 0 : dblResult = dblResult + 1

``sumablocks``/``numBlock`` no se reinician al cambiar de PCS y ``dblResult`` solo avanza al
cerrar: un evento abierto al final del período se "arrastra" al siguiente PCS (F-06).
"""

from __future__ import annotations

from dataclasses import dataclass, field

from etl_arena.config import ConfiguracionCalculo
from etl_arena.excel_semantics import codigo_falla_excel, datetime_a_serial, es_vacio, fecha_vba_a_celda, igual_texto
from etl_arena.fault_events.resumen import FilaResumenCodigo, resumen_por_codigo
from etl_arena.model import Anomalia, DatosExclusion, EventoFalla, MatrizPCS, RegistroLista

MAX_INTEGER_VBA = 32767  # numBlock As Integer
F13 = "F13 NO MODULES"
F1 = "F1 Watchdog"


@dataclass
class ResultadoEventos:
    cerrados: list[RegistroLista]  # filas B:I completas, en orden de escritura
    incompleto: RegistroLista | None  # fila con B,C,F,G escritos pero sin cierre (F-06)
    horas_rack_totales: float  # L10
    horas_periodo: float  # L8 = (L4 - L2 + 1) * 24
    disponibilidad_periodo: float | None  # L12 = 1 - L10/(L6*L8)
    resumen: list[FilaResumenCodigo] = field(default_factory=list)  # N:Q
    anomalias: list[Anomalia] = field(default_factory=list)

    def eventos(self) -> list[EventoFalla]:
        return [EventoFalla.desde_registro(r) for r in self.cerrados]


def _descripcion(actual: object, anterior: object, hay_anterior: bool) -> tuple[object, bool]:
    """Árbol de descripción de ``mcoCreateList`` (R8.8, F-04). Devuelve ``(G, fallback)``."""
    g, fallback = actual, False
    if igual_texto(g, "NO FAULTS"):
        prev = anterior if hay_anterior else None
        if not igual_texto(prev, "NO FAULTS"):
            g, fallback = prev, True  # anterior vacío → G vacío → "F1 Watchdog"
        else:
            g = F13
    if es_vacio(g):
        g = F1
    return g, fallback


def detectar_eventos(
    m: MatrizPCS,
    exc: DatosExclusion | None,
    cfg: ConfiguracionCalculo,
    c23: float | None,
    catalogo_codigos: list[tuple[str, str]] | None = None,
) -> ResultadoEventos:
    """``c23``: ``Calculation-Availability!C23`` de la corrida KPI (minutos entre las dos
    primeras filas procesadas). ``catalogo_codigos``: ``[(codigo, descripcion_pe)]`` de N:O."""
    if c23 is None:
        raise ValueError("C23 indefinido (menos de 2 filas en el período KPI): no se pueden calcular inicios")
    l2 = datetime_a_serial(cfg.inicio_periodo_eventos)
    l4_mas_uno = datetime_a_serial(cfg.fin_periodo_eventos) + 1
    umbral = float(cfg.baterias_por_pcs)  # literal 4 del VBA
    racks = float(cfg.racks_por_pcs)  # literal 12 del VBA
    desplazamiento = c23 / (24 * 60)
    excusable_eventos = cfg.aplicar_evento_excusable_eventos

    n = m.n
    serial = m.serial.tolist()
    modulos = m.modulos.tolist()
    nulos = m.modulos_nulo.tolist()
    if exc is None:
        exc = DatosExclusion.vacia(m.numero_fila, m.p)
    em, previas = exc.valor.tolist(), exc.baterias_previas.tolist()
    sig_mod = m.siguiente_modulos.tolist()
    sig_nulo = m.siguiente_nulo.tolist()

    anomalias: list[Anomalia] = []
    cerrados: list[RegistroLista] = []
    actual = RegistroLista(orden_excel=1)
    suma_bloques, num_bloques, l10 = 0.0, 0, 0.0

    for j, pcs in enumerate(m.pcs):  # bloques mientras haya encabezado (validado = C2, F-17)
        for i in range(n):
            s = serial[i]
            if not (s >= l2 and s < l4_mas_uno):
                continue
            if nulos[i][j]:
                continue
            x = modulos[i][j]
            if not x < umbral:
                continue

            if excusable_eventos:
                # misma ponderación que el KPI (F-37): valor 2 → baterías previas; 0/1 → (C3 − M)·(1 − EM)
                e = em[i][j]
                suma_bloques = suma_bloques + (previas[i][j] if e == 2.0 else (umbral - x) * (1 - e))
            else:
                suma_bloques = suma_bloques + umbral - x  # (sumablocks + 4) - x
            num_bloques += 1
            if em[i][j] != 0.0:
                actual.tiene_exclusion = True
            if num_bloques == MAX_INTEGER_VBA + 1:
                actual.excel_habria_fallado = True
                anomalias.append(
                    Anomalia(
                        "excel_overflow_numblock",
                        "advertencia",
                        numero_fila=int(m.numero_fila[i]),
                        numero_pcs=pcs,
                        detalle="numBlock > 32767: la macro fallaría con Overflow",
                    )
                )

            # ---- inicio (R8.4) ----
            if i == 0:
                inicia = True
                if m.numero_fila[0] == 2:
                    # la anterior es el encabezado: "Arena - PCS …" = 4 → Type mismatch (F-11, D-08)
                    actual.excel_habria_fallado = True
                    anomalias.append(
                        Anomalia(
                            "excel_habria_fallado",
                            "advertencia",
                            numero_fila=2,
                            numero_pcs=pcs,
                            detalle="evento en la fila 2: mcoCreateList falla (D-08)",
                        )
                    )
            else:
                prev_nulo = nulos[i - 1][j]
                inicia = (not prev_nulo and modulos[i - 1][j] == umbral) or prev_nulo or serial[i - 1] < l2
            if inicia:
                if actual.numero_pcs is not None:  # fila escrita y no cerrada: se sobrescribe
                    actual.arrastrado_excel = True
                actual.numero_pcs = pcs
                actual.pcs_iniciados.add(pcs)
                actual.serial_inicio = fecha_vba_a_celda(s - desplazamiento)  # Date → celda (al segundo)
                g, fallback = _descripcion(m.falla[i, j], m.falla[i - 1, j] if i > 0 else None, i > 0)
                actual.descripcion_falla = g
                actual.fallback = fallback
                actual.codigo_falla = codigo_falla_excel(g)

            # ---- cierre (R8.5) ----
            if i + 1 < n:
                sig_es_nulo, sig_valor, sig_serial = nulos[i + 1][j], modulos[i + 1][j], serial[i + 1]
            else:  # fila donde corta el VBA (A vacía = 0)
                sig_es_nulo, sig_valor, sig_serial = sig_nulo[j], sig_mod[j], 0.0
            cierra = (not sig_es_nulo and sig_valor == umbral) or sig_es_nulo or sig_serial > l4_mas_uno
            if cierra:
                if len(actual.pcs_iniciados) > 1 or pcs not in actual.pcs_iniciados:
                    actual.arrastrado_excel = True
                actual.cerrado_por_nulo = bool(sig_es_nulo)
                actual.serial_fin = fecha_vba_a_celda(s)  # última fila en falla (F-03)
                actual.duracion_horas = 24 * (actual.serial_fin - actual.serial_inicio)
                actual.promedio_baterias = suma_bloques / num_bloques
                actual.horas_rack = racks * actual.duracion_horas * actual.promedio_baterias
                actual.numero_bloques, actual.suma_bloques = num_bloques, suma_bloques
                l10 = l10 + racks * actual.duracion_horas * actual.promedio_baterias
                cerrados.append(actual)
                actual = RegistroLista(orden_excel=actual.orden_excel + 1)
                suma_bloques, num_bloques = 0.0, 0

    incompleto = actual if actual.numero_pcs is not None else None
    if incompleto is not None:
        anomalias.append(
            Anomalia(
                "evento_incompleto_excel",
                "advertencia",
                numero_pcs=incompleto.numero_pcs,
                detalle=f"ListOfFaults fila {incompleto.orden_excel + 5}: evento sin cierre (F-06)",
            )
        )
    for r in cerrados:
        if r.arrastrado_excel:
            anomalias.append(
                Anomalia(
                    "evento_arrastrado_excel",
                    "advertencia",
                    numero_pcs=r.numero_pcs,
                    detalle=f"ListOfFaults fila {r.orden_excel + 5}: PCS {sorted(r.pcs_iniciados)} (F-06)",
                )
            )

    horas_periodo = (datetime_a_serial(cfg.fin_periodo_eventos) - l2 + 1) * 24
    den = cfg.total_racks * horas_periodo
    return ResultadoEventos(
        cerrados=cerrados,
        incompleto=incompleto,
        horas_rack_totales=l10,
        horas_periodo=horas_periodo,
        disponibilidad_periodo=1 - l10 / den if den else None,
        resumen=resumen_por_codigo(cerrados, catalogo_codigos or [], l10),
        anomalias=anomalias,
    )

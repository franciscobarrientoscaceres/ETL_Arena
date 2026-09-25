"""Notificación a Power BI / Misael (tarea 4.8; R18.1–R18.3).

Con ``ETL_ARENA_NOTIFY_WEBHOOK`` hace POST JSON (p. ej. un Incoming Webhook de Teams o Power
Automate); sin webhook, o si el POST falla, escribe ``notificacion.md`` en la carpeta del corte.
Nunca incluye credenciales ni datos crudos: solo identificadores, período, estado y resúmenes.
No intenta refrescar Power BI por API (fuera de alcance hasta tener service principal).
"""

from __future__ import annotations

import json
import logging
import os
import urllib.request
from pathlib import Path

log = logging.getLogger("etl_arena.notificacion")
VARIABLE_WEBHOOK = "ETL_ARENA_NOTIFY_WEBHOOK"


def construir_mensaje(
    id_corrida: str,
    periodo: tuple[str, str],
    estado: str,
    estado_exclusiones: str,
    kpi: dict,
    reconciliacion: dict | None,
    calidad: dict | None,
    cierres_pendientes: list[str] | None = None,
) -> dict:
    etiqueta = "Con Exclusiones" if estado_exclusiones == "con_exclusiones" else "Sin Exclusiones"
    anomalias = (calidad or {}).get("anomalias", {})
    return {
        "titulo": f"ETL Arena — {estado.upper()} — {periodo[0]} a {periodo[1]} ({etiqueta})",
        "id_corrida": id_corrida,
        "periodo": {"inicio": periodo[0], "fin": periodo[1]},
        "estado": estado,
        "exclusiones": etiqueta,
        "kpi": kpi,
        "reconciliacion": (reconciliacion or {}).get("estado", "sin_referencia"),
        "anomalias": {"total": anomalias.get("total", 0), "por_severidad": anomalias.get("por_severidad", {})},
        "cierres_pendientes": list(cierres_pendientes or []),
        "accion": (
            "Actualizar el dataset de Power BI (vistas v_*_vigente)."
            if estado == "success"
            else "No actualizar Power BI: revisar la corrida (etl_run / reconciliation_result)."
        ),
    }


def a_markdown(m: dict) -> str:
    k = m["kpi"]
    lineas = [
        f"# {m['titulo']}",
        "",
        f"- **IdCorrida:** `{m['id_corrida']}`",
        f"- **Estado:** {m['estado']} · **Reconciliación:** {m['reconciliacion']} · **{m['exclusiones']}**",
        f"- **KPI:** C12 = {k.get('C12')}, C14 = {k.get('C14')}, disponibilidad del período (C16) = {k.get('C16')}",
        f"- **Anomalías:** {m['anomalias']['total']} {m['anomalias']['por_severidad']}",
    ]
    if m.get("cierres_pendientes"):
        lineas.append(
            f"- **Cierres mensuales pendientes:** {', '.join(m['cierres_pendientes'])} "
            "(esperan la Exclusion_Matrix de Alex; `run_lunes.py --stage cierre-mensual`)"
        )
    lineas += [
        "",
        f"**Acción:** {m['accion']}",
        "",
    ]
    return "\n".join(lineas)


def enviar(mensaje: dict, dir_corte: str | Path, webhook: str | None = None) -> str:
    """Devuelve ``"webhook"`` o la ruta del ``notificacion.md`` escrito."""
    webhook = webhook if webhook is not None else os.environ.get(VARIABLE_WEBHOOK, "")
    if webhook:
        try:
            datos = json.dumps({"text": a_markdown(mensaje), **mensaje}, ensure_ascii=False, default=str).encode()
            req = urllib.request.Request(webhook, data=datos, headers={"Content-Type": "application/json"})
            with urllib.request.urlopen(req, timeout=30) as resp:
                if 200 <= resp.status < 300:
                    log.info("notificación enviada al webhook (%s)", resp.status)
                    return "webhook"
        except Exception as exc:  # no bloquear la corrida por la notificación
            log.warning("webhook falló (%s); se escribe notificacion.md", exc)
    ruta = Path(dir_corte) / "notificacion.md"
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(a_markdown(mensaje), encoding="utf-8")
    print(a_markdown(mensaje))
    return str(ruta)

"""Estado entre cortes del orquestador semanal (tarea 4.9; R3.1, R19.3, D-07, D-17).

* **Libro base** (R3.1): el libro de trabajo de la última corrida oficial exitosa es la base de la
  siguiente; en la primera corrida, el maestro de ``data/``. Se guarda como puntero en
  ``data/work/libro_base.json``: el libro base nunca se edita, cada corte trabaja sobre su copia.
* **Cola de cierres** (R19.3): cuando los datos cargados cubren el último bloque de un mes sin
  ``cierre_mensual`` oficial, el mes se encola en ``data/work/cola_cierres.json``. El cierre oficial
  se ejecuta con ``run_lunes.py --stage cierre-mensual`` cuando Alex entrega la ``Exclusion_Matrix``
  (D-17) o, explícitamente, ``--sin-exclusiones`` (R19.6).
"""

from __future__ import annotations

import json
from calendar import monthrange
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

from etl_arena.excel_semantics import datetime_a_serial

ARCHIVO_LIBRO_BASE = "libro_base.json"
ARCHIVO_COLA = "cola_cierres.json"


def _ahora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _leer(ruta: Path, defecto):
    return json.loads(ruta.read_text(encoding="utf-8")) if ruta.exists() else defecto


def _escribir(ruta: Path, datos) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    tmp = ruta.with_suffix(ruta.suffix + ".tmp")
    tmp.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(ruta)  # nunca queda un puntero a medio escribir


# ------------------------------------------------------------------ libro base
def libro_base(work: Path, maestro: Path) -> Path:
    """Libro de trabajo de la última corrida oficial exitosa, o el maestro si aún no hay ninguna."""
    puntero = _leer(Path(work) / ARCHIVO_LIBRO_BASE, None)
    if puntero is None:
        if not Path(maestro).exists():
            raise FileNotFoundError(f"no hay libro base promovido ni maestro en {maestro}")
        return Path(maestro)
    ruta = Path(puntero["ruta"])
    if not ruta.exists():
        raise FileNotFoundError(f"el libro base {ruta} (corte {puntero['corte']}) ya no existe")
    return ruta


def promover_libro_base(work: Path, libro: Path, corte: str, id_corrida: str, sha256: str) -> dict:
    """Registra ``libro`` como base de la próxima corrida (solo tras una corrida oficial ``success``)."""
    puntero = {
        "ruta": str(Path(libro).resolve()),
        "corte": corte,
        "id_corrida": id_corrida,
        "sha256": sha256,
        "promovido_en": _ahora(),
    }
    _escribir(Path(work) / ARCHIVO_LIBRO_BASE, puntero)
    return puntero


# ------------------------------------------------------------------ meses completos
def ultimo_dia(anio: int, mes: int) -> date:
    return date(anio, mes, monthrange(anio, mes)[1])


def mes_cubierto(ultimo_serial: float, anio: int, mes: int, minutos_muestreo: int) -> bool:
    """¿Los datos llegan al último bloque del mes (último día 23:45 con muestreo de 15 min)?"""
    ultimo_bloque = datetime.combine(ultimo_dia(anio, mes) + timedelta(days=1), datetime.min.time()) - timedelta(
        minutes=minutos_muestreo
    )
    return ultimo_serial >= datetime_a_serial(ultimo_bloque)


def meses_candidatos(ultimo_dato: date) -> list[tuple[int, int]]:
    """Mes anterior y mes del último dato: un corte semanal cruza a lo más un cambio de mes."""
    previo = ultimo_dato.replace(day=1) - timedelta(days=1)
    return [(previo.year, previo.month), (ultimo_dato.year, ultimo_dato.month)]


# ------------------------------------------------------------------ cola de cierres
@dataclass
class ColaCierres:
    work: Path

    @property
    def ruta(self) -> Path:
        return Path(self.work) / ARCHIVO_COLA

    def entradas(self) -> list[dict]:
        return _leer(self.ruta, [])

    def pendientes(self) -> list[str]:
        return [e["mes"] for e in self.entradas() if e["estado"] == "pendiente"]

    def encolar(self, mes: str, corte: str) -> bool:
        """Agrega ``mes`` (``AAAA-MM``) si no estaba; devuelve si lo agregó."""
        entradas = self.entradas()
        if any(e["mes"] == mes for e in entradas):
            return False
        entradas.append({"mes": mes, "estado": "pendiente", "encolado_en": _ahora(), "corte_origen": corte})
        _escribir(self.ruta, sorted(entradas, key=lambda e: e["mes"]))
        return True

    def marcar_ejecutado(self, mes: str, id_corrida: str, estado_exclusiones: str) -> None:
        entradas = self.entradas()
        entrada = next((e for e in entradas if e["mes"] == mes), None)
        if entrada is None:
            entrada = {"mes": mes, "encolado_en": _ahora(), "corte_origen": None}
            entradas.append(entrada)
        entrada.update(
            estado="ejecutado", id_corrida=id_corrida, estado_exclusiones=estado_exclusiones, ejecutado_en=_ahora()
        )
        _escribir(self.ruta, sorted(entradas, key=lambda e: e["mes"]))


def encolar_cierres(
    cola: ColaCierres,
    ultimo_serial: float,
    ultimo_dato: date,
    minutos_muestreo: int,
    corte: str,
    desde: date,
    mes_cerrado: Callable[[int, int], bool] | None = None,
) -> list[str]:
    """Encola los meses cubiertos por los datos, desde ``desde`` y sin cierre oficial (R19.3).

    ``mes_cerrado(anio, mes)`` consulta la BD (cierre oficial o ``excel_manual``); sin BD no se filtra.
    """
    nuevos = []
    for anio, mes in meses_candidatos(ultimo_dato):
        if date(anio, mes, 1) < desde.replace(day=1) or not mes_cubierto(ultimo_serial, anio, mes, minutos_muestreo):
            continue
        if mes_cerrado is not None and mes_cerrado(anio, mes):
            continue
        if cola.encolar(f"{anio}-{mes:02d}", corte):
            nuevos.append(f"{anio}-{mes:02d}")
    return nuevos

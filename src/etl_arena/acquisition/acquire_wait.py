"""Etapa ``acquire-wait`` (tarea 4.5; R2.1–R2.3, R2.6).

Espera el export SCADA en ``data/inbox``, valida que exista, no esté vacío y no se esté
copiando todavía (tamaño estable), calcula su sha256 y lo **mueve** a
``data/processed/<corte>/`` junto a ``<archivo>.sha256``. Desde ahí el original es inmutable.
El contrato de columnas y fechas (4.6) llega con la muestra real (0.6/0.7).
"""

from __future__ import annotations

import fnmatch
import hashlib
import logging
import shutil
import time
from dataclasses import dataclass
from pathlib import Path

log = logging.getLogger("etl_arena.adquisicion")

PATRON_POR_DEFECTO = "raw_pcs_*.*"


class ErrorAdquisicion(RuntimeError):
    """No llegó el archivo a tiempo o no es utilizable (mensaje accionable, R2.2)."""


@dataclass(frozen=True)
class ArchivoAdquirido:
    ruta: Path  # ya dentro de data/processed/<corte>/
    sha256: str
    bytes: int
    corte: str


def sha256_de(ruta: Path) -> str:
    h = hashlib.sha256()
    with open(ruta, "rb") as fh:
        for bloque in iter(lambda: fh.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def _candidatos(inbox: Path, patron: str) -> list[Path]:
    return sorted(
        p
        for p in inbox.iterdir()
        if p.is_file() and fnmatch.fnmatch(p.name, patron) and not p.name.endswith((".tmp", ".part", ".crdownload"))
    )


def acquire_wait(
    inbox: str | Path,
    procesados: str | Path,
    corte: str,
    patron: str = PATRON_POR_DEFECTO,
    timeout_s: float = 3600,
    poll_s: float = 30,
    estabilidad_s: float = 5,
) -> ArchivoAdquirido:
    """Devuelve el archivo adquirido. Con más de un candidato falla (no adivina cuál usar)."""
    inbox, destino_dir = Path(inbox), Path(procesados) / corte
    if not inbox.is_dir():
        raise ErrorAdquisicion(f"no existe la carpeta de entrada {inbox}")
    limite = time.monotonic() + timeout_s
    while True:
        candidatos = _candidatos(inbox, patron)
        if len(candidatos) > 1:
            raise ErrorAdquisicion(
                f"hay {len(candidatos)} archivos {patron!r} en {inbox}: dejar solo el del corte "
                f"({', '.join(p.name for p in candidatos)})"
            )
        if candidatos:
            archivo = candidatos[0]
            tamano = archivo.stat().st_size
            time.sleep(estabilidad_s)  # TeamViewer puede seguir copiando
            if archivo.stat().st_size == tamano:
                break
            log.info("%s todavía se está copiando", archivo.name)
            continue
        if time.monotonic() >= limite:
            raise ErrorAdquisicion(
                f"no llegó ningún archivo {patron!r} a {inbox} en {timeout_s:.0f} s: exportar desde "
                "SCADA y copiarlo por TeamViewer (runbook §1)"
            )
        time.sleep(poll_s)

    if tamano == 0:
        raise ErrorAdquisicion(f"{archivo.name} está vacío")
    digest = sha256_de(archivo)
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / archivo.name
    if destino.exists():
        raise ErrorAdquisicion(f"{destino} ya existe: data/processed es inmutable (usar otro corte)")
    shutil.move(str(archivo), destino)
    (destino_dir / f"{archivo.name}.sha256").write_text(f"{digest}  {archivo.name}\n", encoding="ascii")
    log.info("adquirido %s (%d bytes, sha256 %s…)", destino, tamano, digest[:12])
    return ArchivoAdquirido(destino, digest, tamano, corte)

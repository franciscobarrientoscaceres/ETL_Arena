"""Copia de trabajo del libro (tarea 4.7, parte que no depende del contrato SCADA; ADR-06).

El libro base (maestro o el de la última corrida oficial) **nunca** se modifica: se copia a
``data/work/<corte>/`` y se deja un respaldo ``.bak`` de la copia antes de tocarla. La escritura
de las filas nuevas del export SCADA llega con 4.6/4.7 cuando exista la muestra real (0.6).
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

NOMBRE_LIBRO = "libro.xlsm"


def quitar_marca_de_internet(ruta: Path) -> bool:
    """Quita *Mark-of-the-Web* (``Zone.Identifier``) **de la copia**: Office bloquea las macros
    de archivos descargados aunque se abran por COM. Devuelve si había marca."""
    try:
        os.remove(f"{ruta}:Zone.Identifier")
        return True
    except OSError:
        return False


def copiar_libro_trabajo(libro_base: str | Path, dir_corte: str | Path, nombre: str = NOMBRE_LIBRO) -> Path:
    """``libro_base`` → ``dir_corte/nombre`` (+ ``nombre.bak``). Falla si la copia ya existe
    (cada corte trabaja sobre su propia copia; no se pisa trabajo previo)."""
    base, destino_dir = Path(libro_base).resolve(), Path(dir_corte).resolve()
    destino = destino_dir / nombre
    if destino == base:
        raise ValueError("la copia de trabajo no puede ser el libro base")
    if destino.exists():
        raise FileExistsError(f"ya existe {destino}; usar otro corte o borrarla a mano")
    destino_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(base, destino)
    shutil.copy2(destino, destino.with_suffix(destino.suffix + ".bak"))
    quitar_marca_de_internet(destino)
    return destino

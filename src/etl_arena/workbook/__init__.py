"""Libro de trabajo: copia, sesión COM, macros (solo PC local con Excel) y referencia por XML."""

from etl_arena.workbook.preparar import copiar_libro_trabajo, quitar_marca_de_internet
from etl_arena.workbook.referencia import extraer_referencia, sha256_archivo


def ejecutar_macros(*args, **kwargs):
    """Import diferido: requiere ``pywin32`` y Excel (extra ``com``)."""
    from etl_arena.workbook.macros import ejecutar_macros as _ejecutar

    return _ejecutar(*args, **kwargs)


__all__ = [
    "copiar_libro_trabajo",
    "ejecutar_macros",
    "extraer_referencia",
    "quitar_marca_de_internet",
    "sha256_archivo",
]

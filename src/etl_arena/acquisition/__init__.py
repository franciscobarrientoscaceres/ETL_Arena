"""Adquisición semanal del export SCADA (Fase S). El contrato y el lector llegan con la muestra real (0.6)."""

from etl_arena.acquisition.acquire_wait import ArchivoAdquirido, ErrorAdquisicion, acquire_wait, sha256_de

__all__ = ["ArchivoAdquirido", "ErrorAdquisicion", "acquire_wait", "sha256_de"]

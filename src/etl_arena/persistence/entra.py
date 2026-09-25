"""Autenticación Microsoft Entra ID para Azure SQL Database (ADR-10, 2026-09-24).

El token se obtiene con ``azure-identity`` (se instala con pip, sin permisos de administrador)
y se entrega al ODBC Driver 18 como ``SQL_COPT_SS_ACCESS_TOKEN``: no hay contraseñas en ``.env``.
El login abre el navegador una sola vez; el token queda en la caché cifrada de Windows (DPAPI)
y la cuenta elegida en ``~/.etl_arena/entra_auth_record.json``.
"""

from __future__ import annotations

import struct
from functools import cache
from pathlib import Path

SQL_COPT_SS_ACCESS_TOKEN = 1256
ALCANCE_AZURE_SQL = "https://database.windows.net/.default"
DIR_ESTADO = Path.home() / ".etl_arena"
REGISTRO_CUENTA = DIR_ESTADO / "entra_auth_record.json"


@cache
def _credencial():
    from azure.identity import (
        AuthenticationRecord,
        InteractiveBrowserCredential,
        TokenCachePersistenceOptions,
    )

    cache_persistente = TokenCachePersistenceOptions(name="etl_arena")
    if REGISTRO_CUENTA.exists():
        registro = AuthenticationRecord.deserialize(REGISTRO_CUENTA.read_text(encoding="utf-8"))
        return InteractiveBrowserCredential(cache_persistence_options=cache_persistente, authentication_record=registro)
    credencial = InteractiveBrowserCredential(cache_persistence_options=cache_persistente)
    registro = credencial.authenticate(scopes=[ALCANCE_AZURE_SQL])  # abre el navegador una vez
    DIR_ESTADO.mkdir(parents=True, exist_ok=True)
    REGISTRO_CUENTA.write_text(registro.serialize(), encoding="utf-8")
    return credencial


def token_odbc() -> bytes:
    """Token de acceso en el formato que espera el driver ODBC (longitud + UTF-16-LE)."""
    token = _credencial().get_token(ALCANCE_AZURE_SQL).token.encode("utf-16-le")
    return struct.pack(f"<I{len(token)}s", len(token), token)


def olvidar_cuenta() -> None:
    """Borra la cuenta recordada (el próximo uso vuelve a abrir el navegador)."""
    REGISTRO_CUENTA.unlink(missing_ok=True)
    _credencial.cache_clear()

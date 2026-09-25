"""Ambientes de base de datos del proyecto (sin dependencias; lo usan la conexión y los scripts).

Servidor Azure SQL: ``trina-etl.database.windows.net`` (Brazil South, oferta gratuita, Entra ID).

| Ambiente | ``--entorno`` / ``ETL_ARENA_ENTORNO`` | Base | Para qué |
|---|---|---|---|
| **PROD** | ``produccion`` (por defecto; alias ``prod``) | ``trina_etl`` | Resultados oficiales que ve Power BI |
| **TEST/QA** | ``prueba`` (alias ``qa``, ``test``) | ``trina_etl_prueba`` | Cargar, validar y repetir sin tocar PROD |

Las direcciones son fijas en el código para que cualquier PC apunte a las bases correctas sin configurar
nada; ``ETL_ARENA_DB_URL`` / ``ETL_ARENA_DB_URL_PRUEBA`` (``.env``) solo sirven para reemplazarlas.
"""

from __future__ import annotations

SERVIDOR_AZURE = "trina-etl.database.windows.net"
VARIABLE_URL = "ETL_ARENA_DB_URL"
VARIABLE_URL_PRUEBA = "ETL_ARENA_DB_URL_PRUEBA"
VARIABLE_ENTORNO = "ETL_ARENA_ENTORNO"

AMBIENTES = {
    "produccion": {"etiqueta": "PROD", "base": "trina_etl", "variable": VARIABLE_URL},
    "prueba": {"etiqueta": "TEST/QA", "base": "trina_etl_prueba", "variable": VARIABLE_URL_PRUEBA},
}
ALIAS_ENTORNO = {
    "produccion": "produccion",
    "producción": "produccion",
    "prod": "produccion",
    "production": "produccion",
    "prueba": "prueba",
    "qa": "prueba",
    "test": "prueba",
}
OPCIONES_ENTORNO = ("produccion", "prod", "prueba", "qa", "test")  # choices de argparse
AYUDA_ENTORNO = "produccion/prod = PROD (trina_etl) · prueba/qa/test = TEST/QA (trina_etl_prueba)"


class ErrorAmbiente(ValueError):
    """Nombre de ambiente desconocido."""


def normalizar_entorno(valor: str | None) -> str:
    """``produccion`` o ``prueba`` desde un nombre o alias; vacío = ``produccion``."""
    texto = (valor or "").strip().lower() or "produccion"
    if texto not in ALIAS_ENTORNO:
        raise ErrorAmbiente(
            f"{VARIABLE_ENTORNO}={valor!r}: se admite produccion (PROD, trina_etl) o prueba (TEST/QA, trina_etl_prueba)"
        )
    return ALIAS_ENTORNO[texto]


BASES_DE_AMBIENTE = frozenset(datos["base"] for datos in AMBIENTES.values())


def es_base_de_ambiente(base: str | None) -> bool:
    """``True`` para ``trina_etl`` (PROD) y ``trina_etl_prueba`` (TEST/QA): nunca se reinician ni se usan
    para los tests automáticos de integración."""
    return (base or "").strip().lower() in BASES_DE_AMBIENTE


def url_azure(base: str) -> str:
    """URL ``mssql+pyodbc`` de una base del servidor Azure del proyecto (ODBC Driver 18, cifrada)."""
    return (
        f"mssql+pyodbc://@{SERVIDOR_AZURE}:1433/{base}"
        "?driver=ODBC+Driver+18+for+SQL+Server&Encrypt=yes&TrustServerCertificate=no"
    )

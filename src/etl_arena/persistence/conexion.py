"""Conexión a SQL Server / Azure SQL Database (ADR-10).

* URL en ``ETL_ARENA_DB_URL`` (variable de entorno o ``.env``).
* ``ETL_ARENA_DB_AUTH=entra``: token de Microsoft Entra ID por conexión (``persistence.entra``);
  sin valor, manda lo que diga la URL (autenticación Windows o usuario SQL).
* Toda conexión física reintenta ante errores transitorios de Azure (p. ej. 40613 mientras una
  base serverless se reanuda, ~1 min).
* En Azure (``*.database.windows.net``) nunca se crean ni eliminan bases desde el código: un
  ``CREATE DATABASE`` crearía una base de pago; la base se crea en el portal (oferta gratuita).
"""

from __future__ import annotations

import logging
import os
import re
import time
from pathlib import Path

import sqlalchemy as sa
from sqlalchemy import event
from sqlalchemy.engine import URL, Engine, make_url

VARIABLE_URL = "ETL_ARENA_DB_URL"
VARIABLE_AUTH = "ETL_ARENA_DB_AUTH"
# Errores transitorios documentados para Azure SQL Database (reintentar la conexión).
CODIGOS_TRANSITORIOS = (
    "40613",
    "40197",
    "40501",
    "49918",
    "49919",
    "49920",
    "4060",
    "4221",
    "10928",
    "10929",
    "10053",
    "10054",
    "10060",
    "08S01",
    "HYT00",
)
REINTENTOS = 6
TIMEOUT_LOGIN_AZURE_S = 60
RECICLAJE_POOL_AZURE_S = 1800
ESPERA_S = 10.0
log = logging.getLogger(__name__)


class ErrorConexion(RuntimeError):
    """No hay URL configurada o la base no responde."""


def cargar_env(ruta: str | Path = ".env") -> dict[str, str]:
    """Lee ``CLAVE=valor`` de un ``.env`` (sin sobrescribir variables ya definidas). Sin dependencias."""
    ruta = Path(ruta)
    leidas: dict[str, str] = {}
    if not ruta.exists():
        return leidas
    for linea in ruta.read_text(encoding="utf-8").splitlines():
        linea = linea.strip()
        if not linea or linea.startswith("#") or "=" not in linea:
            continue
        clave, valor = linea.split("=", 1)
        leidas[clave.strip()] = valor.strip()
        os.environ.setdefault(clave.strip(), valor.strip())
    return leidas


def url_configurada(base_datos: str | None = None) -> URL:
    """URL de ``ETL_ARENA_DB_URL``; con ``base_datos`` reemplaza la base (p. ej. ``master`` o una de pruebas)."""
    if VARIABLE_URL not in os.environ:
        cargar_env()
    texto = os.environ.get(VARIABLE_URL)
    if not texto:
        raise ErrorConexion(f"falta {VARIABLE_URL} (ver .env.example)")
    url = make_url(texto)
    return url.set(database=base_datos) if base_datos else url


def es_azure(url: URL) -> bool:
    return (url.host or "").lower().endswith(".database.windows.net")


def _usa_entra() -> bool:
    if VARIABLE_AUTH not in os.environ:
        cargar_env()
    return os.environ.get(VARIABLE_AUTH, "").strip().lower() == "entra"


def _instalar_conexion(engine: Engine, entra: bool) -> Engine:
    """Cada conexión física: token Entra (si corresponde) + reintento ante errores transitorios."""

    @event.listens_for(engine, "do_connect")
    def _conectar(dialect, conn_rec, cargs, cparams):
        for intento in range(1, REINTENTOS + 1):
            if entra:
                from etl_arena.persistence.entra import SQL_COPT_SS_ACCESS_TOKEN, token_odbc

                cparams["attrs_before"] = {**cparams.get("attrs_before", {}), SQL_COPT_SS_ACCESS_TOKEN: token_odbc()}
                # sin usuario en la URL, SQLAlchemy agrega Trusted_Connection=Yes: incompatible con el token
                cargs = (re.sub(r"Trusted_Connection=Yes;?", "", cargs[0], flags=re.IGNORECASE), *cargs[1:])
            try:
                return dialect.dbapi.connect(*cargs, **cparams)
            except dialect.dbapi.Error as exc:
                transitorio = any(c in str(exc) for c in CODIGOS_TRANSITORIOS)
                if not transitorio or intento == REINTENTOS:
                    raise
                log.warning(
                    "conexión transitoria fallida (%s/%s), reintento en %.0f s: %s", intento, REINTENTOS, ESPERA_S, exc
                )
                time.sleep(ESPERA_S)

    return engine


def crear_engine(url: str | URL | None = None, **kwargs) -> Engine:
    """Engine ``mssql+pyodbc`` con ``fast_executemany`` (R11.5), token Entra opcional y reintentos."""
    url = _a_url(url)
    if es_azure(url):  # una base serverless pausada tarda ~1 min en reanudarse
        kwargs.setdefault("connect_args", {}).setdefault("timeout", TIMEOUT_LOGIN_AZURE_S)
        # Azure corta conexiones inactivas (~30 min) y el token Entra expira (~60-90 min)
        kwargs.setdefault("pool_recycle", RECICLAJE_POOL_AZURE_S)
    engine = sa.create_engine(url, fast_executemany=True, pool_pre_ping=True, **kwargs)
    return _instalar_conexion(engine, _usa_entra() and es_azure(url))  # Entra solo contra Azure


def _a_url(url: str | URL | None) -> URL:
    """Sin pasar por ``str(URL)``, que oculta la contraseña como ``***``."""
    if url is None:
        return url_configurada()
    return url if isinstance(url, URL) else make_url(url)


def _validar_nombre(nombre: str) -> str:
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,120}", nombre):
        raise ValueError(f"nombre de base inválido: {nombre!r}")
    return nombre


def crear_base_si_no_existe(nombre: str, url: str | URL | None = None) -> None:
    """``CREATE DATABASE`` desde ``master`` (autocommit)."""
    nombre = _validar_nombre(nombre)
    base = _a_url(url)
    if es_azure(base):
        raise ErrorConexion("en Azure SQL la base se crea desde el portal (oferta gratuita), no desde el código")
    engine = _instalar_conexion(sa.create_engine(base.set(database="master"), isolation_level="AUTOCOMMIT"), False)
    try:
        with engine.connect() as c:
            c.exec_driver_sql(f"IF DB_ID(N'{nombre}') IS NULL CREATE DATABASE [{nombre}]")
    finally:
        engine.dispose()


def eliminar_base(nombre: str, url: str | URL | None = None) -> None:
    """Solo para pruebas: cierra conexiones y elimina la base."""
    nombre = _validar_nombre(nombre)
    base = _a_url(url)
    if es_azure(base):
        raise ErrorConexion("no se eliminan bases de Azure SQL desde el código")
    engine = _instalar_conexion(sa.create_engine(base.set(database="master"), isolation_level="AUTOCOMMIT"), False)
    try:
        with engine.connect() as c:
            c.exec_driver_sql(
                f"IF DB_ID(N'{nombre}') IS NOT NULL BEGIN "
                f"ALTER DATABASE [{nombre}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [{nombre}]; END"
            )
    finally:
        engine.dispose()

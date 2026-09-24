"""Conexión a SQL Server (ADR-10): URL en ``ETL_ARENA_DB_URL`` (variable de entorno o ``.env``)."""

from __future__ import annotations

import os
import re
from pathlib import Path

import sqlalchemy as sa
from sqlalchemy.engine import URL, Engine, make_url

VARIABLE_URL = "ETL_ARENA_DB_URL"


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


def crear_engine(url: str | URL | None = None, **kwargs) -> Engine:
    """Engine ``mssql+pyodbc`` con ``fast_executemany`` (R11.5)."""
    return sa.create_engine(url or url_configurada(), fast_executemany=True, pool_pre_ping=True, **kwargs)


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
    engine = sa.create_engine(base.set(database="master"), isolation_level="AUTOCOMMIT")
    try:
        with engine.connect() as c:
            c.exec_driver_sql(f"IF DB_ID(N'{nombre}') IS NULL CREATE DATABASE [{nombre}]")
    finally:
        engine.dispose()


def eliminar_base(nombre: str, url: str | URL | None = None) -> None:
    """Solo para pruebas: cierra conexiones y elimina la base."""
    nombre = _validar_nombre(nombre)
    base = _a_url(url)
    engine = sa.create_engine(base.set(database="master"), isolation_level="AUTOCOMMIT")
    try:
        with engine.connect() as c:
            c.exec_driver_sql(
                f"IF DB_ID(N'{nombre}') IS NOT NULL BEGIN "
                f"ALTER DATABASE [{nombre}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [{nombre}]; END"
            )
    finally:
        engine.dispose()

"""Reglas de conexión sin base de datos (ADR-10: Azure SQL + Microsoft Entra ID)."""

import pytest
from sqlalchemy.engine import make_url

from etl_arena.persistence import conexion
from etl_arena.persistence.conexion import ErrorConexion, crear_base_si_no_existe, eliminar_base, es_azure
from etl_arena.persistence.esquema import reiniciar_esquema

AZURE = "mssql+pyodbc://@trina-etl.database.windows.net:1433/trina_etl?driver=ODBC+Driver+18+for+SQL+Server"
LOCAL = "mssql+pyodbc://@FRANCISCO-PC\\SQLSERVER2025DEV/ETL_Arena?driver=ODBC+Driver+18+for+SQL+Server"


def test_es_azure():
    assert es_azure(make_url(AZURE)) and not es_azure(make_url(LOCAL))


def test_nunca_crea_ni_borra_bases_en_azure():  # un CREATE DATABASE en Azure crearía una base de pago
    with pytest.raises(ErrorConexion):
        crear_base_si_no_existe("trina_etl", AZURE)
    with pytest.raises(ErrorConexion):
        eliminar_base("trina_etl", AZURE)


def test_nombre_de_base_invalido():
    with pytest.raises(ValueError):
        crear_base_si_no_existe("x]; DROP DATABASE y; --", LOCAL)


def test_entra_solo_contra_azure(monkeypatch):
    monkeypatch.setenv("ETL_ARENA_DB_AUTH", "entra")
    capturado = {}

    def falso(engine, entra):
        capturado["entra"] = entra
        return engine

    monkeypatch.setattr(conexion, "_instalar_conexion", falso)
    conexion.crear_engine(LOCAL)
    assert capturado["entra"] is False
    capturado.clear()
    engine = conexion.crear_engine(AZURE)
    assert capturado["entra"] is True
    assert engine.dialect.create_connect_args(engine.url)  # URL válida para pyodbc


def test_reiniciar_esquema_solo_en_bases_de_prueba():
    engine = conexion.sa.create_engine(AZURE)
    with pytest.raises(RuntimeError):
        reiniciar_esquema(engine)

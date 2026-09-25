"""Tests de integración SQL Server (tarea 2.5, marker ``sql``; R11, R12, R15, R19.7).

Base de pruebas:
* ``ETL_ARENA_TEST_DB_URL`` en Azure (``*.database.windows.net``): base fija de pruebas (oferta
  gratuita) cuyo esquema se reinicia al inicio de la sesión (``reiniciar_esquema`` exige "test" en el nombre).
* ``ETL_ARENA_TEST_DB_URL`` local, o sin definir con ``ETL_ARENA_DB_URL`` local: se crea
  ``ETL_Arena_test`` al inicio y se elimina al final.
* Sin base de pruebas y con ``ETL_ARENA_DB_URL`` en Azure: se omiten (nunca se toca la base real).
"""

from __future__ import annotations

import hashlib
import os
import time
import uuid
from datetime import date, datetime

import pyodbc
import pytest
from fabricas import config_prueba, crear_libro, fila_raw, serial_min

from etl_arena.model import DiaReferencia, EventoReferencia, FilaReconciliacion, ReferenciaExcel
from etl_arena.persistence import (
    MetadatosCorrida,
    RepositorioCorridas,
    aplicar_esquema,
    cargar_env,
    construir_paquete,
    crear_base_si_no_existe,
    crear_engine,
    eliminar_base,
    reiniciar_esquema,
    url_configurada,
)
from etl_arena.persistence.conexion import es_azure
from etl_arena.persistence.esquema import DIR_SQL, aplicar_script
from etl_arena.pipeline import calcular_libro

pytestmark = pytest.mark.sql
BASE = "ETL_Arena_test"


def _url_pruebas():
    from sqlalchemy.engine import make_url

    cargar_env()
    if os.environ.get("ETL_ARENA_TEST_DB_URL"):
        url = make_url(os.environ["ETL_ARENA_TEST_DB_URL"])
        return url if es_azure(url) else url.set(database=url.database or BASE)
    url = url_configurada()
    if es_azure(url):
        pytest.skip("ETL_ARENA_DB_URL apunta a Azure: definir ETL_ARENA_TEST_DB_URL con una base de pruebas")
    return url.set(database=BASE)


@pytest.fixture(scope="session")
def engine():
    try:
        url = _url_pruebas()
        local = not es_azure(url)
        if local:
            eliminar_base(url.database, url)
            crear_base_si_no_existe(url.database, url)
        eng = crear_engine(url)
        (aplicar_esquema if local else reiniciar_esquema)(eng)
    except pytest.skip.Exception:
        raise
    except Exception as exc:  # sin URL o instancia caída
        pytest.skip(f"SQL Server no disponible: {exc}")
    yield eng
    eng.dispose()
    if local:
        eliminar_base(url.database, url)


@pytest.fixture(scope="session")
def repo(engine):
    return RepositorioCorridas(engine)


def _consulta(engine, sql, *params):
    with engine.connect() as c:
        return c.exec_driver_sql(sql, params).fetchall() if params else c.exec_driver_sql(sql).fetchall()


@pytest.fixture(scope="session")
def libro_sintetico(tmp_path_factory):
    """4 filas × 2 PCS con un evento F55 en el PCS 1 y otro sin catálogo en el PCS 2."""
    s = [serial_min(15 * k) for k in range(1, 6)]
    filas = [
        fila_raw(s[0], [4, 4]),
        fila_raw(s[1], [3, 2], ["F55 X", 999.0]),
        fila_raw(s[2], [2, 2], ["F55 X", 999.0]),
        fila_raw(s[3], [4, 4]),
        fila_raw(s[4], [4, 4]),
    ]
    actividad = {r: [None, s[r - 2], 1, 1] for r in range(2, 7)}
    return crear_libro(tmp_path_factory.mktemp("libro") / "l.xlsx", filas, actividad=actividad)


def _corrida(repo, libro, *, oficial=True, exclusiones="sin_exclusiones", tipo="semanal", proyecto=1, guardar=True):
    """Corrida del 01-09-2026; los tests de vigencia usan otro ``proyecto`` para no mezclarse."""
    cfg = config_prueba(es_oficial=oficial, tipo_corrida=tipo, archivo_origen=libro.name, id_proyecto=proyecto)
    r = calcular_libro(libro, cfg, codigos_resumen=[("F55", "EXTERNAL")])
    meta = MetadatosCorrida("0" * 64, estado_exclusiones=exclusiones)
    repo.iniciar(cfg, meta)
    paquete = construir_paquete(r, meta, repo.cargar_catalogo())
    if guardar:
        repo.guardar_corrida(paquete)
        repo.finalizar(
            cfg.id_corrida,
            "success",
            {"anomalias": len(r.anomalias)},
            minutos_muestreo_derivado=r.disponibilidad.minutos_muestreo_derivado,
        )
    return cfg, r, paquete


# ------------------------------------------------------------------ DDL y seeds
def test_esquema_idempotente(engine):
    aplicar_esquema(engine)  # tercera aplicación (la fixture ya aplicó una)
    aplicar_esquema(engine)
    assert _consulta(engine, "SELECT COUNT(*) FROM dbo.proyecto")[0][0] == 4
    assert _consulta(engine, "SELECT COUNT(*) FROM dbo.tipo_detencion")[0][0] == 163
    assert _consulta(engine, "SELECT COUNT(*) FROM dbo.monthly_official_kpi WHERE Origen = 'excel_manual'")[0][0] == 2
    assert _consulta(engine, "SELECT TotalRacks FROM dbo.proyecto WHERE IdProyecto = 1")[0][0] == 2928


def test_sin_date_ni_datetime2(engine):  # convención del usuario (2026-09-24)
    tipos = _consulta(
        engine,
        "SELECT DISTINCT t.name FROM sys.columns c JOIN sys.types t ON t.user_type_id = c.user_type_id "
        "JOIN sys.tables tb ON tb.object_id = c.object_id WHERE t.name IN ('date', 'datetime2')",
    )
    assert tipos == []


def test_duplicados_d14_resueltos(engine):
    filas = _consulta(
        engine,
        "SELECT CodigoFalla, Significado FROM dbo.tipo_detencion "
        "WHERE CodigoFalla IN ('F228','F230','F231','F232') ORDER BY CodigoFalla",
    )
    assert [tuple(f) for f in filas] == [(c, "Crítico") for c in ("F228", "F230", "F231", "F232")]


# ------------------------------------------------------------------ corrida completa y trazabilidad
def test_corrida_completa(engine, repo, libro_sintetico):
    cfg, r, paquete = _corrida(repo, libro_sintetico)
    idc = cfg.id_corrida
    run = _consulta(
        engine,
        "SELECT Estado, EstadoExclusiones, VersionAlgoritmo, MinutosMuestreoDerivado, "
        "ISJSON(ResumenCalidad) FROM dbo.etl_run WHERE IdCorrida = ?",
        idc,
    )[0]
    assert tuple(run) == ("success", "sin_exclusiones", "availability-v1.1-exclusion-matrix", 15.0, 1)
    c14, suma = _consulta(
        engine,
        "SELECT a.BloquesRacksIndisponibles, (SELECT SUM(ImpactoRackPonderado) FROM "
        "dbo.availability_sample_result s WHERE s.IdCorrida = a.IdCorrida) "
        "FROM dbo.availability_run_result a WHERE a.IdCorrida = ?",
        idc,
    )[0]
    assert c14 == r.disponibilidad.bloques_racks_indisponibles == suma  # trazabilidad R15.1, sin pérdida de precisión
    for tabla in paquete.tablas:
        clave = "IdCorrida" if "IdCorrida" in tabla.columnas else None
        if clave:
            n = _consulta(engine, f"SELECT COUNT(*) FROM dbo.{tabla.nombre} WHERE IdCorrida = ?", idc)[0][0]
            assert n == len(tabla.filas), tabla.nombre
    sin_catalogo = _consulta(
        engine, "SELECT CodigoFalla, IdTipoDetencion FROM dbo.detencion WHERE IdCorrida = ? ORDER BY OrdenExcel", idc
    )
    assert [tuple(f) for f in sin_catalogo] == [("F55", 55), ("F999", None)]
    tipos = {t for (t,) in _consulta(engine, "SELECT Tipo FROM dbo.data_quality_issue WHERE IdCorrida = ?", idc)}
    assert "codigo_sin_catalogo" in tipos and "exclusion_matrix_ausente" in tipos


def test_append_only_dos_corridas_coexisten(engine, repo, libro_sintetico):
    a, _, _ = _corrida(repo, libro_sintetico)
    b, _, _ = _corrida(repo, libro_sintetico)
    n = _consulta(
        engine,
        "SELECT COUNT(DISTINCT IdCorrida) FROM dbo.fault_event WHERE IdCorrida IN (?, ?)",
        a.id_corrida,
        b.id_corrida,
    )[0][0]
    assert n == 2


def test_rollback_total_y_estado_failed(engine, repo, libro_sintetico):
    cfg, _, paquete = _corrida(repo, libro_sintetico, guardar=False)
    det = paquete.tabla("detencion")
    det.filas[-1] = det.filas[-1][:7] + (999_999,) + det.filas[-1][8:]  # FK inexistente en la penúltima tabla
    with pytest.raises(pyodbc.IntegrityError):
        repo.guardar_corrida(paquete)
    repo.finalizar(cfg.id_corrida, "failed", mensaje_error="FK inyectada")
    for tabla in paquete.tablas:
        if "IdCorrida" in tabla.columnas:
            assert (
                _consulta(engine, f"SELECT COUNT(*) FROM dbo.{tabla.nombre} WHERE IdCorrida = ?", cfg.id_corrida)[0][0]
                == 0
            ), tabla.nombre
    assert _consulta(engine, "SELECT Estado FROM dbo.etl_run WHERE IdCorrida = ?", cfg.id_corrida)[0][0] == "failed"


def test_no_se_agregan_datos_a_corrida_cerrada(repo, libro_sintetico):  # Checkpoint B-2
    cfg, _, paquete = _corrida(repo, libro_sintetico)  # ya finalizada en success
    with pytest.raises(RuntimeError):
        repo.guardar_corrida(paquete)


def test_finalizar_solo_desde_running(repo, libro_sintetico):
    cfg, _, _ = _corrida(repo, libro_sintetico)
    with pytest.raises(RuntimeError):
        repo.finalizar(cfg.id_corrida, "failed")


# ------------------------------------------------------------------ vigencia y revisión
def test_vigencia_prefiere_con_exclusiones(engine, repo, libro_sintetico):
    con, _, _ = _corrida(repo, libro_sintetico, exclusiones="con_exclusiones", tipo="cierre_mensual", proyecto=2)
    time.sleep(1)
    _corrida(repo, libro_sintetico, exclusiones="sin_exclusiones", proyecto=2)  # posterior, pero sin exclusiones
    _corrida(repo, libro_sintetico, tipo="golden", proyecto=2)  # nunca vigente
    vig = _consulta(
        engine,
        "SELECT IdCorrida, EtiquetaExclusiones FROM dbo.v_corrida_oficial_vigente "
        "WHERE IdProyecto = 2 AND Anio = 2026 AND Mes = 9",
    )
    assert [(str(i).lower(), e) for i, e in vig] == [(con.id_corrida, "Con Exclusiones")]


def test_revision_sobrevive_a_nueva_corrida(engine, repo, libro_sintetico):
    primera, r, _ = _corrida(repo, libro_sintetico, proyecto=3)
    evento = r.eventos.eventos()[0]
    repo.registrar_revision(3, evento.numero_pcs, evento.marca_tiempo_inicio, "revisado", "francisco", "ok")
    time.sleep(1)
    segunda, _, _ = _corrida(repo, libro_sintetico, proyecto=3)  # reproceso semanal del mismo mes
    fila = _consulta(
        engine,
        "SELECT IdCorrida, EstadoRevision, Observacion FROM dbo.v_detencion_vigente "
        "WHERE IdProyecto = 3 AND NumeroPCS = ? AND FechaInicio = ?",
        evento.numero_pcs,
        evento.marca_tiempo_inicio,
    )
    assert len(fila) == 1 and str(fila[0][0]).lower() == segunda.id_corrida != primera.id_corrida
    assert tuple(fila[0][1:]) == ("revisado", "ok")


# ------------------------------------------------------------------ seguridad (06_roles.sql)
def _como_usuario(engine, rol, sentencias):
    """Ejecuta ``sentencias`` como un usuario sin login miembro de ``rol``; devuelve los errores."""
    usuario = f"prueba_{rol}_{uuid.uuid4().hex[:6]}"
    conn = engine.raw_connection()
    errores = []
    try:
        conn.driver_connection.autocommit = True
        cur = conn.cursor()
        cur.execute(f"CREATE USER [{usuario}] WITHOUT LOGIN; ALTER ROLE {rol} ADD MEMBER [{usuario}];")
        for sql in sentencias:
            cur.execute(f"EXECUTE AS USER = '{usuario}'")
            try:
                cur.execute(sql)
                errores.append(None)
            except pyodbc.Error as exc:
                errores.append(str(exc))
            finally:
                cur.execute("REVERT")
    finally:
        conn.driver_connection.autocommit = False
        conn.close()
    return errores


def test_bi_reader_solo_vistas(engine):
    ok, tabla = _como_usuario(
        engine,
        "bi_reader",
        ["SELECT TOP 1 * FROM dbo.v_kpi_vigente", "SELECT TOP 1 * FROM dbo.availability_run_result"],
    )
    assert ok is None and tabla is not None and "229" in tabla  # 229 = permiso denegado


def test_etl_writer_append_only(engine, repo, libro_sintetico):
    cfg, _, _ = _corrida(repo, libro_sintetico, guardar=False)
    res = _como_usuario(
        engine,
        "etl_writer",
        [
            f"UPDATE dbo.etl_run SET Estado = 'failed' WHERE IdCorrida = '{cfg.id_corrida}'",
            "DELETE FROM dbo.fault_event WHERE 1 = 0",
            "UPDATE dbo.availability_run_result SET BloquesMuestreo = 0 WHERE 1 = 0",
            # Checkpoint B-1: maestros y revisión humana fuera del alcance de la carga
            "INSERT INTO dbo.proyecto (IdProyecto, Nombre, Estado) VALUES (99, N'x', N'por_implementar')",
            "INSERT INTO dbo.tipo_detencion (IdTipoDetencion, CodigoFalla, DescripcionFallaPE, CodigoDescripcion) "
            "VALUES (9999, N'F9999', N'x', N'x')",
            "INSERT INTO dbo.detencion_revision (IdProyecto, NumeroPCS, FechaInicio, EstadoRevision, RevisadoPor) "
            "VALUES (1, 1, GETDATE(), N'revisado', N'carga')",
        ],
    )
    assert res[0] is None and all(e is not None for e in res[1:]), res


# ------------------------------------------------------------------ referencia, reconciliación, cargas, auditoría
def test_referencia_y_reconciliacion(engine, repo, libro_sintetico):
    cfg, _, _ = _corrida(repo, libro_sintetico)
    ref = ReferenciaExcel(
        "2026-09-01",
        "libro.xlsm",
        "a" * 64,
        datetime(2026, 9, 24, 12),
        date(2026, 9, 1),
        date(2026, 9, 1),
        date(2026, 9, 1),
        date(2026, 9, 1),
        c21="No",
        c31="Yes",
        l14="Yes",
        c12=4,
        c14=36.0,
        tabla=[(4, serial_min(30), 1, 1.0)],
        eventos=[EventoReferencia(1, 1.0, 46266.0, 46266.1, 0.5, "F55", "F55 X", 1.5, 9.0)],
        diario=[DiaReferencia(1, date(2026, 9, 1), 36.0, 36.0, 0.99, 0.0)],
    )
    idr = repo.guardar_referencia_excel(ref)
    n = repo.guardar_reconciliacion(
        cfg.id_corrida, idr, [FilaReconciliacion(3, "C14", None, "36", "36", 0.0, 1e-6, True)]
    )
    assert n == 1
    assert (
        _consulta(engine, "SELECT COUNT(*) FROM dbo.excel_reference_fault_event WHERE IdReferencia = ?", idr)[0][0] == 1
    )


def test_carga_mensual_exclusion(repo):
    assert not repo.exclusiones_cargadas(1, 2026, 8)
    assert repo.registrar_carga_exclusion(1, 2026, 8, "exclusion_agosto.xlsx", "b" * 64) > 0
    assert repo.exclusiones_cargadas(1, 2026, 8)


def test_consultas_de_auditoria_compilan(engine):
    assert aplicar_script(engine, DIR_SQL / "07_audit_queries.sql") == 1


# ------------------------------------------------------------------ rendimiento (tarea 2.6)
@pytest.mark.golden
def test_rendimiento_corrida_real(engine, repo, ruta_libro_real):
    from etl_arena.config import construir_config, parametros_desde_celdas
    from etl_arena.ingestion import leer_celdas_parametros

    cfg = construir_config(
        **parametros_desde_celdas(leer_celdas_parametros(ruta_libro_real)).valores,
        archivo_origen=ruta_libro_real.name,
        tipo_corrida="golden",
    )
    r = calcular_libro(ruta_libro_real, cfg)
    meta = MetadatosCorrida(hashlib.sha256(ruta_libro_real.read_bytes()).hexdigest())
    repo.iniciar(cfg, meta)
    resultado = repo.guardar_corrida(construir_paquete(r, meta, repo.cargar_catalogo()))
    repo.finalizar(cfg.id_corrida, "success")
    assert resultado.segundos < 180, resultado  # objetivo 2.6: < 3 min
    c14 = _consulta(
        engine, "SELECT BloquesRacksIndisponibles FROM dbo.availability_run_result WHERE IdCorrida = ?", cfg.id_corrida
    )[0][0]
    assert c14 == 104134.2920000001

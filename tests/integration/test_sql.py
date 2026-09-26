"""Tests de integración SQL Server (tareas 2.5 y 6.6, marker ``sql``; R11 rev. 3, R12, R15, R19.7; ADR-12).

Base de pruebas:
* ``ETL_ARENA_TEST_DB_URL`` en Azure (``*.database.windows.net``): base fija de tests (oferta
  gratuita) cuyo esquema se reinicia al inicio de la sesión (``reiniciar_esquema`` exige "test" en el nombre).
  **Nunca** ``trina_etl`` (PROD) ni ``trina_etl_prueba`` (TEST/QA): con ellas los tests se omiten.
* ``ETL_ARENA_TEST_DB_URL`` local, o sin definir con ``ETL_ARENA_DB_URL`` local: se crea
  ``ETL_Arena_test`` al inicio y se elimina al final.
* Sin base de pruebas y con ``ETL_ARENA_DB_URL`` en Azure: se omiten (nunca se toca la base real).

Cada test de estado vigente usa su propio ``IdProyecto`` (1–4) y lo deja vacío al empezar (``limpiar``).
"""

from __future__ import annotations

import dataclasses
import hashlib
import os
import uuid
from datetime import date, datetime

import pyodbc
import pytest
from fabricas import config_prueba, crear_libro, fila_raw, serial_min

from etl_arena.ambientes import es_base_de_ambiente
from etl_arena.model import FilaReconciliacion, ReferenciaExcel
from etl_arena.persistence import (
    ErrorPublicacion,
    MetadatosCorrida,
    OpcionesPublicacion,
    RepositorioCorridas,
    aplicar_esquema,
    cargar_env,
    construir_paquete_mes,
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
ESTADO = (
    "muestra_pcs",
    "muestra_planta",
    "detencion",
    "disponibilidad_diaria",
    "disponibilidad_mensual",
    "calidad_dato",
)


def _url_pruebas():
    from sqlalchemy.engine import make_url

    cargar_env()
    if os.environ.get("ETL_ARENA_TEST_DB_URL"):
        url = make_url(os.environ["ETL_ARENA_TEST_DB_URL"])
        if es_base_de_ambiente(url.database):  # nunca vaciar PROD ni TEST/QA
            pytest.skip(
                f"ETL_ARENA_TEST_DB_URL apunta a {url.database} (un ambiente PROD o TEST/QA): usar una base propia "
                "con 'test' en el nombre (p. ej. la instancia local ETL_Arena_test)"
            )
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


def limpiar(engine, proyecto: int) -> None:
    """Deja vacío el estado vigente del proyecto (los tests comparten la base de la sesión)."""
    with engine.begin() as c:
        for tabla in ESTADO:
            c.exec_driver_sql(f"DELETE FROM dbo.{tabla} WHERE IdProyecto = ?", (proyecto,))


def _libro(tmp_path_factory, nombre, modulos_pcs2=(2, 2), ultimo=5, repetido=False):
    """Septiembre 00:15–01:15 × 2 PCS: evento F55 en el PCS 1 y otro sin catálogo en el PCS 2."""
    s = [serial_min(15 * k) for k in range(1, 6)]
    filas = [
        fila_raw(s[0], [4, 4]),
        fila_raw(s[1], [3, modulos_pcs2[0]], ["F55 X", 999.0]),
        fila_raw(s[2], [2, modulos_pcs2[1]], ["F55 X", 999.0]),
        fila_raw(s[3], [4, 4]),
        fila_raw(s[4], [4, 4]),
    ][:ultimo]
    if repetido:  # cambio de hora (F-32): el mismo timestamp dos veces
        filas.insert(2, fila_raw(s[1], [4, 4]))
    actividad = {r: [None, f[0], 1, 1] for r, f in enumerate(filas, start=2)}
    return crear_libro(tmp_path_factory.mktemp("libro") / nombre, filas, actividad=actividad)


@pytest.fixture(scope="session")
def libro_sintetico(tmp_path_factory):
    return _libro(tmp_path_factory, "l.xlsx")


def publicar(
    repo,
    libro,
    *,
    proyecto,
    exclusiones="sin_exclusiones",
    tipo="semanal",
    fin=date(2026, 9, 1),
    inicio=date(2026, 9, 1),
    opciones=None,
    alterar=None,
):
    """Calcula el libro, registra la corrida y publica su mes; devuelve (cfg, resultado, ResumenPublicacion)."""
    cfg = config_prueba(
        inicio_periodo=inicio, fin_periodo=fin, tipo_corrida=tipo, archivo_origen=libro.name, id_proyecto=proyecto
    )
    r = calcular_libro(libro, cfg, codigos_resumen=[("F55", "EXTERNAL")])
    meta = MetadatosCorrida("0" * 64, estado_exclusiones=exclusiones)
    num = repo.iniciar(cfg, meta)
    paquete = construir_paquete_mes(r, meta, repo.cargar_catalogo())
    if alterar:
        alterar(paquete)
    try:
        resumen = repo.publicar_mes(cfg.id_corrida, num, paquete, opciones, libro.name, "0" * 64)
    except Exception:
        repo.finalizar(cfg.id_corrida, "failed", mensaje_error="publicación rechazada")
        raise
    repo.finalizar(cfg.id_corrida, "success", {}, publicada=True, meses_publicados=[resumen.mes_texto])
    publicar.ultimo_num = num
    return cfg, r, resumen


def _conteos(engine, proyecto):
    return {t: _consulta(engine, f"SELECT COUNT(*) FROM dbo.{t} WHERE IdProyecto = ?", proyecto)[0][0] for t in ESTADO}


# ------------------------------------------------------------------ DDL y seeds
def test_esquema_idempotente(engine):
    aplicar_esquema(engine)  # tercera aplicación (la fixture ya aplicó una)
    aplicar_esquema(engine)
    assert _consulta(engine, "SELECT COUNT(*) FROM dbo.proyecto")[0][0] == 4
    assert _consulta(engine, "SELECT COUNT(*) FROM dbo.tipo_detencion")[0][0] == 163
    manual = _consulta(
        engine,
        "SELECT Mes, TotalRacks FROM dbo.disponibilidad_mensual WHERE IdProyecto = 1 AND Origen = 'excel_manual' "
        "ORDER BY Mes",
    )
    assert [tuple(f) for f in manual] == [(7, 2928), (8, 2928)]
    tablas_v1 = _consulta(
        engine, "SELECT name FROM sys.tables WHERE name IN ('fault_event', 'raw_pcs_sample', 'availability_run_result')"
    )
    assert tablas_v1 == []


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


def test_migracion_v2_se_niega_con_corridas(engine, repo):
    """02a solo borra tablas v1 si no hay corridas registradas (nunca pierde historial)."""
    repo.iniciar(config_prueba(archivo_origen="x.xlsm"), MetadatosCorrida("0" * 64))
    with engine.begin() as c:
        c.exec_driver_sql("CREATE TABLE dbo.fault_event (x INT)")  # simula una base v1
    try:
        with pytest.raises(Exception, match="50001"):
            aplicar_script(engine, DIR_SQL / "02a_migracion_v2.sql")
        assert _consulta(engine, "SELECT COUNT(*) FROM dbo.etl_run")[0][0] > 0
    finally:
        with engine.begin() as c:
            c.exec_driver_sql("DROP TABLE IF EXISTS dbo.fault_event")


# ------------------------------------------------------------------ publicación del mes (ADR-12)
def test_publicacion_completa_y_trazable(engine, repo, libro_sintetico):  # propiedad 17
    limpiar(engine, 1)
    cfg, r, res = publicar(repo, libro_sintetico, proyecto=1)
    num = publicar.ultimo_num
    run = _consulta(
        engine,
        "SELECT Estado, EstadoExclusiones, VersionAlgoritmo, Publicada, MesesPublicados FROM dbo.etl_run "
        "WHERE IdCorrida = ?",
        cfg.id_corrida,
    )[0]
    assert tuple(run) == ("success", "sin_exclusiones", "availability-v1.1-exclusion-matrix", True, "2026-09")
    c14, suma, n_num = _consulta(
        engine,
        "SELECT m.BloquesRacksIndisponibles, (SELECT SUM(ImpactoRackPonderado) FROM dbo.muestra_pcs p "
        "WHERE p.IdProyecto = m.IdProyecto AND p.Anio = m.Anio AND p.Mes = m.Mes), m.NumCorrida "
        "FROM dbo.disponibilidad_mensual m WHERE IdProyecto = 1 AND Anio = 2026 AND Mes = 9",
    )[0]
    assert c14 == r.disponibilidad.bloques_racks_indisponibles == suma and n_num == num  # R15.1, sin perder precisión
    assert res.filas_insertadas["muestra_pcs"] == 5 * 2 and res.detenciones == {
        "nuevas": 2,
        "actualizadas": 0,
        "eliminadas": 0,
    }
    det = _consulta(
        engine,
        "SELECT CodigoFalla, IdTipoDetencion, DuracionSegundos, DuracionHoras, NumCorrida FROM dbo.detencion "
        "WHERE IdProyecto = 1 ORDER BY OrdenExcel",
    )
    assert [(c, t) for c, t, *_ in det] == [("F55", 55), ("F999", None)]
    assert all(s == round(h * 3600) and n == num for _, _, s, h, n in det)
    tipos = {t for (t,) in _consulta(engine, "SELECT Tipo FROM dbo.calidad_dato WHERE IdProyecto = 1")}
    assert "codigo_sin_catalogo" in tipos and "exclusion_matrix_ausente" in tipos
    assert _consulta(engine, "SELECT COUNT(*) FROM dbo.v_ejecuciones WHERE NumCorrida = ?", num)[0][0] == 1


def test_recarga_idempotente_sin_duplicados(engine, repo, libro_sintetico):  # propiedades 12 y 13
    limpiar(engine, 2)
    publicar(repo, libro_sintetico, proyecto=2)
    antes = _conteos(engine, 2)
    num_a = publicar.ultimo_num
    _, _, res = publicar(repo, libro_sintetico, proyecto=2)  # el lunes siguiente recarga el mismo mes
    assert publicar.ultimo_num == num_a + 1  # NumCorrida correlativo sin saltos (IDENTITY_CACHE = OFF)
    assert _conteos(engine, 2) == antes and res.correcciones == 0
    assert res.detenciones == {"nuevas": 0, "actualizadas": 2, "eliminadas": 0}
    assert res.filas_borradas["muestra_pcs"] == res.filas_insertadas["muestra_pcs"] == 5 * 2
    numeros = _consulta(engine, "SELECT DISTINCT NumCorrida FROM dbo.muestra_pcs WHERE IdProyecto = 2")
    assert [n for (n,) in numeros] == [num_a + 1]  # todo el mes quedó con la última carga


def test_recarga_con_datos_corregidos(engine, repo, libro_sintetico, tmp_path_factory):  # propiedades 14 y 15
    limpiar(engine, 3)
    _, r, _ = publicar(repo, libro_sintetico, proyecto=3)
    evento = r.eventos.eventos()[0]
    repo.registrar_revision(3, evento.numero_pcs, evento.marca_tiempo_inicio, "revisado", "francisco", "ok")
    ids = dict(_consulta(engine, "SELECT CodigoFalla, IdDetencion FROM dbo.detencion WHERE IdProyecto = 3"))

    corregido = _libro(tmp_path_factory, "corregido.xlsx", modulos_pcs2=(4, 4))  # SCADA corrigió el PCS 2
    cfg, _, res = publicar(repo, corregido, proyecto=3)
    assert res.correcciones == 2 and res.detenciones == {"nuevas": 0, "actualizadas": 1, "eliminadas": 1}
    cambios = _consulta(
        engine,
        "SELECT TipoCorreccion, Hoja, NumeroPCS, Campo, ValorAnterior, ValorNuevo FROM dbo.v_correccion_dato "
        "WHERE IdCorrida = ? ORDER BY NumeroFilaOrigen",
        cfg.id_corrida,
    )
    assert [tuple(c) for c in cambios] == [
        ("recarga", "RawData-PCS", 2, "MODULES", "2.0", "4.0"),
        ("recarga", "RawData-PCS", 2, "MODULES", "2.0", "4.0"),
    ]
    vig = _consulta(
        engine, "SELECT IdDetencion, CodigoFalla, EstadoRevision, Observacion FROM dbo.v_detencion WHERE IdProyecto = 3"
    )
    assert [tuple(v) for v in vig] == [(ids["F55"], "F55", "revisado", "ok")]  # IdDetencion estable (D-25)


def test_rollback_deja_intacto_el_estado_anterior(engine, repo, libro_sintetico):  # propiedad 16
    limpiar(engine, 4)
    publicar(repo, libro_sintetico, proyecto=4)
    antes = _conteos(engine, 4)
    num_vigente = publicar.ultimo_num

    def fk_invalida(paquete):
        det = paquete.detenciones
        k = det.columnas.index("IdTipoDetencion")
        det.filas[-1] = (*det.filas[-1][:k], 999_999, *det.filas[-1][k + 1 :])

    with pytest.raises(pyodbc.IntegrityError):
        publicar(repo, libro_sintetico, proyecto=4, alterar=fk_invalida)
    assert _conteos(engine, 4) == antes
    assert (
        _consulta(engine, "SELECT DISTINCT NumCorrida FROM dbo.muestra_pcs WHERE IdProyecto = 4")[0][0] == num_vigente
    )
    assert _consulta(engine, "SELECT TOP 1 Estado FROM dbo.etl_run ORDER BY NumCorrida DESC")[0][0] == "failed"


def test_protecciones(engine, repo, libro_sintetico, tmp_path_factory):  # D-21, D-23
    limpiar(engine, 4)
    publicar(repo, libro_sintetico, proyecto=4, exclusiones="con_exclusiones", tipo="cierre_mensual")
    with pytest.raises(ErrorPublicacion, match="forzar-sin-exclusiones"):
        publicar(repo, libro_sintetico, proyecto=4)
    publicar(repo, libro_sintetico, proyecto=4, opciones=OpcionesPublicacion(forzar_sin_exclusiones=True))

    corto = _libro(tmp_path_factory, "corto.xlsx", ultimo=3)  # llega hasta 00:45 en vez de 01:15
    with pytest.raises(ErrorPublicacion, match="permitir-recorte"):
        publicar(repo, corto, proyecto=4)
    publicar(repo, corto, proyecto=4, opciones=OpcionesPublicacion(permitir_recorte=True))
    assert _conteos(engine, 4)["muestra_pcs"] == 3 * 2

    with engine.begin() as c:
        c.exec_driver_sql("UPDATE dbo.disponibilidad_mensual SET Origen = 'excel_manual' WHERE IdProyecto = 4")
    with pytest.raises(ErrorPublicacion, match="reemplazar-manual"):
        publicar(repo, corto, proyecto=4)
    publicar(repo, corto, proyecto=4, opciones=OpcionesPublicacion(reemplazar_manual=True))
    assert repo.estado_mes(4, 2026, 9)["Origen"] == "corrida"


def test_meses_independientes(engine, repo, tmp_path_factory):  # propiedad multi-mes (D-20)
    limpiar(engine, 2)
    agosto = [serial_min(15 * k, date(2026, 8, 31)) for k in (93, 94, 95)]  # 23:15 … 23:45
    septiembre = [serial_min(15 * k) for k in (1, 2)]
    filas = [fila_raw(s, [3, 4], ["F55 X", None]) for s in agosto] + [fila_raw(s, [4, 4]) for s in septiembre]
    act = {r: [None, f[0], 1, 1] for r, f in enumerate(filas, start=2)}
    libro = crear_libro(tmp_path_factory.mktemp("multi") / "m.xlsx", filas, actividad=act)
    publicar(repo, libro, proyecto=2, inicio=date(2026, 8, 1), fin=date(2026, 8, 31))
    publicar(repo, libro, proyecto=2, inicio=date(2026, 9, 1), fin=date(2026, 9, 1))
    por_mes = _consulta(
        engine, "SELECT Mes, COUNT(*) FROM dbo.muestra_pcs WHERE IdProyecto = 2 GROUP BY Mes ORDER BY Mes"
    )
    assert [tuple(f) for f in por_mes] == [(8, 3 * 2), (9, 2 * 2)]
    publicar(repo, libro, proyecto=2, inicio=date(2026, 9, 1), fin=date(2026, 9, 1))  # recargar sept no toca agosto
    assert _consulta(engine, "SELECT COUNT(*) FROM dbo.muestra_pcs WHERE IdProyecto = 2 AND Mes = 8")[0][0] == 6
    assert _consulta(engine, "SELECT COUNT(*) FROM dbo.detencion WHERE IdProyecto = 2 AND Mes = 8")[0][0] == 1


def test_timestamp_repetido_usa_ocurrencia(engine, repo, tmp_path_factory):  # F-32
    limpiar(engine, 3)
    libro = _libro(tmp_path_factory, "dst.xlsx", repetido=True)
    publicar(repo, libro, proyecto=3)
    publicar(repo, libro, proyecto=3)  # y se recarga sin chocar con la llave
    occ = _consulta(
        engine,
        "SELECT Ocurrencia, COUNT(*) FROM dbo.muestra_pcs WHERE IdProyecto = 3 GROUP BY Ocurrencia ORDER BY Ocurrencia",
    )
    assert [tuple(f) for f in occ] == [(1, 10), (2, 2)]


def test_anual_con_meses_manuales(engine, repo, libro_sintetico):
    limpiar(engine, 1)
    aplicar_script(engine, DIR_SQL / "06_seed_annual_manual.sql")  # limpiar borra también jul/ago
    _, r, _ = publicar(repo, libro_sintetico, proyecto=1)
    c12 = float(r.disponibilidad.bloques_muestreo)
    anual = _consulta(
        engine,
        "SELECT Mes, BloquesMuestreoAcumulados FROM dbo.v_disponibilidad_anual WHERE IdProyecto = 1 AND Anio = 2026 "
        "ORDER BY Mes",
    )
    assert [tuple(a) for a in anual] == [(7, 2976.0), (8, 5952.0), (9, 5952.0 + c12)]
    assert sorted(k.mes for k in repo.kpi_mensuales_vigentes(1, 2026)) == [7, 8, 9]


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


def test_bi_reader_lee_estado_y_vistas(engine):
    res = _como_usuario(
        engine,
        "bi_reader",
        [
            "SELECT TOP 1 * FROM dbo.disponibilidad_mensual",
            "SELECT TOP 1 * FROM dbo.v_detencion",
            "SELECT TOP 1 * FROM dbo.etl_run",
            "DELETE FROM dbo.detencion WHERE 1 = 0",
        ],
    )
    assert res[0] is None and res[1] is None and "229" in res[2] and "229" in res[3]  # 229 = permiso denegado


def test_etl_writer_reemplaza_estado_pero_no_registro(engine, repo, libro_sintetico):
    cfg = config_prueba(archivo_origen=libro_sintetico.name)
    repo.iniciar(cfg, MetadatosCorrida("0" * 64))
    res = _como_usuario(
        engine,
        "etl_writer",
        [
            f"UPDATE dbo.etl_run SET Estado = 'failed' WHERE IdCorrida = '{cfg.id_corrida}'",
            "DELETE FROM dbo.muestra_pcs WHERE 1 = 0",
            "UPDATE dbo.disponibilidad_mensual SET BloquesMuestreo = 0 WHERE 1 = 0",
            # historial y maestros fuera del alcance de la carga
            "DELETE FROM dbo.etl_run WHERE 1 = 0",
            "DELETE FROM dbo.correccion_dato WHERE 1 = 0",
            "UPDATE dbo.reconciliation_result SET Aprobado = 0 WHERE 1 = 0",
            "INSERT INTO dbo.proyecto (IdProyecto, Nombre, Estado) VALUES (99, N'x', N'por_implementar')",
            "INSERT INTO dbo.tipo_detencion (IdTipoDetencion, CodigoFalla, DescripcionFallaPE, CodigoDescripcion) "
            "VALUES (9999, N'F9999', N'x', N'x')",
            "INSERT INTO dbo.detencion_revision (IdProyecto, NumeroPCS, FechaInicio, EstadoRevision, RevisadoPor) "
            "VALUES (1, 1, GETDATE(), N'revisado', N'carga')",
        ],
    )
    assert res[:3] == [None, None, None] and all(e is not None for e in res[3:]), res


# ------------------------------------------------------------------ referencia, reconciliación, cargas, auditoría
def test_referencia_y_reconciliacion(engine, repo, libro_sintetico):
    limpiar(engine, 1)
    cfg, _, _ = publicar(repo, libro_sintetico, proyecto=1)
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
    )
    idr = repo.guardar_referencia_excel(ref)
    n = repo.guardar_reconciliacion(
        cfg.id_corrida, idr, [FilaReconciliacion(3, "C14", None, "36", "36", 0.0, 1e-6, True)]
    )
    assert n == 1
    assert _consulta(engine, "SELECT C14 FROM dbo.excel_reference_run WHERE IdReferencia = ?", idr)[0][0] == 36.0


def test_carga_mensual_exclusion(engine, repo, libro_sintetico):
    assert not repo.exclusiones_cargadas(1, 2026, 8)
    id_carga = repo.registrar_carga_exclusion(1, 2026, 8, "exclusion_agosto.xlsx", "b" * 64)
    assert id_carga > 0 and repo.exclusiones_cargadas(1, 2026, 8)
    limpiar(engine, 1)
    cfg, _, _ = publicar(repo, libro_sintetico, proyecto=1, exclusiones="con_exclusiones", tipo="cierre_mensual")
    assert repo.vincular_carga_con_cierre(id_carga, cfg.id_corrida)
    assert not repo.vincular_carga_con_cierre(id_carga, cfg.id_corrida)  # una sola vez: no se pisa
    vinculada = _consulta(engine, "SELECT IdCorridaCierre FROM dbo.exclusion_matrix_carga WHERE IdCarga = ?", id_carga)
    assert str(vinculada[0][0]).lower() == cfg.id_corrida


def test_mes_cerrado_para_la_cola_de_cierres(engine, repo, libro_sintetico, tmp_path_factory):  # R19.3
    limpiar(engine, 1)
    aplicar_script(engine, DIR_SQL / "06_seed_annual_manual.sql")
    assert repo.mes_cerrado(1, 2026, 7) and repo.mes_cerrado(1, 2026, 8)  # excel_manual (jul/ago)
    limpiar(engine, 2)
    publicar(repo, libro_sintetico, proyecto=2, tipo="cierre_mensual", fin=date(2026, 9, 30))
    assert not repo.mes_cerrado(2, 2026, 9)  # los datos no llegan al 30-sep 23:45
    ultimo = serial_min(30 * 24 * 60 - 15)  # 30-sep 23:45
    filas = [fila_raw(serial_min(15), [4, 4]), fila_raw(ultimo, [4, 4])]
    completo = crear_libro(
        tmp_path_factory.mktemp("completo") / "c.xlsx",
        filas,
        actividad={2: [None, filas[0][0], 1, 1], 3: [None, ultimo, 1, 1]},
    )
    publicar(repo, completo, proyecto=2, tipo="cierre_mensual", fin=date(2026, 9, 30))
    assert repo.mes_cerrado(2, 2026, 9) and not repo.mes_cerrado(2, 2026, 10)


def test_correcciones_de_la_matriz_quedan_ligadas_a_la_corrida(engine, repo, libro_sintetico):  # 4.12, D-13
    limpiar(engine, 1)
    cfg, _, _ = publicar(repo, libro_sintetico, proyecto=1, exclusiones="con_exclusiones", tipo="cierre_mensual")
    ts = datetime(2026, 9, 1, 0, 15)
    filas = [
        ("exclusion_matrix", "Exclusion_Matrix", 2, 46266.0104166667, ts, 1, "PCS01", 0.0, 1.0, "em.xlsx", "c" * 64),
        (
            "exclusion_matrix",
            "Exclusion_Matrix",
            2,
            46266.0104166667,
            ts,
            None,
            "Comments",
            None,
            "x" * 300,
            "em.xlsx",
            "c" * 64,
        ),
    ]
    assert repo.guardar_correcciones(cfg.id_corrida, filas) == 2
    vista = _consulta(
        engine,
        "SELECT Campo, ValorAnterior, ValorNuevo, NumeroPCS, TipoCorrida FROM dbo.v_correccion_dato "
        "WHERE IdCorrida = ? ORDER BY IdCorreccion",
        cfg.id_corrida,
    )
    assert [tuple(f) for f in vista] == [
        ("PCS01", "0.0", "1.0", 1, "cierre_mensual"),
        ("Comments", None, "x" * 255, None, "cierre_mensual"),
    ]


def test_consultas_de_auditoria_compilan(engine):
    assert aplicar_script(engine, DIR_SQL / "07_audit_queries.sql") == 1


# ------------------------------------------------------------------ rendimiento (tarea 2.6)
@pytest.mark.golden
def test_rendimiento_corrida_real(engine, repo, ruta_libro_real):
    from etl_arena.config import construir_config, parametros_desde_celdas
    from etl_arena.ingestion import leer_celdas_parametros

    limpiar(engine, 1)
    cfg = construir_config(
        **parametros_desde_celdas(leer_celdas_parametros(ruta_libro_real)).valores,
        archivo_origen=ruta_libro_real.name,
    )
    r = calcular_libro(ruta_libro_real, cfg)
    meta = MetadatosCorrida(hashlib.sha256(ruta_libro_real.read_bytes()).hexdigest())
    num = repo.iniciar(cfg, meta)
    resumen = repo.publicar_mes(cfg.id_corrida, num, construir_paquete_mes(r, meta, repo.cargar_catalogo()))
    repo.finalizar(cfg.id_corrida, "success", publicada=True, meses_publicados=[resumen.mes_texto])
    assert resumen.segundos < 180, resumen  # objetivo 2.6: < 3 min
    c14 = _consulta(
        engine,
        "SELECT BloquesRacksIndisponibles FROM dbo.disponibilidad_mensual WHERE IdProyecto = 1 AND Anio = 2026 "
        "AND Mes = 9",
    )[0][0]
    assert c14 == 104134.2920000001
    # recarga idéntica: mismas filas, todas las detenciones actualizadas, ninguna corrección
    cfg2 = dataclasses.replace(cfg, id_corrida=str(uuid.uuid4()))
    num2 = repo.iniciar(cfg2, meta)
    otra = repo.publicar_mes(cfg2.id_corrida, num2, construir_paquete_mes(r, meta, repo.cargar_catalogo()))
    repo.finalizar(cfg2.id_corrida, "success", publicada=True, meses_publicados=[otra.mes_texto])
    assert otra.correcciones == 0 and otra.detenciones["nuevas"] == otra.detenciones["eliminadas"] == 0


# ------------------------------------------------------------------ E2E (tarea 3.4)
@pytest.mark.golden
def test_e2e_corrida_real_con_reconciliacion(engine, repo, ruta_libro_real):
    """Libro real de septiembre → motores → calidad → reconciliación contra sus valores cacheados →
    ``success`` → publica septiembre con el anual desde julio (jul/ago ``excel_manual``)."""
    import json

    from etl_arena.config import construir_config, parametros_desde_celdas
    from etl_arena.ejecucion import ejecutar_corrida
    from etl_arena.ingestion import leer_celdas_parametros
    from etl_arena.workbook import extraer_referencia, sha256_archivo

    limpiar(engine, 1)
    aplicar_script(engine, DIR_SQL / "06_seed_annual_manual.sql")
    cfg = construir_config(
        **parametros_desde_celdas(leer_celdas_parametros(ruta_libro_real)).valores,
        archivo_origen=ruta_libro_real.name,
    )
    res = ejecutar_corrida(
        ruta_libro_real,
        cfg,
        repo,
        hash_archivo=sha256_archivo(ruta_libro_real),
        referencia=extraer_referencia(ruta_libro_real),
    )
    assert res.estado == "success" and res.error is None and res.publicada

    estado, id_ref, resumen, publicada = _consulta(
        engine,
        "SELECT Estado, IdReferenciaExcel, ResumenCalidad, Publicada FROM dbo.etl_run WHERE IdCorrida = ?",
        cfg.id_corrida,
    )[0]
    assert estado == "success" and id_ref is not None and publicada
    calidad = json.loads(resumen)
    assert calidad["reconciliacion"]["estado"] == "pass" and calidad["publicacion"]["mes"] == "2026-09"
    assert calidad["plant_activity"]["pa_sin_timestamp"] == 659 and calidad["timestamp"] == {"dst_salto": 1}

    niveles = _consulta(
        engine,
        "SELECT Nivel, Metrica, Aprobado FROM dbo.reconciliation_result WHERE IdCorrida = ? AND Metrica "
        "IN ('resumen', 'no_aplica') ORDER BY Nivel",
        cfg.id_corrida,
    )
    assert [(n, m, bool(a)) for n, m, a in niveles] == [
        (1, "resumen", True),
        (2, "resumen", True),
        (3, "resumen", True),
        (4, "resumen", True),
        (5, "resumen", True),
        (9, "no_aplica", True),
    ]  # invariantes: C31 ≠ L14 en este libro
    assert _consulta(engine, "SELECT C14 FROM dbo.excel_reference_run WHERE IdReferencia = ?", id_ref)[0][0] == (
        104134.2920000001
    )
    mensual = _consulta(
        engine,
        "SELECT BloquesMuestreo, BloquesRacksIndisponibles, Origen FROM dbo.disponibilidad_mensual "
        "WHERE IdProyecto = 1 AND Anio = 2026 AND Mes = 9",
    )[0]
    assert tuple(mensual) == (1975.0, 104134.2920000001, "corrida")
    anual = _consulta(
        engine,
        "SELECT Mes, BloquesMuestreoAcumulados FROM dbo.v_disponibilidad_anual WHERE IdProyecto = 1 AND Anio = 2026 "
        "ORDER BY Mes",
    )
    assert [tuple(a) for a in anual] == [(7, 2976.0), (8, 5952.0), (9, 7927.0)]
    assert _consulta(engine, "SELECT COUNT(*) FROM dbo.detencion WHERE IdProyecto = 1 AND Mes = 9")[0][0] == 334

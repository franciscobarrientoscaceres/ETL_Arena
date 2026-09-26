"""Revisa que este computador tenga todo lo necesario para correr el proyecto (docs/instalacion.md).

    .venv\\Scripts\\python scripts\\verificar_entorno.py              # todo
    .venv\\Scripts\\python scripts\\verificar_entorno.py --sin-bd     # sin conectarse a la base de datos
    .venv\\Scripts\\python scripts\\verificar_entorno.py --sin-excel  # sin abrir Excel

Cada revisión muestra OK, AVISO o FALTA, y qué hacer si algo falta. No modifica nada.
Código de salida: 0 = todo OK (o solo avisos) · 1 = falta algo.
"""

from __future__ import annotations

import argparse
import contextlib
import importlib
import os
import socket
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))
os.chdir(RAIZ)  # .env y data/ se buscan desde la raíz del proyecto

from etl_arena.ambientes import AYUDA_ENTORNO, OPCIONES_ENTORNO  # noqa: E402

DRIVER = "ODBC Driver 18 for SQL Server"
PAQUETES = {
    "numpy": "numpy",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "sqlalchemy": "sqlalchemy",
    "pyodbc": "pyodbc",
    "azure.identity": "azure-identity",
    "win32com.client": "pywin32",
    "oletools": "oletools",
    "etl_arena": "el propio proyecto",
}
resultados: list[str] = []


def informar(estado: str, que: str, detalle: str = "", solucion: str = "") -> None:
    resultados.append(estado)
    print(f"[{estado:5}] {que}" + (f": {detalle}" if detalle else ""))
    if solucion and estado != "OK":
        print(f"        -> {solucion}")


def revisar_python() -> None:
    v = sys.version_info
    texto = f"{v.major}.{v.minor}.{v.micro}"
    if (v.major, v.minor) >= (3, 13):
        informar("OK", "Python", texto)
    else:
        informar(
            "FALTA",
            "Python",
            f"{texto} (se necesita 3.13 o más nuevo)",
            "instalar Python 3.13 (docs/instalacion.md paso 2)",
        )
    if ".venv" not in sys.executable:
        informar(
            "AVISO",
            "Entorno virtual",
            f"se está usando {sys.executable}",
            "correr con .venv\\Scripts\\python (docs/instalacion.md paso 4)",
        )
    else:
        informar("OK", "Entorno virtual", ".venv")


def revisar_paquetes() -> None:
    faltan = []
    for modulo, nombre in PAQUETES.items():
        try:
            importlib.import_module(modulo)
        except Exception:
            faltan.append(nombre)
    if faltan:
        informar("FALTA", "Paquetes", ", ".join(faltan), '.venv\\Scripts\\python -m pip install -e ".[dev]"')
    else:
        informar("OK", "Paquetes", "todos instalados")


def revisar_libros() -> None:
    libros = sorted((RAIZ / "data").glob("*.xlsm"))
    if libros:
        informar("OK", "Libros Excel en data/", ", ".join(p.name[:45] for p in libros))
    else:
        informar("FALTA", "Libros Excel en data/", "no hay .xlsm", "volver a clonar el repositorio (git clone)")


def revisar_driver() -> None:
    try:
        import pyodbc

        drivers = pyodbc.drivers()
    except Exception as exc:
        informar("FALTA", "Driver ODBC", str(exc), "instalar pyodbc (paso 4)")
        return
    if DRIVER in drivers:
        informar("OK", "Driver ODBC", DRIVER)
    else:
        informar(
            "FALTA",
            "Driver ODBC",
            f"no está '{DRIVER}' (hay: {', '.join(drivers) or 'ninguno'})",
            "pedir a TI que instale 'ODBC Driver 18 for SQL Server' (necesita administrador)",
        )


def revisar_env() -> object | None:
    from etl_arena.ambientes import AMBIENTES, url_azure
    from etl_arena.persistence.conexion import _texto_url, cargar_env, entorno, url_configurada

    if (RAIZ / ".env").exists():
        cargar_env(RAIZ / ".env")
        informar("OK", "Archivo .env", "encontrado")
    else:
        informar("AVISO", "Archivo .env", "no existe: se usan las bases por defecto", "opcional (paso 5)")
    try:
        url = url_configurada()
    except Exception as exc:
        informar(
            "FALTA", "Base de datos configurada", str(exc), "revisar ETL_ARENA_ENTORNO / ETL_ARENA_DB_URL* en .env"
        )
        return None
    activo = entorno()
    for nombre, datos in AMBIENTES.items():
        texto = _texto_url(nombre)
        origen = "por defecto" if texto == url_azure(datos["base"]) else "reemplazada en .env"
        marca = "  <- activo" if nombre == activo else ""
        informar("OK", f"Ambiente {datos['etiqueta']}", f"{texto.split('@', 1)[-1].split('?')[0]} ({origen}){marca}")
    prueba_auto = os.environ.get("ETL_ARENA_TEST_DB_URL", "")
    if prueba_auto:
        from sqlalchemy.engine import make_url

        from etl_arena.ambientes import es_base_de_ambiente

        base_tests = make_url(prueba_auto).database
        if es_base_de_ambiente(base_tests):
            informar(
                "FALTA",
                "Base de tests automáticos",
                f"ETL_ARENA_TEST_DB_URL apunta a {base_tests}, que es un ambiente",
                "borrar esa línea del .env o usar una base propia con 'test' en el nombre (nunca PROD ni TEST/QA)",
            )
        else:
            informar("OK", "Base de tests automáticos", base_tests or "(sin nombre)")
    if (url.host or "").endswith(".database.windows.net") and os.environ.get(
        "ETL_ARENA_DB_AUTH", "entra"
    ).lower() not in (
        "",
        "entra",
    ):
        informar(
            "AVISO",
            "Inicio de sesión Azure",
            f"ETL_ARENA_DB_AUTH={os.environ['ETL_ARENA_DB_AUTH']}",
            "contra Azure se usa ETL_ARENA_DB_AUTH=entra (o borrar la línea)",
        )
    return url


def revisar_red(url) -> bool:
    host, puerto = url.host, int(url.port or 1433)
    if not host or "\\" in host:
        informar("OK", "Red", f"instancia local {host}")
        return True
    try:
        with socket.create_connection((host, puerto), timeout=8):
            informar("OK", "Red hacia la base", f"{host}:{puerto} responde")
            return True
    except OSError as exc:
        informar(
            "FALTA",
            "Red hacia la base",
            f"{host}:{puerto} no responde ({exc})",
            "la red bloquea el puerto 1433: probar desde otra red o pedirlo a TI",
        )
        return False


def revisar_bd() -> None:
    import sqlalchemy as sa

    from etl_arena.persistence import crear_engine

    print("        (puede tardar ~1 min si la base estaba pausada, y abrir el navegador para iniciar sesión)")
    try:
        engine = crear_engine()
        with engine.connect() as c:
            base = c.execute(sa.text("SELECT DB_NAME()")).scalar()
            corridas = c.execute(sa.text("SELECT COUNT(*) FROM dbo.etl_run")).scalar()
            v2 = c.execute(sa.text("SELECT OBJECT_ID('dbo.muestra_pcs')")).scalar() is not None
        engine.dispose()
        if not v2:
            informar(
                "FALTA",
                "Base de datos",
                f"{base} tiene el esquema antiguo (una copia por corrida)",
                "actualizar al estado vigente por mes: scripts\\crear_base.py (ADR-12; docs/instalacion.md)",
            )
            return
        informar(
            "OK", "Base de datos", f"conectado a {base}; esquema v2 (estado vigente por mes); {corridas} ejecuciones"
        )
    except Exception as exc:
        texto = str(exc)
        if "Client with IP address" in texto or "40615" in texto:
            solucion = "agregar la IP de esta red en el firewall de Azure (paso 6)"
        elif "etl_run" in texto:
            solucion = "la base existe pero no tiene las tablas: scripts\\crear_base.py (solo si es una base nueva)"
        else:
            solucion = "revisar .env, el inicio de sesión y el firewall (docs/instalacion.md)"
        informar("FALTA", "Base de datos", texto.splitlines()[0][:200], solucion)


def revisar_excel() -> None:
    try:
        from etl_arena.workbook.com import SesionExcel

        with SesionExcel(visible=False, timeout_s=90) as s:
            version = s.version
            build = s.build
    except Exception as exc:
        informar("FALTA", "Excel", str(exc)[:200], "instalar Excel de escritorio (Office) en este computador")
        return
    if build < 10000:
        informar("OK", "Excel", f"{version} (Excel 2016: avisos 438 esperables, F-40)")
    else:
        informar("OK", "Excel", version)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--sin-bd", action="store_true", help="no conectarse a la base de datos")
    ap.add_argument("--sin-excel", action="store_true", help="no abrir Excel")
    ap.add_argument("--entorno", choices=OPCIONES_ENTORNO, help=AYUDA_ENTORNO)
    args = ap.parse_args(argv)
    with contextlib.suppress(Exception):  # consolas sin UTF-8: no caerse por un carácter
        sys.stdout.reconfigure(errors="replace")
    if args.entorno:
        os.environ["ETL_ARENA_ENTORNO"] = args.entorno

    print("Revisando este computador...\n")
    revisar_python()
    revisar_paquetes()
    revisar_libros()
    revisar_driver()
    url = revisar_env()
    if url is not None and not args.sin_bd and revisar_red(url):
        revisar_bd()
    if not args.sin_excel:
        revisar_excel()

    faltan, avisos = resultados.count("FALTA"), resultados.count("AVISO")
    print()
    if faltan:
        print(f"Falta {faltan} cosa(s). Resuélvelas siguiendo las flechas '->' y vuelve a correr este script.")
        return 1
    print("Todo listo." + (f" ({avisos} aviso(s) opcional(es).)" if avisos else ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

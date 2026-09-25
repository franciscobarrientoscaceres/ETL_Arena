"""Ejecución de macros sin Excel (4.2): errores VBA y tolerancia 438 de mcoOrder en Excel 2016."""

import pytest

from etl_arena.config import construir_config
from etl_arena.workbook import macros
from etl_arena.workbook.com import ErrorVBA, SesionExcel

MSG_438 = "Se ha producido el error '438' en tiempo de ejecución: El objeto no admite esta propiedad o método"


class _Libro:
    def __getattr__(self, _):
        return lambda *a, **k: None


class _Sesion:
    def __init__(self, build, fallas):
        self.build, self.version, self.fallas, self.corridas = build, f"16.0 build {build}", fallas, []

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return None

    def abrir(self, ruta):
        return _Libro()

    def ejecutar_macro(self, wb, macro):
        self.corridas.append(macro)
        if macro in self.fallas:
            raise ErrorVBA(macro, self.fallas[macro], 3.0)
        return 1.0


@pytest.fixture
def cfg():
    from datetime import date

    return construir_config(inicio_periodo=date(2026, 9, 1), fin_periodo=date(2026, 9, 7))


@pytest.fixture
def copia(tmp_path):
    libro = tmp_path / "libro.xlsm"
    libro.write_bytes(b"")
    (tmp_path / "libro.xlsm.bak").write_bytes(b"")
    return libro


def _correr(monkeypatch, cfg, sesion, libro):
    monkeypatch.setattr(macros, "SesionExcel", lambda **k: sesion)
    monkeypatch.setattr(macros, "escribir_parametros", lambda wb, c, matriz: None)
    return macros.ejecutar_macros(libro, cfg, excel_aplica_matriz=False)


def test_438_de_add2_se_tolera_en_excel_2016(monkeypatch, cfg, copia, caplog):
    sesion = _Sesion(4266, {"mcoCreateList": MSG_438, "Graphupdate": MSG_438})
    tiempos = _correr(monkeypatch, cfg, sesion, copia)
    assert sesion.corridas == list(macros.MACROS)
    assert tiempos["mcoCreateList"] == tiempos["Graphupdate"] == 3.0
    assert "N:Q" in caplog.text and "Graph" in caplog.text


@pytest.mark.parametrize(
    ("build", "fallas"),
    [
        (17928, {"mcoCreateList": MSG_438}),  # Excel moderno: el 438 es un error real
        (4266, {"mcoCreateList": "Se ha producido el error '9': Subíndice fuera del intervalo"}),
        (4266, {"cmdCalcAvailability": MSG_438}),
    ],
)
def test_otros_errores_vba_detienen_la_cadena(monkeypatch, cfg, copia, build, fallas):
    sesion = _Sesion(build, fallas)
    with pytest.raises(ErrorVBA):
        _correr(monkeypatch, cfg, sesion, copia)
    assert sesion.corridas[-1] == next(iter(fallas))


def test_no_corre_macros_sobre_un_libro_que_no_es_copia_de_trabajo(monkeypatch, cfg, tmp_path):
    maestro = tmp_path / "maestro.xlsm"
    maestro.write_bytes(b"")
    sesion = _Sesion(4266, {})
    with pytest.raises(ValueError, match="copia de trabajo"):
        _correr(monkeypatch, cfg, sesion, maestro)
    assert sesion.corridas == []


def test_error_vba_lleva_el_texto_del_dialogo():
    exc = ErrorVBA("mcoCreateList", MSG_438, 2.5)
    assert "438" in str(exc) and exc.macro == "mcoCreateList" and exc.segundos == 2.5


def test_sesion_sin_excel_no_arranca_vigilante():
    s = SesionExcel(timeout_s=5)
    assert s.errores_vba == [] and not s._vigilando.is_set()


def test_abrir_falla_si_el_libro_quedo_de_solo_lectura(tmp_path):
    """Otro Excel tiene el libro abierto: Workbooks.Open no falla, pero Save fallaría al final."""
    from types import SimpleNamespace

    from etl_arena.workbook.com import ErrorExcel

    libro = tmp_path / "libro.xlsm"
    libro.write_bytes(b"")
    s = SesionExcel()
    s.app = SimpleNamespace(Workbooks=SimpleNamespace(Open=lambda *a, **k: SimpleNamespace(ReadOnly=True)))
    with pytest.raises(ErrorExcel, match="solo lectura"):
        s.abrir(libro)
    assert len(s._libros) == 1  # queda registrado para cerrarlo al salir


class _Rango:
    def __init__(self):
        self.Value2 = None
        self.Formula = None

    def ClearContents(self):
        self.Value2 = self.Formula = None


class _Hoja:
    def __init__(self):
        self.celdas = {}

    def Range(self, ref):
        return self.celdas.setdefault(ref, _Rango())


class _LibroFalso:
    def __init__(self):
        self.hojas = {}

    def Worksheets(self, nombre):
        return self.hojas.setdefault(nombre, _Hoja())


@pytest.mark.parametrize(("aplica", "esperado"), [(False, "No"), (True, "Yes")])
def test_c31_l14_yes_solo_si_las_macros_aplican_la_matriz(aplica, esperado):
    """Las macros de septiembre excusarían con PlantActivity!D: nunca se les escribe "Yes" (F-44)."""
    from datetime import date

    cfg = construir_config(inicio_periodo=date(2026, 9, 1), fin_periodo=date(2026, 9, 7))  # oficial: Yes/Yes
    wb = _LibroFalso()
    macros.escribir_parametros(wb, cfg, excel_aplica_matriz=aplica)
    assert wb.hojas["Calculation-Availability"].celdas["C31"].Value2 == esperado
    assert wb.hojas["ListOfFaults"].celdas["L14"].Value2 == esperado
    assert wb.hojas["Calculation-Availability"].celdas["C21"].Value2 == "No"


def test_deteccion_de_la_regla_de_la_matriz_en_el_vba():
    from etl_arena.workbook.vba import _procedimiento

    v11 = """Sub cmdCalcAvailability()
 Set EM = Worksheets("Exclusion_Matrix")
End Sub
Sub mcoCreateList()
 x = Worksheets("Exclusion_Matrix").Cells(1, 1)
End Sub
"""
    assert "Exclusion_Matrix" in _procedimiento(v11, "cmdCalcAvailability")
    assert "Exclusion_Matrix" in _procedimiento(v11, "mcoCreateList")
    solo_kpi = v11.replace('x = Worksheets("Exclusion_Matrix").Cells(1, 1)', "x = 1")
    assert "Exclusion_Matrix" not in _procedimiento(solo_kpi, "mcoCreateList")

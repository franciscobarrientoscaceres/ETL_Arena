"""Tests de etl_arena.excel_semantics (R17). Casos tomados del libro de septiembre."""

from datetime import date, datetime

import pytest

from etl_arena.excel_semantics import (
    ExcelTypeMismatch,
    codigo_falla_excel,
    comparar_variants,
    datetime_a_serial,
    es_vacio,
    flag_no,
    flag_si,
    igual_numero,
    igual_texto,
    menor_que,
    redondear_excel,
    serial_a_datetime,
    texto_excel,
)


class TestVacio:
    def test_none_y_texto_vacio(self):
        assert es_vacio(None) and es_vacio("")

    @pytest.mark.parametrize("v", [0.0, " ", "NO FAULTS", False])
    def test_no_vacios(self, v):
        assert not es_vacio(v)


class TestComparacionNumerica:
    """Cells(...) = 4 / Cells(...) < 4 — R17.1, F-16."""

    def test_modulos_fraccionarios(self):  # F-02: 3.2347 < 4
        assert menor_que(3.2346666666666666, 4) and not igual_numero(3.2346666666666666, 4)

    def test_empty_vale_cero(self):
        assert menor_que(None, 4) and not igual_numero(None, 4)

    def test_texto_numerico_se_convierte(self):
        assert igual_numero("4", 4) and menor_que(" 3.5 ", 4)

    @pytest.mark.parametrize(
        "v", ["Arena - PCS 04 - POWERELECTRONICS HEM-k NUMBER OF MODULES", "", "inf", "nan", "1_0"]
    )
    def test_texto_no_numerico_es_type_mismatch(self, v):  # F-11 (encabezado), F-16
        with pytest.raises(ExcelTypeMismatch):
            igual_numero(v, 4)

    def test_booleano_true_es_menos_uno(self):
        assert menor_que(True, 0) and igual_numero(False, 0)


class TestComparacionTexto:
    """Variant vs literal String: comparación de texto binaria — F-12."""

    def test_flags_exactos(self):
        assert flag_si("Yes") and not flag_si("yes") and not flag_si("YES") and not flag_si(None)
        assert flag_no("No") and not flag_no("no") and not flag_no("") and not flag_no(None)

    def test_numero_vs_texto_no_falla(self):  # falla numérica 169 vs "NO FAULTS"
        assert not igual_texto(169.0, "NO FAULTS")
        assert igual_texto(169.0, "169")

    def test_empty_vs_texto_vacio(self):  # anterior vacío <> "NO FAULTS" → verdadero (F-04)
        assert igual_texto(None, "") and not igual_texto(None, "NO FAULTS")

    def test_numero_no_vacio(self):  # Cells(...) <> "" con número
        assert not igual_texto(0.0, "")


class TestComparacionVariants:
    def test_numeros(self):
        assert comparar_variants(46266.0, 46287.0) == -1

    def test_texto_es_mayor_que_numero(self):  # F-21: fecha como texto queda fuera de "< C7+1"
        assert comparar_variants("01-09-2026 00:00:00", 46287.0) == 1
        assert comparar_variants(46287.0, "x") == -1

    def test_empty(self):
        assert comparar_variants(None, 0.0) == 0 and comparar_variants(None, "") == 0
        assert comparar_variants(None, 1.0) == -1


class TestTextoExcel:
    @pytest.mark.parametrize(
        ("v", "esperado"),
        [
            (55.0, "55"),
            (3.5, "3.5"),
            (None, ""),
            (True, "TRUE"),
            (0.1 + 0.2, "0.3"),
            (1e16, "1E+16"),
            ("F55 X", "F55 X"),
        ],
    )
    def test_general(self, v, esperado):
        assert texto_excel(v) == esperado


class TestCodigoFalla:
    """IFERROR(MID(G,1,FIND(" ",G,1)-1),CONCATENATE("F",G)) — F-14, R8.9."""

    @pytest.mark.parametrize(
        ("g", "esperado"),
        [
            (55.0, "F55"),
            (169.0, "F169"),
            ("F55 EXTERNAL FAULT/OVGR", "F55"),
            ("F1 Watchdog", "F1"),
            ("F13 NO MODULES", "F13"),
            (" X", ""),
            ("ABC", "FABC"),
            ("", "F"),
        ],
    )
    def test_casos_del_libro(self, g, esperado):
        assert codigo_falla_excel(g) == esperado


class TestRedondeo:
    @pytest.mark.parametrize(
        ("x", "d", "esperado"),
        [(2.5, 0, 3.0), (-2.5, 0, -3.0), (2.675, 2, 2.68), (15.000000000058208, 2, 15.0), (60.00000000001, 2, 60.0)],
    )
    def test_mitad_lejos_de_cero(self, x, d, esperado):
        assert redondear_excel(x, d) == esperado

    def test_c23_del_libro(self):  # C23 = ROUND((E5-E4)*24*60, 2) con E4/E5 de septiembre
        e4, e5 = 46266.0, 46266.010416666664
        assert redondear_excel((e5 - e4) * 24 * 60, 2) == 15.0


class TestFechas:
    def test_serial_fila_2(self):
        assert serial_a_datetime(46120.010416666664) == datetime(2026, 4, 8, 0, 15)

    def test_ida_y_vuelta_exacta(self):  # el serial de 00:15 es el mismo double que guarda Excel
        assert datetime_a_serial(datetime(2026, 4, 8, 0, 15)) == 46120.010416666664

    def test_fecha_es_entero(self):
        assert datetime_a_serial(date(2026, 9, 1)) == 46266.0
        assert serial_a_datetime(46286.59375) == datetime(2026, 9, 21, 14, 15)


class TestCasosBorde:
    def test_cstr(self):
        from etl_arena.excel_semantics import cstr_vba

        assert (cstr_vba(None), cstr_vba(True), cstr_vba(3.5), cstr_vba(4.0), cstr_vba("x")) == (
            "",
            "True",
            "3.5",
            "4",
            "x",
        )

    def test_comparar_booleanos_y_texto(self):
        assert comparar_variants(True, 0.0) == -1  # True = -1
        assert comparar_variants("b", "a") == 1 and comparar_variants("", None) == 0

    def test_tipo_no_soportado(self):
        with pytest.raises(ExcelTypeMismatch):
            menor_que([1], 4)

"""Tests de aggregation: Daily (R9) y mensual/anual (R10)."""

from datetime import date

import pytest
from fabricas import actividad, config_prueba, matriz, serial_min

from etl_arena.aggregation import (
    ErrorMensual,
    bloques_calendario,
    calcular_anual,
    calcular_diaria,
    registrar_mes_oficial,
)
from etl_arena.availability import calcular
from etl_arena.model import KpiMensual

DOS_DIAS = {"inicio_periodo": date(2026, 9, 1), "fin_periodo": date(2026, 9, 2)}


def _res(modulos, seriales, operacional=1.0, **cfg):
    config = config_prueba(total_pcs=len(modulos[0]), **{**DOS_DIAS, **cfg})
    m = matriz(modulos, seriales=seriales)
    return calcular(m, actividad(m.n, operacional=operacional), config), config


class TestDiaria:
    def test_dos_dias(self):
        seriales = [serial_min(15), serial_min(30), serial_min(24 * 60 + 15)]
        res, cfg = _res([[3, 4], [2, 4], [4, 0]], seriales)
        dias = calcular_diaria(res, cfg)
        assert [d.dia for d in dias] == [date(2026, 9, 1), date(2026, 9, 2)]
        assert [d.diario for d in dias] == [12 * 1 + 12 * 2, 12 * 4]
        assert [d.acumulado for d in dias] == [36.0, 84.0]
        assert dias[0].disponibilidad == 1 - 36 / (cfg.total_racks * 24 * 60 * 1 / 15)
        assert dias[1].disponibilidad == 1 - 84 / (cfg.total_racks * 24 * 60 * 2 / 15)
        assert dias[0].variacion == 0.0 and dias[1].variacion == dias[1].disponibilidad - dias[0].disponibilidad

    def test_dia_sin_filas_vale_cero(self):
        res, cfg = _res([[3]], [serial_min(24 * 60 + 15)])
        assert [d.diario for d in calcular_diaria(res, cfg)] == [0.0, 12.0]

    def test_sin_factor_operacional(self):  # F-08
        seriales = [serial_min(15), serial_min(30)]
        res, cfg = _res([[3], [3]], seriales, operacional=[0.0, 0.0], solo_tiempo_operacional=True)
        assert res.bloques_racks_indisponibles == 0.0
        assert calcular_diaria(res, cfg)[0].diario == 24.0

    def test_c23_indefinido_da_disponibilidad_vacia(self):
        res, cfg = _res([[3]], [serial_min(15)])
        dias = calcular_diaria(res, cfg)
        assert dias[0].disponibilidad is None and dias[0].variacion == 0.0 and dias[1].variacion is None

    def test_fin_diario_menor_que_fin_periodo(self):  # F-33: la serie corta en Daily!C
        seriales = [serial_min(15), serial_min(24 * 60 + 15)]
        res, cfg = _res([[3], [3]], seriales, fin_diario=date(2026, 9, 1))
        assert [d.diario for d in calcular_diaria(res, cfg)] == [12.0]


class TestMensualAnual:
    def test_mes_incompleto(self):  # sep-2026: DiasMes = último serial − día 1; bloques = C12 (D-07)
        res, cfg = _res([[3], [4]], [serial_min(15), serial_min(24 * 60 + 60 * 14 + 15)])
        k = registrar_mes_oficial(res, cfg)
        assert (k.anio, k.mes, k.bloques_muestreo, k.origen) == (2026, 9, 2.0, "corrida")
        assert k.dias_mes == pytest.approx(1 + 14.25 / 24, abs=1e-12)

    def test_mes_completo(self):
        res, cfg = _res([[3]], [serial_min(15)], fin_periodo=date(2026, 9, 30), fin_diario=date(2026, 9, 30))
        assert registrar_mes_oficial(res, cfg).dias_mes == 30.0

    @pytest.mark.parametrize(
        ("ini", "fin"), [(date(2026, 9, 2), date(2026, 9, 3)), (date(2026, 9, 1), date(2026, 10, 1))]
    )
    def test_periodo_no_mensual(self, ini, fin):
        res, cfg = _res([[3]], [serial_min(15)], inicio_periodo=ini, fin_periodo=fin, fin_diario=ini)
        with pytest.raises(ErrorMensual):
            registrar_mes_oficial(res, cfg)

    def test_bloques_calendario(self):
        assert bloques_calendario(31) == 2976 and bloques_calendario(20.59375) == 1977

    def test_annual_ava_jul_ago_sep(self):  # GT-5: valores de Annual_AVA del libro de septiembre
        meses = [
            KpiMensual(2026, 7, 31, 2976, 350972.29999999946, "excel_manual"),
            KpiMensual(2026, 8, 31, 2976, 305182.96799999941, "excel_manual"),
            KpiMensual(2026, 9, 20.59375, 1977, 104134.2920000001, "corrida"),
        ]
        filas = calcular_anual(meses, 2928, 2026, 7)
        assert [f.bloques_muestreo_acumulados for f in filas] == [2976, 5952, 7929]
        assert filas[-1].bloques_indisponibles_acumulados == 760289.55999999901
        assert filas[-1].disponibilidad_acumulada == 0.96725164144625086
        assert filas[-1].disponibilidad_mensual == 0.98201062699182673

    def test_mes_sin_dato_reinicia_bloques(self):  # H = IF(F<>"", E+H_anterior, 0)
        meses = [KpiMensual(2026, 7, 31, 2976, 100.0, "corrida"), KpiMensual(2026, 9, 30, 2880, 50.0, "corrida")]
        filas = calcular_anual(meses, 2928, 2026, 7)
        assert [f.bloques_muestreo_acumulados for f in filas] == [2976, 0.0, 2880]
        assert [f.bloques_indisponibles_acumulados for f in filas] == [100.0, 100.0, 150.0]
        assert filas[1].disponibilidad_acumulada is None and filas[1].disponibilidad_mensual is None

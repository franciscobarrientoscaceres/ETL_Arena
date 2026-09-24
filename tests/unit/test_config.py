"""Tests de etl_arena.config (R1)."""

from datetime import date, datetime

import pytest

from etl_arena.config import (
    CONFIG_POR_DEFECTO,
    VERSION_ALGORITMO,
    ErrorConfiguracion,
    construir_config,
    parametros_desde_celdas,
)

SEP = {"inicio_periodo": date(2026, 9, 1), "fin_periodo": date(2026, 9, 21)}


def _celdas_septiembre(**cambios):
    """Celdas de parámetros tal como están en el libro de septiembre."""
    c = {
        "Calculation-Availability!C2": 61.0,
        "Calculation-Availability!C3": 4.0,
        "Calculation-Availability!C5": 46266.0,
        "Calculation-Availability!C7": 46286.0,
        "Calculation-Availability!C11": 2928.0,
        "Calculation-Availability!C21": "No",
        "Calculation-Availability!C31": "No",
        "ListOfFaults!L2": 46266.0,
        "ListOfFaults!L4": 46286.0,
        "ListOfFaults!L14": "Yes",
        "Daily!D5": 46286.0,
    }
    c.update({f"Daily!C{9 + i}": 46266.0 + i for i in range(21)})
    c.update(cambios)
    return c


class TestConstruirConfig:
    def test_defaults_de_paridad(self):  # R1.3
        cfg = construir_config(**SEP)
        assert (cfg.total_pcs, cfg.baterias_por_pcs, cfg.racks_por_pcs, cfg.minutos_muestreo) == (61, 4, 12, 15)
        assert cfg.version_algoritmo == VERSION_ALGORITMO == "availability-v1.1-exclusion-matrix"  # F-37
        assert cfg.modo_huecos == "excel" and not cfg.solo_tiempo_operacional

    def test_d03_excusable_por_defecto(self):  # D-03: C31 = L14 = "Yes"
        cfg = construir_config(**SEP)
        assert cfg.aplicar_evento_excusable and cfg.aplicar_evento_excusable_eventos

    def test_total_racks(self):  # R1.2
        assert construir_config(**SEP).total_racks == 2928
        assert construir_config(**SEP, total_pcs=10, racks_por_pcs=6).total_racks == 240

    def test_eventos_y_diario_por_defecto_iguales_al_kpi(self):  # R1.4
        cfg = construir_config(**SEP)
        assert cfg.inicio_periodo_eventos == cfg.inicio_periodo
        assert cfg.fin_periodo_eventos == cfg.fin_periodo
        assert cfg.fin_diario == cfg.fin_periodo and cfg.dias_diario == 21

    def test_eventos_independientes(self):  # F-05
        cfg = construir_config(**SEP, inicio_periodo_eventos=date(2026, 9, 2), aplicar_evento_excusable_eventos=False)
        assert cfg.inicio_periodo_eventos == date(2026, 9, 2) and cfg.aplicar_evento_excusable

    def test_id_corrida_unico(self):
        assert construir_config(**SEP).id_corrida != construir_config(**SEP).id_corrida

    def test_datetime_medianoche_aceptado(self):
        cfg = construir_config(inicio_periodo=datetime(2026, 9, 1), fin_periodo=date(2026, 9, 21))
        assert cfg.inicio_periodo == date(2026, 9, 1)

    @pytest.mark.parametrize(
        "cambios",
        [
            {"fin_periodo": date(2026, 8, 31)},
            {"total_pcs": 0},
            {"total_pcs": 61.0},
            {"racks_por_pcs": True},
            {"modo_huecos": "ignorar"},
            {"tipo_corrida": "diaria"},
            {"fin_diario": date(2026, 10, 2)},  # 32 días
            {"solo_tiempo_operacional": "No"},
            {"inicio_periodo": datetime(2026, 9, 1, 0, 15)},
            {"parametro_inventado": 1},
        ],
    )
    def test_validaciones(self, cambios):
        with pytest.raises(ErrorConfiguracion):
            construir_config(**{**SEP, **cambios})

    def test_falta_periodo(self):
        with pytest.raises(ErrorConfiguracion):
            construir_config(inicio_periodo=date(2026, 9, 1))

    def test_defaults_inmutables(self):
        with pytest.raises(TypeError):
            CONFIG_POR_DEFECTO["total_pcs"] = 1  # type: ignore[index]


class TestDesdeExcel:
    def test_libro_septiembre(self):
        p = parametros_desde_celdas(_celdas_septiembre())
        v = p.valores
        assert v["inicio_periodo"] == date(2026, 9, 1) and v["fin_periodo"] == date(2026, 9, 21)
        assert v["racks_por_pcs"] == 12 and v["total_pcs"] == 61
        assert v["solo_tiempo_operacional"] is False
        assert v["aplicar_evento_excusable"] is False and v["aplicar_evento_excusable_eventos"] is True  # F-05
        assert v["fin_diario"] == date(2026, 9, 21) and p.notas == []
        cfg = construir_config(**v)
        assert cfg.total_racks == 2928

    @pytest.mark.parametrize(
        ("c21", "esperado"), [("No", False), ("no", True), ("Yes", True), (None, True), ("", True)]
    )
    def test_c21_distinto_de_No_aplica_operacional(self, c21, esperado):  # F-12
        p = parametros_desde_celdas(_celdas_septiembre(**{"Calculation-Availability!C21": c21}))
        assert p.valores["solo_tiempo_operacional"] is esperado

    @pytest.mark.parametrize(("c31", "esperado"), [("Yes", True), ("yes", False), ("YES", False), (None, False)])
    def test_c31_solo_Yes_exacto(self, c31, esperado):  # F-12, R1.5
        p = parametros_desde_celdas(_celdas_septiembre(**{"Calculation-Availability!C31": c31}))
        assert p.valores["aplicar_evento_excusable"] is esperado

    def test_fin_diario_sale_de_daily_c_no_de_d5(self):  # F-33
        celdas = _celdas_septiembre(**{"Daily!D5": 46295.0})
        for fila in range(24, 30):
            celdas[f"Daily!C{fila}"] = None
        p = parametros_desde_celdas(celdas)
        assert p.valores["fin_diario"] == date(2026, 9, 15)
        assert len(p.notas) == 1 and "F-33" in p.notas[0]

    def test_c11_incoherente(self):
        with pytest.raises(ErrorConfiguracion):
            parametros_desde_celdas(_celdas_septiembre(**{"Calculation-Availability!C11": 2930.0}))

    def test_fecha_con_hora_rechazada(self):
        with pytest.raises(ErrorConfiguracion):
            parametros_desde_celdas(_celdas_septiembre(**{"Calculation-Availability!C5": 46266.5}))

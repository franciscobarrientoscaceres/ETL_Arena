"""Tests de normalization (R5) y enrichment (R6)."""

import numpy as np
import pytest
from fabricas import BLOQUE, SERIAL_SEP_1, config_prueba, crear_libro, encabezado_raw, fila_raw

from etl_arena.enrichment import asociar_actividad
from etl_arena.ingestion import leer_libro
from etl_arena.model import ErrorParidad
from etl_arena.normalization import a_formato_largo, a_matriz, columna, validar_esquema

S = [SERIAL_SEP_1 + (k + 1) * BLOQUE for k in range(4)]


def _matriz(tmp_path, filas, actividad=None, **cfg):
    config = config_prueba(**cfg)
    libro = leer_libro(crear_libro(tmp_path / "l.xlsx", filas, total_pcs=config.total_pcs, actividad=actividad), config)
    mapa, anomalias = validar_esquema(libro.encabezado, config)
    assert not anomalias
    return a_matriz(libro, mapa, config), libro, config


class TestEsquema:
    def test_columnas_por_posicion(self):  # PCS 1 → E (5); PCS 2 → I (9)
        assert columna(1, "modulos") == 5 and columna(2, "modulos") == 9 and columna(2, "falla") == 6

    def test_encabezado_correcto(self):
        enc = dict(enumerate(encabezado_raw(61), start=1))
        mapa, anomalias = validar_esquema(enc, config_prueba(total_pcs=61))
        assert anomalias == [] and mapa.pcs == list(range(1, 62)) and mapa.pcs_encabezado_eventos == 61

    def test_columna_desplazada(self):  # R5.5
        enc = dict(enumerate(encabezado_raw(2), start=1))
        enc[6], enc[7] = enc[7], enc[6]
        _, anomalias = validar_esquema(enc, config_prueba())
        assert {a.tipo for a in anomalias} == {"columna_desplazada"}

    def test_nombre_inesperado_y_faltante(self):
        enc = dict(enumerate(encabezado_raw(2), start=1))
        enc[3] = "Otra cosa"
        del enc[9]
        _, anomalias = validar_esquema(enc, config_prueba())
        assert sorted(a.tipo for a in anomalias) == ["columna_faltante", "nombre_inesperado"]

    def test_mas_pcs_en_encabezado_que_c2(self):  # F-17
        enc = dict(enumerate(encabezado_raw(3), start=1))
        _, anomalias = validar_esquema(enc, config_prueba(total_pcs=2))
        assert [a.tipo for a in anomalias] == ["pcs_distinto_c2"]


class TestMatriz:
    def test_valores_y_nulos(self, tmp_path):  # R5.3, R5.4
        m, _, cfg = _matriz(
            tmp_path, [fila_raw(S[0], [3.2346666666666666, None], [169, "NO FAULTS"]), fila_raw(S[1], [4, 0])]
        )
        assert m.modulos[0, 0] == 3.2346666666666666 and m.modulos_nulo[0, 1] and not m.modulos_nulo[1, 1]
        assert m.falla[0, 0] == 169.0 and m.falla[0, 1] == "NO FAULTS"
        assert m.modulos_disponibles(cfg.baterias_por_pcs)[0, 1] == 4.0
        assert m.marca_tiempo[0] == np.datetime64("2026-09-01T00:15:00")

    @pytest.mark.parametrize("valor", ["abc", "", "3", 4.5, -1.0, True])
    def test_modulos_invalidos_rechazados(self, tmp_path, valor):  # R4.6, F-16
        with pytest.raises(ErrorParidad) as e:
            _matriz(tmp_path, [fila_raw(S[0], [valor, 4])])
        assert e.value.anomalias[0].numero_pcs == 1

    def test_fila_siguiente(self, tmp_path):
        m, _, _ = _matriz(tmp_path, [fila_raw(S[0], [3, 3]), fila_raw(None, [3.5, None])])
        assert m.n == 1 and m.siguiente_modulos[0] == 3.5 and m.siguiente_nulo[1]

    def test_formato_largo(self, tmp_path):  # R5.1
        m, _, _ = _matriz(tmp_path, [fila_raw(S[0], [3.5, None], [169, "NO FAULTS"]), fila_raw(S[1], [4, 4])])
        df = a_formato_largo(m, 4)
        assert len(df) == 4 and list(df["NumeroPCS"]) == [1, 2, 1, 2]
        fila = df.iloc[1]
        assert fila["ModulosDisponiblesNulo"] and fila["ModulosDisponibles"] == 4.0 and np.isnan(fila["ModulosRaw"])
        assert df.iloc[0]["FallaRawEsNumero"] and not df.iloc[1]["FallaRawEsNumero"]


class TestEnriquecimiento:
    def test_union_por_fila_y_vacio_cero(self, tmp_path):  # R6.1, R6.2, F-01
        act = {2: [None, S[0], 1, 0], 3: [None, None, 0, 1], 4: [None, S[3], None, None]}
        m, libro, _ = _matriz(tmp_path, [fila_raw(S[0], [4, 4]), fila_raw(S[1], [4, 4]), fila_raw(S[2], [4, 4])], act)
        datos, anomalias = asociar_actividad(m, libro.actividad)
        assert list(datos.factor_operacional) == [1.0, 0.0, 0.0]
        assert list(datos.factor_excusable) == [0.0, 1.0, 0.0]
        tipos = [(a.tipo, a.numero_fila) for a in anomalias]
        assert ("pa_sin_timestamp", 3) in tipos
        assert ("pa_desalineado", 4) in tipos  # R6.3: B de la fila 4 = S[3] ≠ S[2]
        assert tipos.count(("pa_vacio", 4)) == 2

    def test_texto_en_factor_rechazado(self, tmp_path):
        m, libro, _ = _matriz(tmp_path, [fila_raw(S[0], [4, 4])], {2: [None, S[0], "Sí", 1]})
        with pytest.raises(ErrorParidad):
            asociar_actividad(m, libro.actividad)

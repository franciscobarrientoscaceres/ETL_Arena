"""Tests de etl_arena.ingestion (R4, R16) con libros sintéticos."""

import pytest
from fabricas import BLOQUE, SERIAL_SEP_1, config_prueba, crear_libro, fila_raw

from etl_arena.ingestion import ErrorLibro, LibroXlsx, detectar_anomalias_timestamp, leer_libro
from etl_arena.model import ErrorParidad


def _tipos(anomalias):
    return [a.tipo for a in anomalias]


def _serie(n, desde=SERIAL_SEP_1):
    return [desde + (k + 1) * BLOQUE for k in range(n)]


class TestLibroXlsx:
    def test_hoja_inexistente(self, tmp_path):
        ruta = crear_libro(tmp_path / "l.xlsx", [fila_raw(SERIAL_SEP_1, [4, 4])])
        with LibroXlsx(ruta) as libro, pytest.raises(ErrorLibro):
            list(libro.filas("NoExiste"))

    def test_archivo_inexistente(self, tmp_path):  # R4.3
        with pytest.raises(ErrorLibro):
            LibroXlsx(tmp_path / "no.xlsx")

    def test_celdas_desde_generador(self, tmp_path):
        ruta = crear_libro(tmp_path / "l.xlsx", [fila_raw(SERIAL_SEP_1, [3.5, 4])])
        with LibroXlsx(ruta) as libro:
            assert libro.celdas(r for r in ["RawData-PCS!E2"]) == {"RawData-PCS!E2": 3.5}

    def test_celdas_sueltas(self, tmp_path):
        ruta = crear_libro(tmp_path / "l.xlsx", [fila_raw(SERIAL_SEP_1, [3.5, 4])])
        with LibroXlsx(ruta) as libro:
            c = libro.celdas(["RawData-PCS!A2", "RawData-PCS!E2", "RawData-PCS!Z99", "PlantActivity!B1"])
        assert c == {
            "RawData-PCS!A2": SERIAL_SEP_1,
            "RawData-PCS!E2": 3.5,
            "RawData-PCS!Z99": None,
            "PlantActivity!B1": "Date/Time",
        }


class TestLeerLibro:
    def test_serial_exacto_y_tipos_crudos(self, tmp_path):  # R4.2, R4.7
        s = _serie(2)
        ruta = crear_libro(
            tmp_path / "l.xlsx",
            [fila_raw(s[0], [3.2346666666666666, None], [169, "NO FAULTS"]), fila_raw(s[1], [4, 4])],
        )
        libro = leer_libro(ruta, config_prueba())
        assert [f.serial for f in libro.filas] == s
        assert libro.filas[0].celdas[2] == 169.0 and libro.filas[0].celdas[5] == 3.2346666666666666
        assert 9 not in libro.filas[0].celdas  # módulos PCS 2 vacío

    def test_modo_excel_corta_en_primera_a_vacia(self, tmp_path):  # R4.5, D-01
        s = _serie(5)
        filas = [fila_raw(s[0], [4, 4]), fila_raw(s[1], [4, 4]), [], fila_raw(s[3], [4, 4]), fila_raw(s[4], [4, 4])]
        libro = leer_libro(crear_libro(tmp_path / "l.xlsx", filas), config_prueba())
        assert [f.numero_fila for f in libro.filas] == [2, 3]
        assert libro.filas_descartadas == 2  # filas 5 y 6
        assert _tipos(libro.anomalias) == ["celda_a_vacia", "filas_truncadas", "exclusion_matrix_ausente"]

    def test_modo_excel_a_vacia_con_datos_es_fila_siguiente(self, tmp_path):
        s = _serie(3)
        filas = [fila_raw(s[0], [3, 4]), fila_raw(None, [3.5, 4]), fila_raw(s[2], [4, 4])]
        libro = leer_libro(crear_libro(tmp_path / "l.xlsx", filas), config_prueba())
        assert [f.numero_fila for f in libro.filas] == [2]
        assert libro.fila_siguiente[5] == 3.5 and libro.filas_descartadas == 1

    def test_modo_continuar_omite_solo_vacias(self, tmp_path):  # R4.5
        s = _serie(4)
        filas = [fila_raw(s[0], [4, 4]), [], fila_raw(s[2], [4, 4]), fila_raw(s[3], [4, 4])]
        libro = leer_libro(crear_libro(tmp_path / "l.xlsx", filas), config_prueba(modo_huecos="continuar"))
        assert [f.numero_fila for f in libro.filas] == [2, 4, 5]
        assert libro.filas_descartadas == 0

    def test_a2_vacia_no_corta(self, tmp_path):  # VBA: la salida se evalúa desde la fila 3
        s = _serie(3)
        filas = [fila_raw(None, [4, 4]), fila_raw(s[1], [4, 4]), fila_raw(s[2], [4, 4])]
        libro = leer_libro(crear_libro(tmp_path / "l.xlsx", filas), config_prueba())
        assert [(f.numero_fila, f.serial) for f in libro.filas] == [(2, 0.0), (3, s[1]), (4, s[2])]
        assert "celda_a_vacia" in _tipos(libro.anomalias)

    def test_a_texto_rechazada(self, tmp_path):  # F-21
        ruta = crear_libro(tmp_path / "l.xlsx", [fila_raw("01-09-2026 00:15:00", [4, 4])])
        with pytest.raises(ErrorParidad) as e:
            leer_libro(ruta, config_prueba())
        assert e.value.anomalias[0].tipo == "a_no_numerico"

    def test_actividad_por_fila(self, tmp_path):
        s = _serie(2)
        ruta = crear_libro(
            tmp_path / "l.xlsx",
            [fila_raw(s[0], [4, 4]), fila_raw(s[1], [4, 4])],
            actividad={2: [None, s[0], 1, 0], 3: [None, None, 0, 1]},
        )
        libro = leer_libro(ruta, config_prueba())
        assert libro.actividad[2] == {2: s[0], 3: 1.0, 4: 0.0}
        assert libro.actividad[3] == {3: 0.0, 4: 1.0}


class TestAnomaliasTimestamp:
    def test_serie_regular_sin_anomalias(self):
        s = _serie(10)
        assert detectar_anomalias_timestamp(s, list(range(2, 12)), 15) == []

    def test_tipos(self):
        s = _serie(8)
        seriales = [s[0], s[1], s[1], s[0], s[4], s[4] + 10 / 1440]
        tipos = _tipos(detectar_anomalias_timestamp(seriales, list(range(2, 8)), 15))
        assert tipos == ["duplicado", "fuera_de_orden", "hueco", "frecuencia_distinta"]

    def test_dst_septiembre_es_informativo(self):  # R16.3: 2026-09-06 00:00 → 01:00
        cero = 46271.0  # 2026-09-06 00:00
        seriales = [cero - BLOQUE, cero, cero + 4 * BLOQUE, cero + 5 * BLOQUE]
        a = detectar_anomalias_timestamp(seriales, [2, 3, 4, 5], 15)
        assert [(x.tipo, x.severidad, x.numero_fila) for x in a] == [("dst_salto", "info", 4)]
        assert "faltan 3 filas" in a[0].detalle

    def test_hueco_real_no_es_dst(self):
        cero = 46270.5  # 2026-09-05 12:00
        a = detectar_anomalias_timestamp([cero, cero + 4 * BLOQUE], [2, 3], 15)
        assert [(x.tipo, x.severidad) for x in a] == [("hueco", "advertencia")]

    def test_repeticion_otono_es_informativa(self):  # retroceso de abril 2027: 23:xx del sábado se repite
        from datetime import datetime

        from etl_arena.excel_semantics import datetime_a_serial

        t = datetime_a_serial(datetime(2027, 4, 3, 23, 30))
        a = detectar_anomalias_timestamp([t, t], [2, 3], 15)
        assert [(x.tipo, x.severidad) for x in a] == [("dst_repeticion", "info")]

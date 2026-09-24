"""Valores por defecto de paridad (R1.3). Único lugar con las constantes del activo Arena BESS."""

from __future__ import annotations

from datetime import date
from types import MappingProxyType

# v1.1 (2026-09-24, F-37): la exclusión viene de Exclusion_Matrix (0/1/2) en vez de PlantActivity!D.
# Sin Exclusion_Matrix (libros hasta septiembre) los resultados son idénticos a v1-excel-parity.
VERSION_ALGORITMO = "availability-v1.1-exclusion-matrix"

CONFIG_POR_DEFECTO = MappingProxyType(
    {
        "id_proyecto": 1,
        "version_algoritmo": VERSION_ALGORITMO,
        "nombre_proyecto": "Arena BESS",
        "fecha_inicio_proyecto": date(2026, 4, 8),
        "total_pcs": 61,
        "baterias_por_pcs": 4,
        "racks_por_pcs": 12,
        "minutos_muestreo": 15,
        "solo_tiempo_operacional": False,
        # D-03 (2026-09-24): el KPI oficial descuenta eventos de exclusión (C31 = L14 = "Yes");
        # desde F-37 la fuente es Exclusion_Matrix, no PlantActivity!D.
        "aplicar_evento_excusable": True,
        "aplicar_evento_excusable_eventos": True,
        "modo_huecos": "excel",
        "tipo_corrida": "semanal",
        "es_oficial": False,
        "archivo_origen": "",
    }
)

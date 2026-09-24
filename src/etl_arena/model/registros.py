"""Registros de salida de los motores (espejo de celdas del libro y filas SQL)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Literal

from etl_arena.excel_semantics import serial_a_datetime, texto_excel


@dataclass
class RegistroLista:
    """Una fila ``B:I`` de ``ListOfFaults`` (la fila ``dblResult`` del VBA), mutable como la hoja."""

    orden_excel: int  # 1 = fila 6
    numero_pcs: int | None = None  # B
    serial_inicio: float | None = None  # C
    serial_fin: float | None = None  # D
    duracion_horas: float | None = None  # E
    codigo_falla: str | None = None  # F
    descripcion_falla: object = None  # G (str | float | None, tipo crudo)
    promedio_baterias: float | None = None  # H
    horas_rack: float | None = None  # I
    numero_bloques: int = 0
    suma_bloques: float = 0.0
    fallback: bool = False  # G tomada de la fila anterior (R8.8)
    arrastrado_excel: bool = False  # F-06
    excel_habria_fallado: bool = False  # F-11 / D-08
    cerrado_por_nulo: bool = False  # cerró porque la fila siguiente tenía módulos vacíos
    tiene_exclusion: bool = False  # algún bloque con Exclusion_Matrix ≠ 0 (F-37)
    pcs_iniciados: set[int] = field(default_factory=set)

    @property
    def cerrado(self) -> bool:
        return self.serial_fin is not None


@dataclass(frozen=True)
class EventoFalla:
    """Fila de ``fault_event`` construida desde un ``RegistroLista`` cerrado."""

    orden_excel: int
    numero_pcs: int
    serial_inicio: float
    serial_fin: float
    marca_tiempo_inicio: datetime
    marca_tiempo_fin: datetime
    duracion_horas: float
    codigo_falla: str
    descripcion_falla: str
    descripcion_falla_fallback: bool
    promedio_baterias: float
    horas_rack_indisponibles: float
    numero_bloques: int
    arrastrado_excel: bool
    excel_habria_fallado: bool
    suma_bloques: float = 0.0
    cerrado_por_nulo: bool = False
    tiene_exclusion: bool = False

    @classmethod
    def desde_registro(cls, r: RegistroLista) -> EventoFalla:
        if not r.cerrado or r.numero_pcs is None or r.serial_inicio is None:
            raise ValueError(f"registro {r.orden_excel} incompleto")
        return cls(
            orden_excel=r.orden_excel,
            numero_pcs=r.numero_pcs,
            serial_inicio=r.serial_inicio,
            serial_fin=r.serial_fin,  # type: ignore[arg-type]
            marca_tiempo_inicio=serial_a_datetime(r.serial_inicio),
            marca_tiempo_fin=serial_a_datetime(r.serial_fin),  # type: ignore[arg-type]
            duracion_horas=r.duracion_horas,  # type: ignore[arg-type]
            codigo_falla=r.codigo_falla or "",
            descripcion_falla=texto_excel(r.descripcion_falla),
            descripcion_falla_fallback=r.fallback,
            promedio_baterias=r.promedio_baterias,  # type: ignore[arg-type]
            horas_rack_indisponibles=r.horas_rack,  # type: ignore[arg-type]
            numero_bloques=r.numero_bloques,
            arrastrado_excel=r.arrastrado_excel,
            excel_habria_fallado=r.excel_habria_fallado,
            suma_bloques=r.suma_bloques,
            cerrado_por_nulo=r.cerrado_por_nulo,
            tiene_exclusion=r.tiene_exclusion,
        )


@dataclass(frozen=True)
class MuestraDisponibilidad:
    """Fila de ``availability_sample_result`` (una por fila en rango × PCS)."""

    numero_fila: int
    serial: float
    numero_pcs: int
    baterias_indisponibles: float
    valor_exclusion: float  # Exclusion_Matrix: 0, 1 o 2 (F-37)
    factor_operacional: float
    baterias_ponderadas: float  # espejo de Calculation-Availability!F:BN
    impacto_rack_ponderado: float  # aporte a C14


@dataclass(frozen=True)
class DiaDisponibilidad:
    """Fila ``B:G`` de ``Daily``."""

    numero_dia: int  # B
    dia: date  # C
    diario: float  # D
    acumulado: float  # E
    disponibilidad: float | None  # F (None = "")
    variacion: float | None  # G


@dataclass(frozen=True)
class KpiMensual:
    """Fila de ``monthly_official_kpi`` / ``Annual_AVA`` (R10)."""

    anio: int
    mes: int
    dias_mes: float
    bloques_muestreo: float
    bloques_racks_indisponibles: float
    origen: Literal["corrida", "excel_manual"]
    id_corrida: str | None = None
    disponibilidad_contractual: float = 0.98

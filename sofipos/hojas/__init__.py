"""Construcción y guardado del libro Excel de SOFIPOS."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

from openpyxl import Workbook

from .catalogo import escribir_catalogo
from .datos_fuente import escribir_datos_fuente
from .diccionario import escribir_diccionario
from .indicadores import escribir_indicadores
from .semaforos import escribir_semaforos


def escribir_xlsx(
    universo: list[dict],
    nombres: dict[str, str],
    filas: dict[int, list[dict]],
    ruta: Path,
) -> None:
    """Genera el libro de Excel y lo guarda en la ruta indicada.

    Crea las hojas del catálogo, los datos fuente, los indicadores, el
    diccionario y los semáforos para 2025 y 2026. Guarda primero el archivo
    en un temporal y luego lo mueve a la ruta final.

    Args:
        universo: Sociedades incluidas en el libro.
        nombres: Nombres de las sociedades, indexados por clave.
        filas: Registros mensuales agrupados por año.
        ruta: Ruta de destino del archivo ``.xlsx``.
    """
    wb = Workbook()
    wb.remove(wb.active)

    # Crea las hojas del libro en el orden deseado.
    escribir_catalogo(wb, universo, nombres)
    escribir_datos_fuente(wb, filas[2026], 2026, len(universo))
    escribir_datos_fuente(wb, filas[2025], 2025, len(universo))
    escribir_indicadores(wb, filas[2026], 2026)
    escribir_indicadores(wb, filas[2025], 2025)
    escribir_diccionario(wb)
    escribir_semaforos(wb)

    # Añade metadatos al archivo.
    wb.properties.title = (
        "SOFIPOS - Catálogo, datos fuente, indicadores, diccionario "
        "y semáforos 2025-2026"
    )
    wb.properties.subject = (
        "Cartera, captación, resultados, capitalización e indicadores mensuales"
    )
    wb.properties.creator = "Generador automático (tablero CNBV + Universo BE)"

    # Solicita a Excel recalcular las fórmulas al abrir el libro.
    wb.calculation.fullCalcOnLoad = True

    # Guarda en un archivo temporal y lo reemplaza por el archivo final.
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        "wb",
        suffix=".xlsx",
        dir=ruta.parent,
        delete=False,
    ) as tmp:
        wb.save(tmp.name)
        temporal = tmp.name

    os.replace(temporal, ruta)

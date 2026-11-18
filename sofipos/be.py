"""Lectura de boletines financieros y del universo de SOFIPOS."""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from pathlib import Path

from .config import COL_TOTAL_SECTOR, HOJA_BE, HOJA_UNIVERSO
from .utilidades import _a_numero, _norm, normalizar_clave
from .xlsx_lector import _abrir_hoja, _columnas_sofipo, _fila_encabezado


# Relaciona los nombres internos de los campos con sus etiquetas en el boletín.
ETIQUETAS_BE = {
    "activos": "activo",
    "cartera_1_2": "cartera de credito etapa 1 + etapa 2",
    "cartera_3": "cartera de credito etapa 3",
    "estimaciones": "estimaciones preventivas para riesgos crediticios",
    "ingresos_int": "ingresos por intereses",
    "gastos_int": "gastos por intereses",
    "gastos_adm": "gastos de administracion y promocion",
}


@dataclass
class Boletin:
    """Contiene los registros, totales y filas de origen de un boletín."""

    registros: list[dict]
    totales: dict
    filas: dict[str, int] = field(default_factory=dict)


def leer_be_sofipos(ruta: Path) -> Boletin:
    """Lee los datos de SOFIPOS de un boletín de estados financieros.

    Localiza las filas de cada concepto por su etiqueta y extrae los valores
    por sociedad, junto con los totales del sector.

    Args:
        ruta: Ruta del archivo que contiene el boletín.

    Returns:
        Un objeto con los registros de las sociedades, los totales y las
        filas donde se encontraron los conceptos.

    Raises:
        SystemExit: Si no se encuentra alguna etiqueta requerida.
    """
    hoja = _abrir_hoja(ruta, HOJA_BE)
    fila_enc = _fila_encabezado(hoja, "Clave CASFIM")

    # Las filas siguientes al encabezado contienen datos de identificación.
    codigos = hoja[fila_enc]
    nombres = hoja.get(fila_enc + 1, {})
    federaciones = hoja.get(fila_enc + 2, {})
    niveles = hoja.get(fila_enc + 3, {})

    # Busca la fila correspondiente a cada concepto financiero.
    filas: dict[str, int] = {}
    for campo, etiqueta in ETIQUETAS_BE.items():
        for num in sorted(hoja):
            if num <= fila_enc + 3:
                continue
            if _norm(hoja[num].get("D", "")) == etiqueta:
                filas[campo] = num
                break

        if campo not in filas:
            raise SystemExit(
                f"No se localizó la etiqueta '{etiqueta}' en {ruta.name}."
            )

    # Extrae los datos de cada sociedad identificada en el encabezado.
    columnas = _columnas_sofipo(codigos)
    registros = []
    for col in columnas:
        crudo = codigos[col]
        clave = normalizar_clave(
            int(crudo) if isinstance(crudo, float) else crudo
        )
        registro = {
            "clave": clave,
            "nombre": str(nombres.get(col, "")).strip(),
            "federacion": str(federaciones.get(col, "")).strip(),
            "nivel": str(niveles.get(col, "")).strip(),
        }

        for campo, fila in filas.items():
            valor = hoja.get(fila, {}).get(col)
            registro[campo] = valor.strip() if isinstance(valor, str) else valor

        registros.append(registro)

    # Obtiene el total del sector para cada concepto.
    totales = {
        campo: hoja.get(fila, {}).get(COL_TOTAL_SECTOR)
        for campo, fila in filas.items()
    }
    return Boletin(registros=registros, totales=totales, filas=filas)


def leer_universo(ruta: Path) -> list[dict]:
    """Lee las sociedades autorizadas que están en operación.

    Args:
        ruta: Ruta del archivo que contiene la hoja del universo.

    Returns:
        Lista de sociedades con su clave, nombre, fecha de autorización,
        nivel, federación y estatus.
    """
    hoja = _abrir_hoja(ruta, HOJA_UNIVERSO)
    fila_enc = _fila_encabezado(hoja, "Clave CASFIM")

    registros = []
    for num in sorted(hoja):
        if num <= fila_enc:
            continue

        fila = hoja[num]
        consecutivo = fila.get("C")
        clave = fila.get("D")
        estatus = _norm(fila.get("J", ""))

        # Incluye solo sociedades autorizadas que están en operación.
        if estatus != "en operacion":
            continue
        if (
            _a_numero(consecutivo) is None
            or not re.fullmatch(
                r"\d{4,6}",
                str(clave).strip().split(".")[0],
            )
        ):
            continue

        # Convierte el número de serie de Excel en una fecha.
        serial = _a_numero(fila.get("F"))
        fecha = (
            dt.date(1899, 12, 30) + dt.timedelta(days=int(serial))
            if serial
            else None
        )

        registros.append(
            {
                "consecutivo": int(_a_numero(consecutivo)),
                "clave": normalizar_clave(
                    int(clave) if isinstance(clave, float) else clave
                ),
                "nombre": str(fila.get("E", "")).strip(),
                "fecha_autorizacion": fecha,
                "nivel": str(fila.get("G", "")).strip(),
                "federacion": str(fila.get("I", "")).strip(),
                "estatus": str(fila.get("J", "")).strip(),
            }
        )

    return registros

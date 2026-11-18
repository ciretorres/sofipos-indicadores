"""Lectura y acceso a los datos del tablero en formato CSV."""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from .config import CLAVE_SECTOR, CONCEPTOS, TIPO_SALDO
from .utilidades import _a_numero, normalizar_clave


@dataclass
class Tablero:
    """Almacena los datos leídos del tablero por sociedad, fecha y concepto."""

    # Valores por sociedad y fecha: (clave, fecha) -> campo -> importe en pesos.
    valores: dict[tuple[str, str], dict[str, float]]

    # Totales del sector por fecha y campo: (fecha, campo) -> importe en pesos.
    sector: dict[tuple[str, str], float]

    def dato(self, clave: str, fecha: str, campo: str):
        """Devuelve un dato o ``None`` si no está disponible.

        Args:
            clave: Clave de la sociedad.
            fecha: Fecha del registro.
            campo: Nombre interno del concepto.
        """
        return self.valores.get((clave, fecha), {}).get(campo)


def leer_tablero(ruta: Path, fechas: set[str]) -> Tablero:
    """Lee el CSV y conserva los conceptos y fechas necesarios para el libro.

    Filtra los registros por tipo de saldo, fecha y concepto configurados.
    Separa los totales del sector de los datos de cada sociedad.

    Args:
        ruta: Ruta del archivo CSV del tablero.
        fechas: Fechas que deben conservarse.

    Returns:
        Un ``Tablero`` con los datos de las sociedades y los totales del sector.
    """
    valores: dict[tuple[str, str], dict[str, float]] = {}
    sector: dict[tuple[str, str], float] = {}
    leidas = 0

    with open(ruta, newline="", encoding="utf-8-sig") as f:
        for registro in csv.DictReader(f):
            if registro.get("tipo_saldo") != TIPO_SALDO:
                continue

            fecha = registro.get("fecha", "")
            if fecha not in fechas:
                continue

            campo = CONCEPTOS.get(registro.get("concepto", ""))
            if not campo:
                continue

            valor = _a_numero(registro.get("saldo_se"))
            if valor is None:
                continue

            institucion = str(registro.get("institucion", "")).strip()
            leidas += 1

            # Guarda aparte la fila agregada del sector.
            if institucion == CLAVE_SECTOR:
                sector[(fecha, campo)] = valor
                continue

            # Agrupa los valores de cada sociedad por clave y fecha.
            clave = normalizar_clave(institucion)
            valores.setdefault((clave, fecha), {})[campo] = valor

    print(
        f"  tablero: {leidas:,} registros útiles de {ruta.name} "
        f"({len(fechas)} meses)"
    )
    return Tablero(valores=valores, sector=sector)

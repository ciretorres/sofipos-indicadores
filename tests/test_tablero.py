"""Pruebas de la lectura del CSV del tablero con el fixture mini."""

import csv
from pathlib import Path

import pytest

from sofipos.config import ANIOS
from sofipos.tablero import leer_tablero
from sofipos.utilidades import fecha_mes

FIXTURE = Path(__file__).parent / "fixtures" / "tablero_mini.csv"

# Los 8 meses de 2026 más el mes anterior de enero.
FECHAS = {fecha_mes(2026, mes) for mes in ANIOS[2026]} | {"202512"}


@pytest.fixture
def tablero():
    """Tablero leído una sola vez desde el fixture."""
    return leer_tablero(FIXTURE, FECHAS)


def test_contenido_del_tablero(tablero):
    """El tablero conserva 10 pares (sociedad, fecha) con 20 importes.

    ``027001`` aparece en los 9 meses (``202512``–``202608``), ``027009``
    solo en ``202601`` y el sector aporta 2 filas aparte.
    """
    esperadas = {
        ("027001", "202512"), ("027001", "202601"), ("027001", "202602"),
        ("027001", "202603"), ("027001", "202604"), ("027001", "202605"),
        ("027001", "202606"), ("027001", "202607"), ("027001", "202608"),
        ("027009", "202601"),
    }
    assert set(tablero.valores) == esperadas
    assert sum(len(campos) for campos in tablero.valores.values()) == 20
    assert len(tablero.sector) == 2


def test_descarta_las_cuatro_filas_invalidas():
    """Las 4 filas inválidas del fixture no llegan al objeto ``Tablero``.

    Descartes: ``tipo_saldo`` 140, fecha ``202501`` fuera del set, concepto
    desconocido ``999999`` y ``saldo_se`` no numérico (``n.d.``). Se cuentan
    las entradas del objeto frente a las filas del CSV, sin depender del
    texto que imprime ``leer_tablero``.
    """
    with open(FIXTURE, newline="", encoding="utf-8-sig") as f:
        filas_csv = sum(1 for _ in csv.DictReader(f))

    tablero = leer_tablero(FIXTURE, FECHAS)
    entradas = (
        sum(len(campos) for campos in tablero.valores.values())
        + len(tablero.sector)
    )
    assert filas_csv - entradas == 4
    # Las filas con fecha fuera del set no crean claves nuevas…
    assert ("027001", "202501") not in tablero.valores
    # …y las de tipo 140 o saldo no numérico no pisan el dato válido.
    assert tablero.dato("027001", "202601", "activos") == 1001000.0


def test_fila_del_sector_no_entra_en_valores(tablero):
    """La institución ``27`` va a ``sector`` y no aparece como sociedad."""
    assert tablero.sector[("202601", "activos")] == 9000000.0
    assert ("27", "202601") not in tablero.valores


def test_claves_de_sociedades_normalizadas_a_seis_digitos(tablero):
    """Toda clave de ``valores`` tiene 6 dígitos."""
    assert all(clave.isdigit() and len(clave) == 6 for clave, _ in tablero.valores)


def test_dato_devuelve_el_valor_del_fixture(tablero):
    """``dato`` recupera el importe leído para sociedad, fecha y campo."""
    assert tablero.dato("027001", "202601", "activos") == 1001000.0
    assert tablero.dato("027009", "202602", "activos") is None

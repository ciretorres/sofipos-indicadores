"""Pruebas de la preparación de filas mensuales con el fixture mini."""

from pathlib import Path

import pytest

from sofipos.config import ANIOS
from sofipos.filas import preparar_filas
from sofipos.tablero import leer_tablero
from sofipos.utilidades import fecha_mes

FIXTURE = Path(__file__).parent / "fixtures" / "tablero_mini.csv"
FECHAS = {fecha_mes(2026, mes) for mes in ANIOS[2026]} | {"202512"}

# Universo y nombres construidos en el test, sin leer el BE.
UNIVERSO = [
    {"clave": "027001", "consecutivo": 1, "nombre": "Sociedad Uno"},
    {"clave": "027009", "consecutivo": 2, "nombre": "Sociedad Dos"},
]
NOMBRES = {"027001": "Sociedad Uno Renombrada"}


@pytest.fixture(scope="module")
def filas():
    """Filas de 2026 preparadas una sola vez con el fixture mini."""
    tablero = leer_tablero(FIXTURE, FECHAS)
    return preparar_filas(tablero, UNIVERSO, NOMBRES, 2026)


def _fila(filas, clave, mes):
    """Devuelve la fila de una sociedad para un mes concreto."""
    return next(
        f for f in filas if f["clave"] == clave and f["mes"] == mes
    )


def test_una_fila_por_sociedad_y_mes(filas):
    """Hay 8 meses × 2 sociedades = 16 filas."""
    assert len(filas) == len(ANIOS[2026]) * len(UNIVERSO) == 16


def test_importes_van_en_miles_con_dos_decimales(filas):
    """Los pesos se convierten a miles redondeando a dos decimales."""
    assert _fila(filas, "027001", "Enero")["activos"] == 1001.0
    assert _fila(filas, "027001", "Enero")["gastos_adm"] == 20.12


def test_cartera_usa_el_valor_directo(filas):
    """Con ``272724`` disponible se toma ese importe (E12 directo)."""
    assert _fila(filas, "027001", "Enero")["cartera_e12"] == 300.0


def test_cartera_suma_las_etapas_uno_y_dos(filas):
    """Sin valor directo, E12 = etapa 1 + etapa 2."""
    assert _fila(filas, "027009", "Enero")["cartera_e12"] == 200.0


def test_cartera_resta_la_etapa_tres_al_total(filas):
    """Sin valor directo ni desglose, E12 = total − etapa 3."""
    assert _fila(filas, "027001", "Febrero")["cartera_e12"] == 360.0
    assert _fila(filas, "027001", "Febrero")["cartera_e3"] == 40.0


def test_tiene_datos_falso_en_meses_sin_registros(filas):
    """``027009`` solo tiene datos en enero; el resto de 2026 queda vacío."""
    assert _fila(filas, "027009", "Enero")["tiene_datos"] is True
    assert _fila(filas, "027009", "Febrero")["tiene_datos"] is False
    assert _fila(filas, "027001", "Agosto")["tiene_datos"] is True


def test_nombres_reemplaza_al_del_universo(filas):
    """La clave presente en ``nombres`` manda; la ausente usa el universo."""
    assert _fila(filas, "027001", "Enero")["sofipo"] == "Sociedad Uno Renombrada"
    assert _fila(filas, "027009", "Enero")["sofipo"] == "Sociedad Dos"


def test_enero_toma_los_totales_de_diciembre(filas):
    """La fila de enero 2026 usa la captación y cartera de 202512."""
    enero = _fila(filas, "027001", "Enero")
    assert enero["captacion_prev"] == 500.0
    assert enero["cartera_total_prev"] == 390.0
    assert enero["mes"] == "Enero"
    assert enero["trimestre"] == "T1"

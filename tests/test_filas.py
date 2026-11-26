"""Pruebas de preparación de filas mensuales usando un tablero de prueba mínimo."""

from pathlib import Path

import pytest

from sofipos.config import ANIOS
from sofipos.filas import preparar_filas
from sofipos.tablero import leer_tablero
from sofipos.utilidades import fecha_mes

FIXTURE = Path(__file__).parent / "fixtures" / "tablero_mini.csv"
FECHAS = {fecha_mes(2026, mes) for mes in ANIOS[2026]} | {"202512"}

# Universo y nombres definidos aquí para no depender del BE.
UNIVERSO = [
    {"clave": "027001", "consecutivo": 1, "nombre": "Sociedad Uno"},
    {"clave": "027009", "consecutivo": 2, "nombre": "Sociedad Dos"},
]
NOMBRES = {"027001": "Sociedad Uno Renombrada"}


@pytest.fixture(scope="module")
def filas():
    """Prepara una vez las filas de 2026 a partir del fixture mínimo."""
    tablero = leer_tablero(FIXTURE, FECHAS)
    return preparar_filas(tablero, UNIVERSO, NOMBRES, 2026)


def _fila(filas, clave, mes):
    """Busca y devuelve la fila de una sociedad para el mes indicado."""
    return next(f for f in filas if f["clave"] == clave and f["mes"] == mes)


def test_una_fila_por_sociedad_y_mes(filas):
    """Comprueba que haya una fila por sociedad y mes de 2026."""
    assert len(filas) == len(ANIOS[2026]) * len(UNIVERSO)


def test_importes_van_en_miles_con_dos_decimales(filas):
    """Comprueba la conversión de pesos a miles y el redondeo a dos decimales."""
    assert _fila(filas, "027001", "Enero")["activos"] == 1001.0
    assert _fila(filas, "027001", "Enero")["gastos_adm"] == 20.12


def test_cartera_usa_el_valor_directo(filas):
    """Usa el importe directo de cartera cuando está disponible."""
    assert _fila(filas, "027001", "Enero")["cartera_e12"] == 300.0


def test_cartera_suma_las_etapas_uno_y_dos(filas):
    """Calcula E12 sumando las etapas 1 y 2 si falta el valor directo."""
    assert _fila(filas, "027009", "Enero")["cartera_e12"] == 200.0


def test_cartera_resta_la_etapa_tres_al_total(filas):
    """Calcula E12 restando la etapa 3 al total si faltan otros datos."""
    assert _fila(filas, "027001", "Febrero")["cartera_e12"] == 360.0
    assert _fila(filas, "027001", "Febrero")["cartera_e3"] == 40.0


def test_tiene_datos_falso_en_meses_sin_registros(filas):
    """Marca como vacíos los meses sin registros para la sociedad indicada."""
    assert _fila(filas, "027009", "Enero")["tiene_datos"] is True
    assert _fila(filas, "027009", "Febrero")["tiene_datos"] is False
    assert _fila(filas, "027001", "Agosto")["tiene_datos"] is True


def test_nombres_reemplaza_al_del_universo(filas):
    """Usa el nombre configurado y, si falta, conserva el del universo."""
    assert _fila(filas, "027001", "Enero")["sofipo"] == "Sociedad Uno Renombrada"
    assert _fila(filas, "027009", "Enero")["sofipo"] == "Sociedad Dos"


def test_enero_toma_los_totales_de_diciembre(filas):
    """Usa los datos de diciembre de 2025 como valores previos de enero de 2026."""
    enero = _fila(filas, "027001", "Enero")
    assert enero["captacion_prev"] == 500.0
    assert enero["cartera_total_prev"] == 390.0
    assert enero["mes"] == "Enero"
    assert enero["trimestre"] == "T1"

"""Pruebas del selector de años ``parse_anios`` de la CLI."""

import pytest

from sofipos.config import ANIOS
from sofipos.main import parse_anios


def test_sin_texto_devuelve_todos_los_anios():
    """Devuelve la configuración completa cuando no se especifican años."""
    assert parse_anios(None) == ANIOS


def test_anio_unico():
    """Devuelve el año indicado con los meses definidos en la configuración."""
    assert parse_anios("2026") == {2026: ANIOS[2026]}


def test_rango_de_anios():
    """Devuelve todos los años, inclusive, dentro del rango indicado."""
    assert parse_anios("2024-2026") == {
        anio: ANIOS[anio] for anio in (2024, 2025, 2026)
    }


def test_anio_fuera_de_la_config():
    """Lanza ``SystemExit`` si el año no está definido en ``ANIOS``."""
    with pytest.raises(SystemExit):
        parse_anios("2027")


def test_formato_invalido():
    """Lanza ``SystemExit`` si el texto no tiene un formato de año válido."""
    with pytest.raises(SystemExit):
        parse_anios("dieciséis")

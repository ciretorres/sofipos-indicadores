"""Pruebas de las utilidades de fechas y conversión."""

from sofipos.utilidades import a_miles, fecha_anterior, fecha_mes


def test_fecha_mes_concatena_anio_y_mes():
    """Enero queda como ``202601`` con el mes de dos dígitos."""
    assert fecha_mes(2026, 1) == "202601"
    assert fecha_mes(2026, 12) == "202612"


def test_fecha_anterior_cruza_el_anio():
    """Enero regresa a diciembre del año anterior."""
    assert fecha_anterior("202601") == "202512"
    assert fecha_anterior("202608") == "202607"


def test_a_miles_convierte_pesos_a_miles():
    """Divide entre mil y redondea a dos decimales."""
    assert a_miles(1_000_000) == 1000.0
    assert a_miles(None) is None

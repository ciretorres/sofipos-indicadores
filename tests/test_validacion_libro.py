"""Pruebas de las hojas esperadas derivadas de los años de ``filas``."""

from sofipos.validacion_libro import hojas_esperadas


def test_hojas_esperadas_con_anios_no_contiguos_en_orden_descendente():
    """Dos años no contiguos generan sus hojas en orden descendente."""
    assert hojas_esperadas({2019: [], 2021: []}) == [
        "Catálogo SOFIPOS",
        "Datos Fuente 2021",
        "Datos Fuente 2019",
        "Indicadores 2021",
        "Indicadores 2019",
        "Diccionario",
        "Semáforos",
    ]


def test_hojas_esperadas_con_anio_unico():
    """Un solo año produce una hoja de datos y una de indicadores."""
    assert hojas_esperadas({2026: []}) == [
        "Catálogo SOFIPOS",
        "Datos Fuente 2026",
        "Indicadores 2026",
        "Diccionario",
        "Semáforos",
    ]


def test_hojas_esperadas_con_los_dos_anios_actuales():
    """La config actual produce las 7 hojas en el orden actual."""
    assert hojas_esperadas({2025: [], 2026: []}) == [
        "Catálogo SOFIPOS",
        "Datos Fuente 2026",
        "Datos Fuente 2025",
        "Indicadores 2026",
        "Indicadores 2025",
        "Diccionario",
        "Semáforos",
    ]

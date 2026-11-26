"""Pruebas de la configuración: conceptos y años."""

from sofipos.config import ANIOS, CONCEPTOS


def test_conceptos_claves_sin_duplicadas():
    """Ninguna clave de CONCEPTOS se repite."""
    assert len(CONCEPTOS) == len(set(CONCEPTOS))


def test_conceptos_tiene_claves_solo_digitos():
    """Toda clave de CONCEPTOS está compuesta únicamente por dígitos."""
    assert all(clave.isdigit() for clave in CONCEPTOS)


def test_conceptos_valores_unicos():
    """Ningún nombre de campo se repite."""
    valores = list(CONCEPTOS.values())
    assert len(valores) == len(set(valores))


def test_anios_son_enteros_con_meses_ordenados():
    """Cada año usa meses enteros de 1 a 12, en orden y sin repetir."""
    for anio, meses in ANIOS.items():
        assert isinstance(anio, int)
        assert list(meses) == sorted(set(meses))
        assert all(1 <= mes <= 12 for mes in meses)

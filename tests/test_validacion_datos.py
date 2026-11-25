"""Pruebas de tolerancia de diferencias en la validación de datos."""

from sofipos.tablero import Tablero
from sofipos.validacion_datos import _tolera_dif, validar_datos


def _fila(clave="027032", anio=2024, mes="Mayo", **campos):
    """Crea una fila mensual con los datos mínimos y los campos adicionales."""
    base = {"clave": clave, "anio": anio, "mes": mes}
    base.update(campos)
    return base


def _tablero(valores):
    """Crea un tablero mínimo con valores crudos y sin datos sectoriales."""
    return Tablero(valores=valores, sector={})


def test_dif_grande_con_publicado_enorme_tolerada():
    """Acepta una diferencia grande cuando el valor publicado también lo es."""
    assert _tolera_dif(125.75, 16_772_979.0)


def test_dif_pequena_con_publicado_pequeno_falla():
    """Rechaza una diferencia que supera las tolerancias absoluta y relativa."""
    assert not _tolera_dif(0.03, 1.0)


def test_dif_dentro_de_la_tolerancia_absoluta():
    """Acepta una diferencia dentro del margen absoluto de redondeo."""
    assert _tolera_dif(0.015, 1.0)


def test_nicap_dif_grande_tolerada_no_marca_falla():
    """La diferencia tolerada de NICAP no hace fallar la validación."""
    valores = {
        ("027032", "202405"): {
            "capital_neto": 16_773_104.75,
            "requerimiento": 100.0,
        }
    }
    tablero = _tablero(valores)
    filas = {2024: [_fila(nicap_pub=16_772_979.0)]}
    universo = [{"clave": "027032", "consecutivo": 1, "nombre": "Stori"}]

    assert validar_datos(tablero, filas, universo, {}) is True


def test_roa_roe_sigue_estricto_aun_con_publicado_grande():
    """La validación de ROA rechaza diferencias mayores al margen absoluto."""
    valores = {
        ("027032", "202405"): {
            "rn_12m": 1_000_000.03,
            "activo_prom_12m": 100.0,
        }
    }
    tablero = _tablero(valores)
    filas = {2024: [_fila(roa_pub=1_000_000.0)]}
    universo = [{"clave": "027032", "consecutivo": 1, "nombre": "Stori"}]

    assert validar_datos(tablero, filas, universo, {}) is False

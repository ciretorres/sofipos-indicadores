"""Prepara las filas mensuales del libro a partir de los datos del tablero."""

from __future__ import annotations

from .config import ANIOS, MESES
from .tablero import Tablero
from .utilidades import a_miles, fecha_anterior, fecha_mes, trimestre


def _cartera_e12(v: dict) -> float | None:
    """Obtiene la cartera de etapas 1 y 2, usando datos alternativos si hace falta.

    Primero usa el valor directo. Si no está disponible, suma las etapas 1 y 2
    o calcula la diferencia entre la cartera total y la etapa 3.

    Args:
        v: Valores del tablero para una sociedad y una fecha.

    Returns:
        El importe de la cartera de etapas 1 y 2, o ``None`` si no se puede
        calcular.
    """
    if v.get("cartera_e12") is not None:
        return v["cartera_e12"]
    if v.get("cartera_e1") is not None and v.get("cartera_e2") is not None:
        return v["cartera_e1"] + v["cartera_e2"]
    if (
        v.get("cartera_total_cruda") is not None
        and v.get("cartera_e3") is not None
    ):
        return v["cartera_total_cruda"] - v["cartera_e3"]
    return None


def _cartera_e3(v: dict, e12: float | None) -> float | None:
    """Obtiene la cartera de etapa 3, calculándola si no viene desglosada.

    Args:
        v: Valores del tablero para una sociedad y una fecha.
        e12: Cartera de etapas 1 y 2, si está disponible.

    Returns:
        El importe de la cartera de etapa 3, o ``None`` si no se puede
        calcular.
    """
    if v.get("cartera_e3") is not None:
        return v["cartera_e3"]
    if v.get("cartera_total_cruda") is not None and e12 is not None:
        return v["cartera_total_cruda"] - e12
    return None


def preparar_filas(
    tablero: Tablero,
    universo: list[dict],
    nombres: dict[str, str],
    anio: int,
) -> list[dict]:
    """Prepara los registros mensuales de las sociedades para un año.

    Combina los datos del tablero con la información del universo, calcula
    los valores de cartera que falten y convierte los importes a miles.

    Args:
        tablero: Datos del tablero indexados por clave y fecha.
        universo: Sociedades que se incluirán en el libro.
        nombres: Nombres de las sociedades indexados por clave.
        anio: Año para el que se prepararán las filas.

    Returns:
        Lista de registros, uno por sociedad y mes, con los datos necesarios
        para generar el libro.
    """
    filas: list[dict] = []

    for mes in ANIOS[anio]:
        fecha = fecha_mes(anio, mes)
        previa = fecha_anterior(fecha)

        for reg in universo:
            clave = reg["clave"]
            v = tablero.valores.get((clave, fecha), {})
            p = tablero.valores.get((clave, previa), {})

            # Calcula la cartera del mes actual y la del mes anterior.
            e12 = _cartera_e12(v)
            e3 = _cartera_e3(v, e12)
            e12_prev = _cartera_e12(p)
            e3_prev = _cartera_e3(p, e12_prev)
            cartera_prev = (
                e12_prev + e3_prev
                if e12_prev is not None and e3_prev is not None
                else None
            )

            filas.append(
                {
                    "no": reg["consecutivo"],
                    "clave": clave,
                    "sofipo": nombres.get(clave) or reg["nombre"],
                    "federacion": reg.get("federacion") or "Sin federación",
                    "nivel": reg.get("nivel") or "Sin nivel",
                    "anio": anio,
                    "mes": MESES[mes],
                    "trimestre": trimestre(mes),
                    "activos": a_miles(v.get("activos")),
                    "cartera_e12": a_miles(e12),
                    "cartera_e3": a_miles(e3),
                    "estimaciones": a_miles(v.get("estimaciones")),
                    "ingresos_int": a_miles(v.get("ingresos_int")),
                    "gastos_int": a_miles(v.get("gastos_int")),
                    "gastos_adm": a_miles(v.get("gastos_adm")),
                    "resultado_neto": a_miles(v.get("resultado_neto")),
                    "captacion": a_miles(v.get("captacion")),
                    "captacion_prev": a_miles(p.get("captacion")),
                    "cartera_total_prev": a_miles(cartera_prev),
                    "capital_neto": a_miles(v.get("capital_neto")),
                    "requerimiento": a_miles(v.get("requerimiento")),
                    "rn_12m": a_miles(v.get("rn_12m")),
                    "activo_prom_12m": a_miles(v.get("activo_prom_12m")),
                    "capital_prom_12m": a_miles(v.get("capital_prom_12m")),
                    # Indicadores publicados, incluidos solo para validación.
                    "nicap_pub": v.get("nicap_pub"),
                    "imor_pub": v.get("imor_pub"),
                    "icor_pub": v.get("icor_pub"),
                    "roa_pub": v.get("roa_pub"),
                    "roe_pub": v.get("roe_pub"),
                    "cartera_total_cruda": a_miles(
                        v.get("cartera_total_cruda")
                    ),
                    "tiene_datos": bool(v),
                }
            )

    return filas

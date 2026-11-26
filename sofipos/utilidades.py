"""Funciones auxiliares para normalizar y convertir datos."""

from __future__ import annotations

import re
import unicodedata

from .config import MILES


def _a_numero(texto):
    """Convierte un valor a número flotante.

    Acepta números o textos con comas y símbolo de dólar. Devuelve ``None``
    si el valor está vacío o no se puede convertir.
    """
    if texto is None:
        return None
    if isinstance(texto, (int, float)):
        return float(texto)
    try:
        return float(str(texto).replace(",", "").replace("$", "").strip())
    except ValueError:
        return None


def _norm(texto) -> str:
    """Normaliza texto para compararlo sin acentos ni diferencias de formato."""
    s = unicodedata.normalize("NFKD", str(texto or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def normalizar_clave(valor) -> str:
    """Deja solo los dígitos de una clave y completa con ceros hasta seis."""
    digitos = re.sub(r"\D", "", str(valor))
    return digitos.zfill(6)


def normalizar_entidad_legal(nombre: str) -> str:
    """Normaliza el sufijo societario y conserva el nombre registral.

    Unifica las variantes de «Sociedad Financiera Popular» como «S.F.P.» y
    ajusta el formato de «S.A. de C.V.».
    """
    s = re.sub(r"\s+", " ", str(nombre)).strip().rstrip(",")
    s = re.sub(
        r",?\s*Sociedad\s+Financiera\s+Popular\.?$",
        "",
        s,
        flags=re.I,
    )
    s = re.sub(r",?\s*S\.\s*F\.\s*P\.?$", "", s, flags=re.I)
    s = s.strip().rstrip(",")

    if re.search(r"S\.A\.\s*de\s*C\.\s*V\.?$", s, re.I):
        s = re.sub(
            r"S\.A\.\s*de\s*C\.\s*V\.?$",
            "S.A. de C.V.",
            s,
            flags=re.I,
        )
    else:
        s = f"{s}, S.A. de C.V."

    return f"{s}, S.F.P."


def trimestre(mes: int) -> str:
    """Devuelve el trimestre correspondiente al número de mes."""
    return f"T{((mes - 1) // 3) + 1}"


def fecha_mes(anio: int, mes: int) -> str:
    """Devuelve el año y el mes concatenados en formato ``AAAAMM``."""
    return f"{anio}{mes:02d}"


def fecha_anterior(fecha: str) -> str:
    """Devuelve el mes anterior a una fecha en formato ``AAAAMM``."""
    anio, mes = int(fecha[:4]), int(fecha[4:])
    mes -= 1
    if mes == 0:
        anio, mes = anio - 1, 12
    return f"{anio}{mes:02d}"


def enumerar(valores) -> str:
    """Une valores con comas y «e» antes del último (2 o más)."""
    v = list(valores)
    if not v:
        return ""
    if len(v) == 1:
        return v[0]
    return ", ".join(v[:-1]) + " e " + v[-1]


def a_miles(valor):
    """Convierte pesos a miles de pesos y redondea a dos decimales.

    Devuelve ``None`` si el valor de entrada es ``None``.
    """
    if valor is None:
        return None
    return round(float(valor) / MILES, 2)

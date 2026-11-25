"""Genera y valida el libro Excel mensual de las SOFIPOS."""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from .be import leer_bajas, leer_be_sofipos, leer_universo
from .catalogo import CATALOGO
from .config import (
    ANIOS,
    BE_UNIVERSO,
    BE_VALIDACION,
    CSV_TABLERO,
    DIR_ENTRADAS,
    DIR_SALIDAS,
    SALIDA_DEFECTO,
)
from .filas import preparar_filas
from .hojas import escribir_xlsx
from .tablero import leer_tablero
from .utilidades import fecha_anterior, fecha_mes
from .validacion_datos import validar_datos
from .validacion_libro import validar_libro

# --------------------------------------------------------------------------
# Principal
# --------------------------------------------------------------------------

def parse_anios(texto: str | None) -> dict[int, tuple[int, ...]]:
    """Selecciona los años configurados a partir de una opción de rango.

    Acepta un año (``AAAA``), un rango inclusivo (``AAAA-AAAA``) o ``None``
    para usar todos los años configurados.

    Args:
        texto: Valor recibido mediante ``--anios``.

    Returns:
        Diccionario con los meses configurados para cada año seleccionado.

    Raises:
        SystemExit: Si el formato es inválido, el rango es descendente o
            incluye años no configurados.
    """
    if texto is None:
        return ANIOS

    rango = re.fullmatch(r"(\d{4})(?:-(\d{4}))?", texto)
    if not rango:
        raise SystemExit(f"--anios inválido: {texto!r} (usa AAAA o AAAA-AAAA)")

    inicio = int(rango.group(1))
    fin = int(rango.group(2) or inicio)
    if inicio > fin:
        raise SystemExit(f"--anios inválido: {texto!r} (rango descendente)")

    elegidos = range(inicio, fin + 1)
    faltantes = [a for a in elegidos if a not in ANIOS]
    if faltantes:
        raise SystemExit(
            f"--anios: año no contemplado {faltantes[0]} "
            f"(disponibles: {', '.join(str(a) for a in sorted(ANIOS))})"
        )

    return {a: ANIOS[a] for a in elegidos}


def main() -> int:
    """Genera el libro Excel y ejecuta sus validaciones.

    Lee las entradas, prepara las filas, escribe el archivo y valida los
    datos y el libro generado.

    Returns:
        ``0`` si las validaciones pasan; ``2`` si hay observaciones.

    Raises:
        SystemExit: Si falta la carpeta de entradas, algún archivo requerido
            o datos de catálogo para una sociedad.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Genera el Excel mensual de las 35 SOFIPOS "
            f"({', '.join(str(a) for a in sorted(ANIOS))})."
        )
    )
    parser.add_argument(
        "--entradas",
        type=Path,
        default=DIR_ENTRADAS,
        help="Carpeta con el BE del Universo y la carpeta tablero/",
    )
    parser.add_argument(
        "--salidas",
        type=Path,
        default=DIR_SALIDAS,
        help="Carpeta donde se escribe el .xlsx",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Ruta exacta del .xlsx (por defecto salidas/SOFIPOS_Indicadores_Datos.xlsx)",
    )
    parser.add_argument(
        "--tablero",
        type=Path,
        default=None,
        help="Ruta exacta del CSV 027_datos.csv",
    )
    parser.add_argument(
        "--anios",
        default=None,
        help="Años a incluir: AAAA o AAAA-AAAA (por defecto, todos los configurados)",
    )
    args = parser.parse_args()

    anios_sel = parse_anios(args.anios)

    # Comprueba que exista la carpeta de entradas.
    dir_entradas = args.entradas.expanduser().resolve()
    if not dir_entradas.is_dir():
        raise SystemExit(f"No existe la carpeta de entradas: {dir_entradas}")

    # Usa la ruta indicada para el CSV o, si no se especificó, la ruta por defecto.
    ruta_csv = (
        args.tablero.expanduser().resolve()
        if args.tablero
        else dir_entradas / CSV_TABLERO
    )
    ruta_universo = dir_entradas / BE_UNIVERSO

    # Informa si faltan archivos requeridos y muestra el contenido disponible.
    faltantes = [str(ruta) for ruta in (ruta_csv, ruta_universo) if not ruta.exists()]
    if faltantes:
        disponibles = sorted(
            archivo.name for archivo in dir_entradas.iterdir() if archivo.is_file()
        )
        raise SystemExit(
            "Faltan archivos:\n  - "
            + "\n  - ".join(faltantes)
            + "\nContenido actual de entradas/:\n  - "
            + ("\n  - ".join(disponibles) or "(vacía)")
        )

    # Define la ruta de salida, usando el nombre predeterminado si hace falta.
    salida = args.out or (
        args.salidas.expanduser().resolve() / SALIDA_DEFECTO.name
    )

    # Lee el universo y los nombres de las sociedades desde el archivo BE.
    universo = leer_universo(ruta_universo)
    # Las bajas (transformadas o revocadas) conservan su histórico en el
    # libro: se agregan al universo de datos con consecutivo después de las
    # sociedades en operación.
    bajas = leer_bajas(ruta_universo)
    for i, baja in enumerate(bajas, start=1):
        baja["consecutivo"] = len(universo) + i
    universo_datos = universo + bajas
    be_principal = leer_be_sofipos(ruta_universo)
    nombres = {registro["clave"]: registro["nombre"] for registro in be_principal.registros}

    print(f"Entradas: {dir_entradas}")
    print(
        f"Universo: {len(universo)} activas · +{len(bajas)} bajas · "
        f"{ruta_universo.name}"
    )

    # Verifica que todas las sociedades del universo estén en el catálogo.
    sin_catalogo = [sociedad["clave"] for sociedad in universo if sociedad["clave"] not in CATALOGO]
    if sin_catalogo:
        raise SystemExit(f"Faltan datos de catálogo para: {sin_catalogo}")

    # Reúne los meses del libro y el mes anterior necesario para los cálculos.
    fechas: set[str] = set()
    for anio, meses in anios_sel.items():
        for mes in meses:
            fecha = fecha_mes(anio, mes)
            fechas.add(fecha)
            fechas.add(fecha_anterior(fecha))

    # Lee del tablero los datos correspondientes a las fechas necesarias.
    print(f"\nLeyendo tablero: {ruta_csv.name} ...")
    tablero = leer_tablero(ruta_csv, fechas)

    # Prepara las filas del libro para cada año seleccionado.
    filas = {
        anio: preparar_filas(tablero, universo_datos, nombres, anio)
        for anio in anios_sel
    }
    print(
        "  filas preparadas: "
        + " · ".join(f"{anio}: {len(registros)}" for anio, registros in filas.items())
    )

    # Carga los archivos BE disponibles para validar los datos.
    be_por_archivo = {}
    for nombre_archivo in BE_VALIDACION:
        ruta = dir_entradas / nombre_archivo
        if ruta.exists():
            be_por_archivo[nombre_archivo] = leer_be_sofipos(ruta)

    # Valida los datos, genera el archivo y valida el libro resultante.
    todo_ok = validar_datos(tablero, filas, universo_datos, be_por_archivo)

    escribir_xlsx(universo, bajas, nombres, filas, salida)
    todo_ok &= validar_libro(salida, universo, filas)

    print(f"\nArchivo generado: {salida}")
    print("Validación general:", "OK" if todo_ok else "CON OBSERVACIONES")
    return 0 if todo_ok else 2

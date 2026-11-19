"""Genera y valida el libro Excel mensual de las SOFIPOS."""

from __future__ import annotations

import argparse
from pathlib import Path

from .be import leer_be_sofipos, leer_universo
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

def main() -> int:
    """Genera el Excel, valida los datos y devuelve el código de salida.

    Returns:
        0 si todas las validaciones son correctas; 2 si hay observaciones.
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
        help="Ruta exacta del .xlsx (por defecto salidas/SOFIPOS_2026_Datos.xlsx)",
    )
    parser.add_argument(
        "--tablero",
        type=Path,
        default=None,
        help="Ruta exacta del CSV 027_datos.csv",
    )
    args = parser.parse_args()

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
    be_principal = leer_be_sofipos(ruta_universo)
    nombres = {registro["clave"]: registro["nombre"] for registro in be_principal.registros}

    print(f"Entradas: {dir_entradas}")
    print(f"Universo: {len(universo)} sociedades · {ruta_universo.name}")

    # Verifica que todas las sociedades del universo estén en el catálogo.
    sin_catalogo = [sociedad["clave"] for sociedad in universo if sociedad["clave"] not in CATALOGO]
    if sin_catalogo:
        raise SystemExit(f"Faltan datos de catálogo para: {sin_catalogo}")

    # Reúne los meses del libro y el mes anterior necesario para los cálculos.
    fechas: set[str] = set()
    for anio, meses in ANIOS.items():
        for mes in meses:
            fecha = fecha_mes(anio, mes)
            fechas.add(fecha)
            fechas.add(fecha_anterior(fecha))

    # Lee del tablero los datos correspondientes a las fechas necesarias.
    print(f"\nLeyendo tablero: {ruta_csv.name} ...")
    tablero = leer_tablero(ruta_csv, fechas)

    # Prepara las filas del libro para cada año configurado.
    filas = {
        anio: preparar_filas(tablero, universo, nombres, anio)
        for anio in ANIOS
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
    todo_ok = validar_datos(tablero, filas, universo, be_por_archivo)

    escribir_xlsx(universo, nombres, filas, salida)
    todo_ok &= validar_libro(salida, universo, filas)

    print(f"\nArchivo generado: {salida}")
    print("Validación general:", "OK" if todo_ok else "CON OBSERVACIONES")
    return 0 if todo_ok else 2

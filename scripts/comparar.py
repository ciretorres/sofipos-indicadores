#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Compara dos libros .xlsx y, opcionalmente, sus salidas de consola.

Ejemplos:
    .venv/bin/python scripts/comparar.py --xlsx-a viejo.xlsx --xlsx-b nuevo.xlsx
    .venv/bin/python scripts/comparar.py --xlsx-a viejo.xlsx --xlsx-b nuevo.xlsx \\
        --stdout-a viejo.txt --stdout-b nuevo.txt

Compara el contenido y la estructura, no los bytes del archivo. Los archivos
.xlsx pueden diferir en bytes por los timestamps que incluye su contenedor ZIP.

Se revisan las hojas, sus celdas y formatos, las propiedades del libro y otros
elementos de configuración de página y formato condicional.

Códigos de salida:
    0: los archivos comparados son idénticos.
    1: se encontraron diferencias.
    2: el uso es incorrecto o falta algún archivo.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook


# Máximo de diferencias detalladas que se muestran en pantalla.
MAX_REPORTE = 40

# --------------------------------------------------------------------------
# Utilidades de comparación
# --------------------------------------------------------------------------

def _color(color) -> str:
    """Representa un color de openpyxl como texto comparable."""
    if color is None:
        return "None"

    tipo = color.type
    valor = (
        color.rgb
        if tipo == "rgb"
        else color.indexed
        if tipo == "indexed"
        else color.theme
        if tipo == "theme"
        else color.auto
    )
    return f"{tipo}:{valor}"


def _fuente(fuente) -> tuple:
    """Extrae las propiedades de fuente que se comparan."""
    if fuente is None:
        return ()

    return (
        fuente.name,
        fuente.sz,
        fuente.b,
        fuente.i,
        fuente.u,
        _color(fuente.color),
    )


def _relleno(relleno) -> tuple:
    """Extrae el tipo de relleno y sus colores."""
    if relleno is None:
        return ()

    return (
        relleno.patternType,
        _color(relleno.fgColor),
        _color(relleno.bgColor),
    )


def _borde(borde) -> tuple:
    """Extrae los colores de los bordes izquierdo, derecho, superior e inferior."""
    if borde is None:
        return ()

    return tuple(
        _color(getattr(borde, lado).color)
        if getattr(borde, lado)
        else "None"
        for lado in ("left", "right", "top", "bottom")
    )


def _alineacion(al) -> tuple:
    """Extrae las propiedades principales de alineación."""
    if al is None:
        return ()

    return (al.horizontal, al.vertical, al.wrap_text, al.indent)


def _celda(celda) -> tuple:
    """Devuelve los valores y formatos comparables de una celda."""
    enlace = celda.hyperlink

    return (
        repr(celda.value),
        celda.data_type,
        celda.number_format,
        _fuente(celda.font),
        _relleno(celda.fill),
        _borde(celda.border),
        _alineacion(celda.alignment),
        (enlace.target, enlace.location, enlace.tooltip) if enlace else None,
    )


def _reglas_formato_condicional(ws) -> list:
    """Devuelve las reglas de formato condicional de una hoja."""
    reglas = []

    for rango in ws.conditional_formatting:
        for regla in rango.rules:
            dxf = regla.dxf
            reglas.append(
                (
                    str(rango.sqref),
                    regla.type,
                    tuple(regla.formula or ()),
                    regla.operator,
                    bool(regla.stopIfTrue),
                    _relleno(dxf.fill) if dxf is not None and dxf.fill else (),
                    _fuente(dxf.font) if dxf is not None and dxf.font else (),
                )
            )

    return sorted(reglas, key=repr)


def _hoja(ws) -> dict:
    """Recopila el contenido y la configuración comparable de una hoja."""
    return {
        "max_row": ws.max_row,
        "max_column": ws.max_column,
        "celdas": {
            (fila, col): _celda(ws.cell(row=fila, column=col))
            for fila in range(1, ws.max_row + 1)
            for col in range(1, ws.max_column + 1)
        },
        "combinadas": sorted(str(r) for r in ws.merged_cells.ranges),
        "anchos": {
            k: (v.width, v.customWidth)
            for k, v in ws.column_dimensions.items()
        },
        "altos": {
            k: v.height
            for k, v in ws.row_dimensions.items()
            if v.height
        },
        "freeze": str(ws.freeze_panes) if ws.freeze_panes else None,
        "autofiltro": ws.auto_filter.ref,
        "gridlines": ws.sheet_view.showGridLines,
        "orientacion": ws.page_setup.orientation,
        "fitToWidth": ws.page_setup.fitToWidth,
        "fitToHeight": ws.page_setup.fitToHeight,
        "fitToPage": bool(ws.sheet_properties.pageSetUpPr.fitToPage),
        "print_title_rows": ws.print_title_rows,
        "condicional": _reglas_formato_condicional(ws),
    }


def _libro(ruta: Path) -> dict:
    """Abre un libro y recopila sus propiedades y el contenido de sus hojas."""
    wb = load_workbook(ruta)

    return {
        "hojas": list(wb.sheetnames),
        "propiedades": (
            wb.properties.title,
            wb.properties.subject,
            wb.properties.creator,
        ),
        "fullCalcOnLoad": bool(wb.calculation.fullCalcOnLoad),
        "contenido": {
            nombre: _hoja(wb[nombre])
            for nombre in wb.sheetnames
        },
    }

# --------------------------------------------------------------------------
# Reporte
# --------------------------------------------------------------------------

def comparar_libros(ruta_a: Path, ruta_b: Path) -> list[str]:
    """Compara dos libros y devuelve sus diferencias.

    Args:
        ruta_a: Ruta del libro de referencia.
        ruta_b: Ruta del libro que se compara con la referencia.

    Returns:
        Lista de diferencias; queda vacía si los libros coinciden.
    """
    a, b = _libro(ruta_a), _libro(ruta_b)
    diferencias: list[str] = []

    if a["hojas"] != b["hojas"]:
        diferencias.append(
            f"orden/nombre de hojas: {a['hojas']} != {b['hojas']}"
        )

    if a["propiedades"] != b["propiedades"]:
        diferencias.append(
            f"propiedades del libro: {a['propiedades']} != {b['propiedades']}"
        )

    if a["fullCalcOnLoad"] != b["fullCalcOnLoad"]:
        diferencias.append(
            "fullCalcOnLoad: "
            f"{a['fullCalcOnLoad']} != {b['fullCalcOnLoad']}"
        )

    # Compara las propiedades y las celdas de cada hoja presente en ambos libros.
    for nombre in a["contenido"]:
        if nombre not in b["contenido"]:
            continue

        hoja_a = a["contenido"][nombre]
        hoja_b = b["contenido"][nombre]

        for clave in hoja_a:
            if clave == "celdas":
                continue
            if hoja_a[clave] != hoja_b[clave]:
                diferencias.append(
                    f"[{nombre}] {clave}: "
                    f"{hoja_a[clave]!r} != {hoja_b[clave]!r}"
                )

        celdas_a = hoja_a["celdas"]
        celdas_b = hoja_b["celdas"]

        for ref in sorted(set(celdas_a) | set(celdas_b)):
            if celdas_a.get(ref) != celdas_b.get(ref):
                fila, col = ref
                diferencias.append(
                    f"[{nombre}] fila {fila} columna {col}: "
                    f"{celdas_a.get(ref)!r} != {celdas_b.get(ref)!r}"
                )

    return diferencias


def comparar_stdout(ruta_a: Path, ruta_b: Path) -> list[str]:
    """Compara dos archivos de salida de consola línea por línea."""
    lineas_a = ruta_a.read_text(encoding="utf-8").splitlines()
    lineas_b = ruta_b.read_text(encoding="utf-8").splitlines()

    if lineas_a == lineas_b:
        return []

    diferencias = [
        f"stdout: {len(lineas_a)} líneas != {len(lineas_b)} líneas"
    ]

    for i in range(max(len(lineas_a), len(lineas_b))):
        la = lineas_a[i] if i < len(lineas_a) else "<sin línea>"
        lb = lineas_b[i] if i < len(lineas_b) else "<sin línea>"

        if la != lb:
            diferencias.append(f"  línea {i + 1}:")
            diferencias.append(f"    A: {la}")
            diferencias.append(f"    B: {lb}")

    return diferencias


def main() -> int:
    """Procesa los argumentos, compara los archivos y muestra el resultado."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--xlsx-a",
        type=Path,
        required=True,
        help="Libro de referencia",
    )
    parser.add_argument(
        "--xlsx-b",
        type=Path,
        required=True,
        help="Libro a comparar",
    )
    parser.add_argument(
        "--stdout-a",
        type=Path,
        default=None,
        help="Salida de consola de referencia (opcional)",
    )
    parser.add_argument(
        "--stdout-b",
        type=Path,
        default=None,
        help="Salida de consola a comparar (opcional)",
    )
    args = parser.parse_args()

    # Comprueba que existan todos los archivos indicados.
    rutas = [args.xlsx_a, args.xlsx_b]
    rutas += [p for p in (args.stdout_a, args.stdout_b) if p]
    faltantes = [str(ruta) for ruta in rutas if not ruta.exists()]

    if faltantes:
        print("No existen: " + ", ".join(faltantes), file=sys.stderr)
        return 2

    # Las salidas de consola deben indicarse en pareja.
    if (args.stdout_a is None) != (args.stdout_b is None):
        print(
            "Hay que indicar --stdout-a y --stdout-b juntos.",
            file=sys.stderr,
        )
        return 2

    diferencias = comparar_libros(args.xlsx_a, args.xlsx_b)

    if args.stdout_a:
        diferencias += comparar_stdout(args.stdout_a, args.stdout_b)

    if not diferencias:
        mensaje = "IDÉNTICO: misma estructura de xlsx"
        if args.stdout_a:
            mensaje += " y mismo stdout"
        print(mensaje + ".")
        return 0

    print(f"{len(diferencias)} diferencias:")
    for linea in diferencias[:MAX_REPORTE]:
        print("  " + linea)

    if len(diferencias) > MAX_REPORTE:
        print(f"  ... y {len(diferencias) - MAX_REPORTE} más")

    return 1


if __name__ == "__main__":
    sys.exit(main())

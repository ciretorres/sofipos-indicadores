"""Funciones para crear la hoja del catálogo de SOFIPOS."""

from __future__ import annotations

import calendar
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from ..catalogo import CATALOGO
from ..config import BE_UNIVERSO, MESES
from ..estilos import (
    AZUL,
    ENC_CATALOGO,
    FILA_DATOS,
    FILA_ENC,
    FMT_FECHA,
    _borde_y_banda,
    _celda_numero,
    _estilo_encabezado,
    _notas,
    _texto,
    _titulo,
)
from ..utilidades import normalizar_entidad_legal


def _corte_be() -> tuple[int, int, str]:
    """Obtiene el año, mes y texto del corte indicado en ``BE_UNIVERSO``.

    Por ejemplo, ``BE_SOFIPOS_202609.xlsx`` devuelve
    ``(2026, 9, "30 de septiembre de 2026")``.
    """
    anio, mes = (int(x) for x in re.search(r"(\d{4})(\d{2})", BE_UNIVERSO).groups())
    ultimo = calendar.monthrange(anio, mes)[1]
    return anio, mes, f"{ultimo} de {MESES[mes].lower()} de {anio}"


def escribir_catalogo(
    wb: Workbook,
    universo: list[dict],
    nombres: dict[str, str],
    bajas: list[dict],
) -> None:
    """Agrega al libro la hoja «Catálogo SOFIPOS».

    Incluye las sociedades en operación y un apartado para las transformadas
    o revocadas, con formato y notas sobre las fuentes de información.

    Args:
        wb: Libro de Excel al que se agrega la hoja.
        universo: Registros de las sociedades en operación.
        nombres: Nombres actualizados asociados a cada clave CASFIM.
        bajas: Registros de sociedades transformadas o revocadas, con clave,
            nombre, motivo y fecha.
    """
    ws = wb.create_sheet("Catálogo SOFIPOS")
    ncols = len(ENC_CATALOGO)
    anio, mes, corte = _corte_be()

    _titulo(
        ws,
        ncols,
        "SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Catálogo",
        f"{len(universo)} sociedades autorizadas en operación · "
        f"Información al {corte}",
        f"Autorización, nivel y estatus: {BE_UNIVERSO}, hoja Universo. "
        "Entidad Legal y Sitio Web: investigación en CONDUSEF/SIPRES y SEGOB/DOF.",
    )

    # Escribe y da formato a los encabezados.
    for j, encabezado in enumerate(ENC_CATALOGO, start=1):
        ws.cell(row=FILA_ENC, column=j, value=encabezado)
    _estilo_encabezado(ws, FILA_ENC, ncols, alto=40)

    # Escribe los datos de cada sociedad.
    for i, reg in enumerate(universo):
        r = FILA_DATOS + i
        crudo, sitio = CATALOGO[reg["clave"]]

        _celda_numero(ws, r, 1, reg["consecutivo"], "0")
        _texto(ws, r, 2, reg["clave"], alineacion="center")
        _texto(ws, r, 3, nombres.get(reg["clave"], reg["nombre"]), negrita=True)
        _texto(ws, r, 4, reg["federacion"])
        _texto(ws, r, 5, normalizar_entidad_legal(crudo))
        _texto(ws, r, 6, sitio, color="0563C1")

        fecha = ws.cell(row=r, column=7, value=reg["fecha_autorizacion"])
        fecha.number_format = FMT_FECHA
        fecha.alignment = Alignment(horizontal="center", vertical="center")
        fecha.font = Font()

        _texto(ws, r, 8, reg["nivel"], alineacion="center")
        _texto(ws, r, 9, reg["estatus"], alineacion="center")

        # Convierte el sitio web en un hipervínculo.
        enlace = ws.cell(row=r, column=6)
        enlace.hyperlink = sitio
        enlace.font = Font(color="0563C1", underline="single", size=9)

        _borde_y_banda(ws, r, ncols, i % 2 == 1)

    ultima = FILA_DATOS + len(universo) - 1

    # Apartado: sociedades que salieron del universo de operación.
    encabezado = ["Clave CASFIM", "Nombre", "Motivo", "Fecha"]
    columnas_baja = [2, 3, 4, 7]  # Misma columna de fecha que la tabla.

    fila = ultima + 2
    ws.merge_cells(
        start_row=fila, start_column=1, end_row=fila, end_column=ncols
    )
    titulo = ws.cell(
        row=fila,
        column=1,
        value="Sociedades transformadas o revocadas",
    )
    titulo.font = Font(bold=True, size=11, color=AZUL)
    titulo.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[fila].height = 20

    fila += 1
    _estilo_encabezado(ws, fila, ncols, alto=20)
    for col, texto in zip(columnas_baja, encabezado):
        ws.cell(row=fila, column=col, value=texto)

    for i, baja in enumerate(bajas):
        r = fila + 1 + i
        _texto(ws, r, 2, baja["clave"], alineacion="center")
        _texto(ws, r, 3, baja["nombre"], negrita=True)
        _texto(ws, r, 4, baja["motivo"])
        fecha = ws.cell(row=r, column=7, value=baja["fecha"])
        fecha.number_format = FMT_FECHA
        fecha.alignment = Alignment(horizontal="center", vertical="center")
        fecha.font = Font()
        _borde_y_banda(ws, r, ncols, i % 2 == 1)

    # Agrega notas sobre las fuentes y los criterios de elaboración.
    notas = [
        f"Fuente del catálogo: {BE_UNIVERSO}, hoja Universo (No., Clave CASFIM, SOFIPO, "
        "Federación, Fecha de autorización, Nivel de operación y Estatus).",
        "Las cifras financieras NO se toman del boletín: provienen del tablero mensual de la CNBV "
        "(entradas/tablero/sh_datos_csv_27_i/027_datos.csv), que se describe en las hojas de datos.",
        "Entidad Legal y Sitio Web investigados el 11 de noviembre de 2026 en el registro de la "
        "CONDUSEF/SIPRES, cuya clave registral es la misma Clave CASFIM; para Nu México se usó el "
        "oficio de SEGOB/DOF y su sitio oficial.",
        "Los sufijos de las denominaciones sociales se normalizaron a ', S.A. de C.V., S.F.P.'; "
        "no se alteró ninguna otra parte del nombre registral.",
        "Sitios activos que rechazan peticiones automáticas (HTTP 403): nu.com.mx y premo.mx.",
        f"La Federación corresponde a la clasificación vigente en el boletín de "
        f"{MESES[mes].lower()} de {anio}.",
        "El apartado inferior lista las sociedades transformadas o revocadas que dejaron de "
        f"estar en operación según {BE_UNIVERSO}; se conservan en el libro con sus datos "
        "históricos hasta cerrar 2026.",
    ]
    _notas(ws, ncols, fila + len(bajas) + 2, notas)

    # Configura el ancho de las columnas y la presentación de la hoja.
    anchos = {
        "A": 6,
        "B": 13,
        "C": 24,
        "D": 20,
        "E": 58,
        "F": 48,
        "G": 15,
        "H": 7,
        "I": 14,
    }
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho

    ws.freeze_panes = f"C{FILA_DATOS}"
    ws.auto_filter.ref = f"A{FILA_ENC}:{get_column_letter(ncols)}{ultima}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{FILA_ENC}:{FILA_ENC}"

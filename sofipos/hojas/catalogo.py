"""Funciones para crear la hoja del catálogo de SOFIPOS."""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from ..catalogo import CATALOGO
from ..estilos import (
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


def escribir_catalogo(
    wb: Workbook,
    universo: list[dict],
    nombres: dict[str, str],
) -> None:
    """Crea y configura la hoja «Catálogo SOFIPOS» del libro.

    Añade los datos de cada sociedad, aplica formato a la tabla e incorpora
    notas sobre las fuentes de información.

    Args:
        wb: Libro de Excel al que se agregará la hoja.
        universo: Registros de las sociedades que formarán parte del catálogo.
        nombres: Nombres de las sociedades asociados a su clave CASFIM.
    """
    ws = wb.create_sheet("Catálogo SOFIPOS")
    ncols = len(ENC_CATALOGO)

    _titulo(
        ws,
        ncols,
        "SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Catálogo",
        "35 sociedades autorizadas en operación · Información al 30 de junio de 2026",
        "Autorización, nivel y estatus: BE_SOFIPOS_202606.xlsx, hoja Universo. "
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

    # Agrega notas sobre las fuentes y los criterios de elaboración.
    notas = [
        "Fuente del catálogo: BE_SOFIPOS_202606.xlsx, hoja Universo (No., Clave CASFIM, SOFIPO, "
        "Federación, Fecha de autorización, Nivel de operación y Estatus).",
        "Las cifras financieras NO se toman del boletín: provienen del tablero mensual de la CNBV "
        "(entradas/tablero/sh_datos_csv_27_i/027_datos.csv), que se describe en las hojas de datos.",
        "Entidad Legal y Sitio Web investigados el 11 de noviembre de 2026 en el registro de la "
        "CONDUSEF/SIPRES, cuya clave registral es la misma Clave CASFIM; para Nu México se usó el "
        "oficio de SEGOB/DOF y su sitio oficial.",
        "Los sufijos de las denominaciones sociales se normalizaron a ', S.A. de C.V., S.F.P.'; "
        "no se alteró ninguna otra parte del nombre registral.",
        "Sitios activos que rechazan peticiones automáticas (HTTP 403): nu.com.mx y premo.mx.",
        "La Federación corresponde a la clasificación vigente en el boletín de junio de 2026.",
    ]
    _notas(ws, ncols, ultima + 2, notas)

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

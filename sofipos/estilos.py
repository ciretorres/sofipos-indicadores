"""Estilos, formatos y funciones auxiliares para las hojas del libro Excel."""

from __future__ import annotations

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import ND


# Colores y bordes usados en las hojas.
AZUL = "1F3864"
GRIS = "F2F2F2"
AMBAR = "FFF2CC"
BORDE = Side(style="thin", color="BFBFBF")

# Formatos numéricos de Excel.
FMT_MILES = '#,##0.00;[Red]-#,##0.00'
FMT_PORC = "0.00%"
FMT_ANIO = "0"
FMT_TEXTO = "@"
FMT_FECHA = "dd-mmm-yyyy"

# Colores para los semáforos, con contraste AAA de WCAG 2.2 (>= 7:1).
# Verde: 7.21:1 · Amarillo: 7.21:1 · Rojo: 7.08:1.
VERDE_FILL = PatternFill("solid", fgColor="D1E7DD")
AMARILLO_FILL = PatternFill("solid", fgColor="FFF3CD")
ROJO_FILL = PatternFill("solid", fgColor="F8D7DA")
VERDE_FONT = Font(color="0F5132", bold=True)
AMARILLO_FONT = Font(color="664D03", bold=True)
ROJO_FONT = Font(color="842029", bold=True)

# Filas comunes para los encabezados y los datos.
FILA_ENC = 4
FILA_DATOS = 5

# Encabezados de las hojas del libro.
ENC_CATALOGO = [
    "No.", "Clave CASFIM", "SOFIPO", "Federación", "Entidad Legal",
    "Sitio Web", "Fecha de Autorización", "Nivel", "Estatus",
]

ENC_DATOS = [
    "No.", "SOFIPO", "Año", "Mes", "Trimestre",
    "Activos Totales\n(miles de pesos)",
    "Cartera de crédito etapa 1 + etapa 2\n(miles de pesos)",
    "Cartera de crédito etapa 3\n(miles de pesos)",
    "Cartera Total\n(miles de pesos)",
    "Estimaciones preventivas para riesgos crediticios\n(miles de pesos)",
    "Cartera Neta\n(miles de pesos)",
    "Ingresos por intereses\n(YTD, miles de pesos)",
    "Gastos por intereses\n(YTD, miles de pesos)",
    "Margen Financiero\n(YTD, miles de pesos)",
    "Gastos de administración y promoción\n(YTD, miles de pesos)",
    "Resultado Neto\n(YTD, miles de pesos)",
    "Captación Tradicional\n(miles de pesos)",
    "Captación del mes anterior\n(miles de pesos)",
    "Cartera Total del mes anterior\n(miles de pesos)",
    "Capital Neto\n(miles de pesos)",
    "Requerimiento de capital\n(miles de pesos)",
    "NICAP\n(%)",
    "Resultado Neto 12M\n(miles de pesos)",
    "Activo promedio 12M\n(miles de pesos)",
    "Capital contable promedio 12M\n(miles de pesos)",
]

ENC_INDICADORES = [
    "No.", "SOFIPO", "Año", "Mes", "Trimestre",
    "NICAP\n(%)", "IMOR\n(%)", "ICOR\n(%)", "ROA\n(%)", "ROE\n(%)",
    "Margen Financiero\n(%)", "Crecimiento de Captación\n(%)",
    "Crecimiento de Cartera\n(%)",
    "Resultado Neto\n(miles de pesos)", "Crédito Neto\n(%)",
    "% Gastos de Administración\n(%)",
]

ENC_DICCIONARIO = ["Indicador", "Fórmula", "Grupo de razón financiera", "Descripción"]
ENC_SEMAFOROS = ["Rango", "Verde", "Amarillo", "Rojo"]


def _estilo_encabezado(ws, fila: int, columnas: int, alto: int = 46) -> None:
    """Aplica formato uniforme a una fila de encabezados."""
    for j in range(1, columnas + 1):
        celda = ws.cell(row=fila, column=j)
        celda.font = Font(bold=True, color="FFFFFF", size=10)
        celda.fill = PatternFill("solid", fgColor=AZUL)
        celda.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        celda.border = Border(
            left=BORDE,
            right=BORDE,
            top=BORDE,
            bottom=Side(style="medium", color=AZUL),
        )

    ws.row_dimensions[fila].height = alto


def _titulo(ws, columnas: int, titulo: str, subtitulo: str, nota: str) -> None:
    """Escribe y da formato al título, subtítulo y nota de una hoja."""
    ultima = get_column_letter(columnas)

    ws.merge_cells(f"A1:{ultima}1")
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, size=14, color=AZUL)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    ws.merge_cells(f"A2:{ultima}2")
    ws["A2"] = subtitulo
    ws["A2"].font = Font(bold=True, size=11)
    ws["A2"].alignment = Alignment(horizontal="center")

    ws.merge_cells(f"A3:{ultima}3")
    ws["A3"] = nota
    ws["A3"].font = Font(size=9, italic=True, color="595959")
    ws["A3"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[3].height = 28


def _notas(ws, columna_final: int, fila: int, textos: list[str]) -> None:
    """Escribe notas en filas sucesivas, combinando las celdas de cada fila."""
    for texto in textos:
        ws.merge_cells(
            start_row=fila,
            start_column=1,
            end_row=fila,
            end_column=columna_final,
        )
        celda = ws.cell(row=fila, column=1, value=texto)
        celda.font = Font(size=8, color="595959")
        celda.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True,
        )
        ws.row_dimensions[fila].height = 26
        fila += 1


def _texto(
    ws,
    fila: int,
    col: int,
    valor,
    negrita=False,
    alineacion="left",
    color=None,
) -> None:
    """Escribe texto en una celda y configura su alineación y fuente."""
    celda = ws.cell(row=fila, column=col, value=valor)
    celda.number_format = FMT_TEXTO

    if color:
        celda.font = Font(
            bold=negrita,
            italic=color is not None and color != "000000",
            color=color,
        )
    else:
        celda.font = Font(bold=negrita)

    celda.alignment = Alignment(
        horizontal=alineacion,
        vertical="center",
        wrap_text=False,
        indent=1 if alineacion == "left" else 0,
    )


def _celda_numero(ws, fila: int, col: int, valor, formato: str) -> None:
    """Escribe un número o muestra «n.d.» en gris cuando no hay dato."""
    celda = ws.cell(row=fila, column=col)

    if isinstance(valor, (int, float)):
        celda.value = float(valor)
        celda.number_format = formato
        celda.font = Font()
        celda.alignment = Alignment(horizontal="right", vertical="center")
    else:
        celda.value = ND
        celda.number_format = FMT_TEXTO
        celda.font = Font(italic=True, color="808080")
        celda.alignment = Alignment(horizontal="center", vertical="center")


def _celda_formula(ws, fila: int, col: int, formula: str, formato: str) -> None:
    """Escribe una fórmula con formato destacado y alineación a la derecha."""
    celda = ws.cell(row=fila, column=col, value=formula)
    celda.number_format = formato
    celda.font = Font(bold=True, color=AZUL)
    celda.alignment = Alignment(horizontal="right", vertical="center")
    celda.fill = PatternFill("solid", fgColor=AMBAR)


def _borde_y_banda(ws, fila: int, columnas: int, banda: bool) -> None:
    """Aplica bordes a una fila y, si corresponde, un fondo alternado."""
    for j in range(1, columnas + 1):
        celda = ws.cell(row=fila, column=j)
        celda.border = Border(
            left=BORDE,
            right=BORDE,
            top=BORDE,
            bottom=BORDE,
        )

        # No reemplaza los fondos que ya tenga la celda.
        if banda and celda.fill.fgColor.rgb in (None, "00000000"):
            celda.fill = PatternFill("solid", fgColor=GRIS)


def _texto_envuelto(ws, fila: int, col: int, valor, alto: int = 30) -> None:
    """Escribe texto con ajuste de línea y establece la altura de la fila."""
    celda = ws.cell(row=fila, column=col, value=valor)
    celda.number_format = FMT_TEXTO
    celda.font = Font(size=10)
    celda.alignment = Alignment(
        horizontal="left",
        vertical="top",
        wrap_text=True,
        indent=1,
    )
    ws.row_dimensions[fila].height = alto

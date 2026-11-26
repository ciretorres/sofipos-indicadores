"""Funciones para crear y dar formato a la hoja «Diccionario»."""

from __future__ import annotations

from openpyxl import Workbook

from ..estilos import (
    ENC_DICCIONARIO,
    FILA_DATOS,
    FILA_ENC,
    _borde_y_banda,
    _estilo_encabezado,
    _notas,
    _texto,
    _texto_envuelto,
    _titulo,
)
from ..fichas import FICHAS
from ..utilidades import enumerar


def escribir_diccionario(wb: Workbook, anios: list[int]) -> None:
    """Agrega al libro la hoja «Diccionario» con las fichas de indicadores.

    Incluye la definición, fórmula, grupo y descripción de cada indicador,
    además de notas aclaratorias y opciones de presentación para la hoja.
    Los años recibidos se usan para mencionar las hojas de indicadores
    correspondientes en el subtítulo.

    Args:
        wb: Libro de Excel al que se agrega la hoja.
        anios: Años cuyas hojas de indicadores se mencionan en el subtítulo.
    """
    ws = wb.create_sheet("Diccionario")
    ncols = len(ENC_DICCIONARIO)

    _titulo(
        ws,
        ncols,
        "SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Diccionario de indicadores",
        "Qué mide cada indicador, cómo se calcula y a qué grupo de razón financiera pertenece",
        "Cada renglón corresponde a una columna de las hojas "
        f"{enumerar([f'Indicadores {a}' for a in anios])}. "
        "Las fórmulas se expresan en función de las columnas de Datos Fuente.",
    )

    # Escribe y aplica formato a los encabezados.
    for j, encabezado in enumerate(ENC_DICCIONARIO, start=1):
        ws.cell(row=FILA_ENC, column=j, value=encabezado)
    _estilo_encabezado(ws, FILA_ENC, ncols, alto=32)

    # Escribe la ficha de cada indicador.
    for i, ficha in enumerate(FICHAS):
        r = FILA_DATOS + i
        _texto(ws, r, 1, ficha.indicador, negrita=True)
        _texto_envuelto(ws, r, 2, ficha.formula, alto=34)
        _texto(ws, r, 3, ficha.grupo, alineacion="center")
        _texto_envuelto(ws, r, 4, ficha.descripcion, alto=40)
        _borde_y_banda(ws, r, ncols, i % 2 == 1)

    ultima = FILA_DATOS + len(FICHAS) - 1

    # Agrega aclaraciones sobre los indicadores y sus umbrales.
    notas = [
        "Grupos de razón financiera: solvencia (capacidad de absorber pérdidas), calidad del activo "
        "(deterioro de la cartera), cobertura financiera (respaldo de la cartera vencida), "
        "rentabilidad (generación de resultado), crecimiento (evolución de captación y cartera), "
        "estructura financiera (composición del balance) y eficiencia operativa (costo por peso de ingreso).",
        "ROA y ROE se calculan con los insumos de 12 meses que publica el tablero "
        "(conceptos 27200043, 27200044 y 27200051) y reproducen los indicadores 27200001 y 27200002.",
        "Resultado Neto se documenta aquí, pero NO tiene semáforo: "
        "no se le asignó un umbral de referencia.",
        "Cuando el margen financiero es cero o negativo, % Gastos de Administración se muestra "
        "como n.a. (no aplicable) y el formato condicional lo marca en rojo.",
        "Los umbrales de los demás indicadores están en la hoja Semáforos y son los mismos que "
        "aplica el formato condicional de las hojas Indicadores.",
    ]
    _notas(ws, ncols, ultima + 2, notas)

    # Ajusta el ancho de las columnas.
    for col, ancho in {"A": 30, "B": 62, "C": 28, "D": 88}.items():
        ws.column_dimensions[col].width = ancho

    # Configura la navegación y la presentación para impresión.
    ws.freeze_panes = f"A{FILA_DATOS}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{FILA_ENC}:{FILA_ENC}"

"""Funciones para crear las hojas de indicadores y aplicar sus semáforos."""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

from ..config import ANIOS, MESES, ND
from ..estilos import (
    ENC_INDICADORES,
    FILA_DATOS,
    FILA_ENC,
    FMT_ANIO,
    FMT_MILES,
    FMT_PORC,
    _borde_y_banda,
    _celda_formula,
    _celda_numero,
    _estilo_encabezado,
    _notas,
    _texto,
    _titulo,
)
from ..fichas import COLORES, COND_NA, FICHAS_SEMAFORO


def aplicar_formato_condicional(ws, nfilas: int) -> None:
    """Aplica los colores de semáforo a los indicadores de la hoja.

    Las reglas y sus tramos se obtienen de ``FICHAS_SEMAFORO``. Las celdas
    no numéricas no reciben los colores de los tramos; la condición especial
    para valores no aplicables se maneja según su propia regla.

    Args:
        ws: Hoja de Excel donde se aplicará el formato condicional.
        nfilas: Número de filas de datos de la hoja.
    """
    primera = FILA_DATOS
    ultima = FILA_DATOS + nfilas - 1

    for ficha in FICHAS_SEMAFORO:
        rango = f"{ficha.columna}{primera}:{ficha.columna}{ultima}"
        celda = f"{ficha.columna}{primera}"

        for _, condicion, color in ficha.tramos:
            relleno, fuente = COLORES[color]

            if condicion == COND_NA:
                regla = condicion.format(c=celda)
            else:
                regla = "AND(ISNUMBER({c}),{cond})".format(
                    c=celda,
                    cond=condicion.format(c=celda),
                )

            ws.conditional_formatting.add(
                rango,
                FormulaRule(
                    formula=[regla],
                    fill=relleno,
                    font=fuente,
                    stopIfTrue=False,
                ),
            )


def escribir_indicadores(wb: Workbook, filas: list[dict], anio: int) -> None:
    """Crea la hoja de indicadores para el año indicado.

    Escribe fórmulas que vinculan los datos con la hoja «Datos Fuente» del
    mismo año, aplica el formato condicional y agrega notas explicativas.

    Args:
        wb: Libro de Excel al que se agregará la hoja.
        filas: Registros que determinan las filas de datos que se escribirán.
        anio: Año de los datos y nombre de las hojas relacionadas.
    """
    nombre_hoja = f"Indicadores {anio}"
    ws = wb.create_sheet(nombre_hoja)
    ncols = len(ENC_INDICADORES)
    ref = f"'Datos Fuente {anio}'"

    # Prepara el periodo y los títulos de la hoja.
    meses = ANIOS[anio]
    rango = f"{MESES[min(meses)]}–{MESES[max(meses)]} {anio}"
    _titulo(
        ws,
        ncols,
        f"SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Indicadores {anio}",
        f"{rango} · {len(filas)} filas",
        "Todas las celdas son fórmulas que apuntan a la hoja Datos Fuente del mismo año, en la "
        "misma fila. NICAP y Resultado Neto se toman como parámetro; ROA y ROE se calculan con "
        "los insumos de 12 meses.",
    )

    # Escribe y aplica formato a los encabezados.
    for j, encabezado in enumerate(ENC_INDICADORES, start=1):
        ws.cell(row=FILA_ENC, column=j, value=encabezado)
    _estilo_encabezado(ws, FILA_ENC, ncols, alto=50)

    # Escribe los datos de identificación y las fórmulas de los indicadores.
    for i, f in enumerate(filas):
        r = FILA_DATOS + i

        _celda_numero(ws, r, 1, f["no"], "0")
        _texto(ws, r, 2, f["sofipo"], negrita=True)
        _celda_numero(ws, r, 3, f["anio"], FMT_ANIO)
        _texto(ws, r, 4, f["mes"], alineacion="center")
        _texto(ws, r, 5, f["trimestre"], alineacion="center")

        # Fórmulas porcentuales que referencian la hoja Datos Fuente.
        formulas_pct = {
            6: f'=IFERROR({ref}!V{r},"{ND}")',  # NICAP
            7: f'=IFERROR({ref}!H{r}/{ref}!I{r},"{ND}")',  # IMOR
            8: f'=IFERROR(ABS({ref}!J{r})/{ref}!H{r},"{ND}")',  # ICOR
            9: f'=IFERROR({ref}!W{r}/{ref}!X{r},"{ND}")',  # ROA
            10: f'=IFERROR({ref}!W{r}/{ref}!Y{r},"{ND}")',  # ROE
            11: f'=IFERROR({ref}!N{r}/{ref}!L{r},"{ND}")',  # Margen
            12: f'=IFERROR({ref}!Q{r}/{ref}!R{r}-1,"{ND}")',  # Crec. captación
            13: f'=IFERROR({ref}!I{r}/{ref}!S{r}-1,"{ND}")',  # Crec. cartera
            15: f'=IFERROR({ref}!K{r}/{ref}!F{r},"{ND}")',  # Crédito neto
            16: (
                f'=IFERROR(IF({ref}!N{r}<=0,"n.a.",'
                f'{ref}!O{r}/{ref}!N{r}),"{ND}")'
            ),  # % Gastos
        }
        for col, formula in formulas_pct.items():
            _celda_formula(ws, r, col, formula, FMT_PORC)

        # Columna N: resultado neto como parámetro, expresado en miles de pesos.
        _celda_formula(ws, r, 14, f'=IFERROR({ref}!P{r},"{ND}")', FMT_MILES)

        _borde_y_banda(ws, r, ncols, i % 2 == 1)

    ultima = FILA_DATOS + len(filas) - 1

    # Aplica los semáforos a las filas de datos.
    aplicar_formato_condicional(ws, len(filas))

    # Agrega notas sobre los cálculos, las unidades y los datos no disponibles.
    notas = [
        "NICAP = 'Datos Fuente'!V (parámetro). IMOR = Cartera E3 / Cartera Total. "
        "ICOR = ABS(Estimaciones preventivas) / Cartera E3.",
        "ROA = Resultado Neto 12M / Activo promedio 12M. "
        "ROE = Resultado Neto 12M / Capital contable promedio 12M. Reproducen los indicadores "
        "27200001 y 27200002 que publica la CNBV.",
        "Margen Financiero = Margen Financiero (columna N de Datos Fuente) / Ingresos por intereses. "
        "El margen es Ingresos por intereses − Gastos por intereses.",
        "Crecimiento de Captación = (Captación actual / Captación del mes anterior) − 1. "
        "Crecimiento de Cartera = (Cartera Total actual / Cartera Total del mes anterior) − 1.",
        "Resultado Neto = parámetro de la columna P de Datos Fuente (YTD, miles de pesos). "
        "No es un porcentaje y no tiene semáforo.",
        "Crédito Neto = Cartera Neta / Activos Totales. "
        "% Gastos de Administración = Gastos de administración y promoción / Margen Financiero.",
        "Cuando el margen financiero es cero o negativo, % Gastos de Administración no aplica: "
        "la fórmula devuelve 'n.a.' y la celda se marca en rojo.",
        "Cuando la fuente no reporta el dato, la fórmula devuelve el texto 'n.d.' en lugar de "
        "un cero no sustentado. Los semáforos se aplican solo a celdas numéricas.",
        "Semáforo en las columnas F a P, salvo N (Resultado Neto), que no tiene semáforo "
        "definido. Los umbrales están en la hoja Semáforos.",
    ]
    _notas(ws, ncols, ultima + 2, notas)

    # Ajusta el ancho de las columnas.
    anchos = {"A": 6, "B": 22, "C": 8, "D": 11, "E": 10}
    for j in range(6, ncols + 1):
        anchos[get_column_letter(j)] = 20
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho

    # Configura la navegación, los filtros y la presentación para impresión.
    ws.freeze_panes = f"C{FILA_DATOS}"
    ws.auto_filter.ref = f"A{FILA_ENC}:{get_column_letter(ncols)}{ultima}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{FILA_ENC}:{FILA_ENC}"

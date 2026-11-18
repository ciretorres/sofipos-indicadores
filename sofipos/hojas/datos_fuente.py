"""Funciones para crear las hojas de datos fuente del libro."""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from ..config import ANIOS, MESES, ND
from ..estilos import (
    ENC_DATOS,
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


def escribir_datos_fuente(
    wb: Workbook,
    filas: list[dict],
    anio: int,
    n_entidades: int,
) -> None:
    """Crea y da formato a la hoja de datos fuente de un año.

    Escribe los datos mensuales de las sociedades, calcula las columnas
    derivadas mediante fórmulas y agrega notas sobre las fuentes y criterios.

    Args:
        wb: Libro de Excel al que se agregará la hoja.
        filas: Registros mensuales que se escribirán en la hoja.
        anio: Año de los datos.
        n_entidades: Número de sociedades incluidas por mes.
    """
    nombre_hoja = f"Datos Fuente {anio}"
    ws = wb.create_sheet(nombre_hoja)
    ncols = len(ENC_DATOS)

    # Prepara el periodo y los títulos de la hoja.
    meses = ANIOS[anio]
    rango = f"{MESES[min(meses)]}–{MESES[max(meses)]} {anio}"
    _titulo(
        ws,
        ncols,
        f"SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Datos Fuente {anio}",
        f"{rango} · {n_entidades} sociedades por mes · {len(filas)} filas",
        "Todas las cifras están en MILES DE PESOS. Fuente: tablero mensual de la CNBV "
        "(027_datos.csv), saldos sin consolidar (tipo_saldo 130). "
        "Los estados de resultados son acumulados del año (YTD).",
    )

    # Escribe y aplica formato a los encabezados.
    for j, encabezado in enumerate(ENC_DATOS, start=1):
        ws.cell(row=FILA_ENC, column=j, value=encabezado)
    _estilo_encabezado(ws, FILA_ENC, ncols, alto=62)

    # Escribe cada registro y calcula las columnas derivadas.
    for i, f in enumerate(filas):
        r = FILA_DATOS + i

        _celda_numero(ws, r, 1, f["no"], "0")
        _texto(ws, r, 2, f["sofipo"], negrita=True)
        _celda_numero(ws, r, 3, f["anio"], FMT_ANIO)
        _texto(ws, r, 4, f["mes"], alineacion="center")
        _texto(ws, r, 5, f["trimestre"], alineacion="center")

        # Columnas F a H: datos del tablero.
        _celda_numero(ws, r, 6, f["activos"], FMT_MILES)
        _celda_numero(ws, r, 7, f["cartera_e12"], FMT_MILES)
        _celda_numero(ws, r, 8, f["cartera_e3"], FMT_MILES)

        # Columna I: cartera total = cartera E1+E2 + cartera E3.
        _celda_formula(ws, r, 9, f'=IFERROR(G{r}+H{r},"{ND}")', FMT_MILES)

        # Columna J: estimaciones; columna K: cartera neta.
        _celda_numero(ws, r, 10, f["estimaciones"], FMT_MILES)
        _celda_formula(ws, r, 11, f'=IFERROR(I{r}-ABS(J{r}),"{ND}")', FMT_MILES)

        # Columnas L y M: datos del tablero; columna N: margen financiero.
        _celda_numero(ws, r, 12, f["ingresos_int"], FMT_MILES)
        _celda_numero(ws, r, 13, f["gastos_int"], FMT_MILES)
        _celda_formula(ws, r, 14, f'=IFERROR(L{r}-M{r},"{ND}")', FMT_MILES)

        # Columnas O a U: datos del tablero.
        _celda_numero(ws, r, 15, f["gastos_adm"], FMT_MILES)
        _celda_numero(ws, r, 16, f["resultado_neto"], FMT_MILES)
        _celda_numero(ws, r, 17, f["captacion"], FMT_MILES)
        _celda_numero(ws, r, 18, f["captacion_prev"], FMT_MILES)
        _celda_numero(ws, r, 19, f["cartera_total_prev"], FMT_MILES)
        _celda_numero(ws, r, 20, f["capital_neto"], FMT_MILES)
        _celda_numero(ws, r, 21, f["requerimiento"], FMT_MILES)

        # Columna V: NICAP = capital neto / requerimiento.
        _celda_formula(ws, r, 22, f'=IFERROR(T{r}/U{r},"{ND}")', FMT_PORC)

        # Columnas W a Y: datos para calcular ROA y ROE.
        _celda_numero(ws, r, 23, f["rn_12m"], FMT_MILES)
        _celda_numero(ws, r, 24, f["activo_prom_12m"], FMT_MILES)
        _celda_numero(ws, r, 25, f["capital_prom_12m"], FMT_MILES)

        _borde_y_banda(ws, r, ncols, i % 2 == 1)

    ultima = FILA_DATOS + len(filas) - 1

    # Define las notas comunes y las específicas de cada año.
    notas_comunes = [
        "UNIDADES: todas las columnas monetarias están en MILES DE PESOS. El tablero de la CNBV "
        "publica en pesos; aquí cada importe se dividió entre 1,000.",
        "FUENTE: tablero de información financiera de la CNBV, descarga mensual 027_datos.csv "
        "(serie 201601–202608), catálogo de conceptos 027 y catálogo de instituciones 027. "
        "Solo se usan saldos sin consolidar (tipo_saldo = 130).",
        "La fila con institución '27' es el TOTAL DEL SECTOR y se excluye: no es una SOFIPO. "
        "La serie completa del tablero (todas las entidades) coincide con esa fila al peso.",
        "El libro cubre las 35 sociedades que están 'en operación' en el Universo de junio de "
        "2026. CAME (027038), revocada el 19 de septiembre de 2025, no se incluye aunque el "
        "tablero conserve algunos meses suyos.",
        "Los estados de resultados (ingresos, gastos, gastos de administración y resultado neto) "
        "son ACUMULADOS DEL AÑO (YTD), igual que los publica el tablero. No son flujos del mes.",
        "PERIODO ANTERIOR: es el mes inmediato anterior. Enero de 2025 se compara contra diciembre "
        "de 2024 y enero de 2026 contra diciembre de 2025. Las columnas R y S son DATO descargado, "
        "no fórmula, para que cada hoja sea autocontenida.",
        "Cartera Total (I) = G + H. Cartera Neta (K) = I - ABS(J). Margen Financiero (N) = L - M. "
        "NICAP (V) = T / U. Las cuatro son fórmulas visibles en cada celda.",
        "Cartera etapa 3 (H) se toma del concepto 100800104003; cuando el tablero no publica el "
        "desglose, se deriva como Cartera de crédito total (130800103001) menos Cartera E1+E2, "
        "lo que en todos esos casos da exactamente cero.",
        "ROA y ROE de la hoja Indicadores se calculan con: Resultado Neto 12M (W) / Activo "
        "promedio 12M (X) y Resultado Neto 12M (W) / Capital contable promedio 12M (Y).",
        "n.d. = no disponible: la entidad no operaba o no reportó ese concepto en el mes. "
        "No se imputan ceros.",
    ]
    notas_2026 = notas_comunes + [
        "Agosto de 2026 es el corte más reciente del tablero y está INCOMPLETO: Nu México "
        "(027014) todavía no reporta ese mes y Capital Neto / Requerimiento no se publican para "
        "agosto, por lo que NICAP y los indicadores que dependen de ellos aparecen como n.d.",
        "La serie del tablero está reexpresada. Por eso las cifras de 2025 pueden diferir de los "
        "boletines BE_SOFIPOS_2025 que se usaron antes; se conserva la serie revisada de la CNBV.",
    ]
    notas_2025 = notas_comunes + [
        "En 2025 varias entidades aún no iniciaban operaciones (Impulso 027040, Trafalgar 027052 y "
        "Acción y Evolución 027028): sus celdas aparecen como n.d. hasta que el tablero las reporta.",
        "La serie del tablero está reexpresada. Por eso las cifras de 2025 pueden diferir de los "
        "boletines BE_SOFIPOS_2025 que se usaron antes; se conserva la serie revisada de la CNBV.",
    ]
    notas = notas_2026 if anio == 2026 else notas_2025
    _notas(ws, ncols, ultima + 2, notas)

    # Ajusta el ancho de las columnas.
    anchos = {"A": 6, "B": 22, "C": 8, "D": 11, "E": 10}
    for j in range(6, ncols + 1):
        anchos[get_column_letter(j)] = 21
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho

    # Configura la navegación, los filtros y la impresión de la hoja.
    ws.freeze_panes = f"C{FILA_DATOS}"
    ws.auto_filter.ref = f"A{FILA_ENC}:{get_column_letter(ncols)}{ultima}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{FILA_ENC}:{FILA_ENC}"

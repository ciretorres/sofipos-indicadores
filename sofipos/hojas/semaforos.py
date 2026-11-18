"""Funciones para crear la hoja de umbrales de los semáforos."""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill

from ..estilos import (
    AZUL,
    BORDE,
    FILA_DATOS,
    FILA_ENC,
    ENC_SEMAFOROS,
    AMARILLO_FILL,
    AMARILLO_FONT,
    ROJO_FILL,
    ROJO_FONT,
    VERDE_FILL,
    VERDE_FONT,
    _estilo_encabezado,
    _notas,
    _titulo,
)
from ..fichas import COLORES, FICHAS_SEMAFORO, TOTAL_REGLAS


def escribir_semaforos(wb: Workbook) -> None:
    """Crea y da formato a la hoja «Semáforos».

    Muestra los umbrales de cada indicador, las condiciones que se usan en
    Excel y una leyenda para interpretar los colores.

    Args:
        wb: Libro de Excel al que se agregará la hoja.
    """
    ws = wb.create_sheet("Semáforos")
    ncols = len(ENC_SEMAFOROS)

    _titulo(
        ws,
        ncols,
        "SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Semáforos",
        "Umbrales de referencia por indicador · se aplican como formato condicional en Indicadores",
        "Verde = dentro del rango objetivo · Amarillo = vigilancia · "
        "Rojo = fuera de rango o n.a. (margen ≤ 0). Las condiciones se evalúan sobre "
        "fracciones (141% = 1.41) y solo si la celda es numérica.",
    )

    # Escribe y aplica formato a los encabezados.
    for j, encabezado in enumerate(ENC_SEMAFOROS, start=1):
        ws.cell(row=FILA_ENC, column=j, value=encabezado)
    _estilo_encabezado(ws, FILA_ENC, ncols, alto=24)

    # Añade los rangos y las fórmulas de cada indicador.
    fila = FILA_DATOS
    for ficha in FICHAS_SEMAFORO:
        # Título del indicador y su grupo.
        ws.merge_cells(
            start_row=fila,
            start_column=1,
            end_row=fila,
            end_column=ncols,
        )
        titulo = ws.cell(
            row=fila,
            column=1,
            value=(
                f"{ficha.indicador}  ·  {ficha.grupo}  ·  "
                f"columna {ficha.columna} de Indicadores"
            ),
        )
        titulo.font = Font(bold=True, color="FFFFFF", size=10)
        titulo.fill = PatternFill("solid", fgColor="2F5597")
        titulo.alignment = Alignment(
            horizontal="left",
            vertical="center",
            indent=1,
        )
        ws.row_dimensions[fila].height = 20
        fila += 1

        # Muestra los umbrales agrupados por color.
        etiqueta = ws.cell(row=fila, column=1, value="Umbral")
        etiqueta.font = Font(bold=True, size=10)
        etiqueta.alignment = Alignment(
            horizontal="left",
            vertical="center",
            indent=1,
        )
        etiqueta.border = Border(
            left=BORDE,
            right=BORDE,
            top=BORDE,
            bottom=BORDE,
        )

        por_color: dict[str, list[str]] = {}
        for texto, _, color in ficha.tramos:
            por_color.setdefault(color, []).append(texto)

        for col, color in enumerate(("verde", "amarillo", "rojo"), start=2):
            celda = ws.cell(row=fila, column=col)
            if color in por_color:
                relleno, fuente = COLORES[color]
                celda.value = " · ".join(por_color[color])
                celda.fill = relleno
                celda.font = fuente
            else:
                celda.value = "—"
                celda.font = Font(color="A6A6A6")
                celda.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                )

            celda.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )
            celda.border = Border(
                left=BORDE,
                right=BORDE,
                top=BORDE,
                bottom=BORDE,
            )

        ws.row_dimensions[fila].height = 20
        fila += 1

        # Expone las condiciones que usa Excel para aplicar cada color.
        celda = ficha.columna + str(FILA_DATOS)
        detalles = []
        for texto, condicion, _ in ficha.tramos:
            detalles.append(f"{texto} → {condicion.format(c=celda)}")

        ws.merge_cells(
            start_row=fila,
            start_column=1,
            end_row=fila,
            end_column=ncols,
        )
        nota = ws.cell(
            row=fila,
            column=1,
            value="Excel (fracción): " + " · ".join(detalles),
        )
        nota.font = Font(size=8, italic=True, color="595959")
        nota.alignment = Alignment(
            horizontal="left",
            vertical="center",
            indent=1,
        )
        ws.row_dimensions[fila].height = 16
        fila += 2

    # Añade una leyenda para explicar el significado de cada color.
    leyenda_fila = fila
    ws.merge_cells(
        start_row=leyenda_fila,
        start_column=1,
        end_row=leyenda_fila,
        end_column=ncols,
    )
    leyenda = ws.cell(row=leyenda_fila, column=1, value="Leyenda")
    leyenda.font = Font(bold=True, size=10, color=AZUL)
    leyenda.alignment = Alignment(
        horizontal="left",
        vertical="center",
        indent=1,
    )
    fila += 1

    for texto, relleno, fuente in (
        ("Verde: comportamiento dentro del rango objetivo", VERDE_FILL, VERDE_FONT),
        ("Amarillo: vigilancia, el indicador se acerca al límite", AMARILLO_FILL, AMARILLO_FONT),
        ("Rojo: fuera de rango o n.a. (margen ≤ 0), requiere seguimiento", ROJO_FILL, ROJO_FONT),
    ):
        for col in range(1, ncols + 1):
            ws.cell(row=fila, column=col).fill = relleno
            ws.cell(row=fila, column=col).border = Border(
                left=BORDE,
                right=BORDE,
                top=BORDE,
                bottom=BORDE,
            )

        ws.merge_cells(
            start_row=fila,
            start_column=1,
            end_row=fila,
            end_column=ncols,
        )
        celda = ws.cell(row=fila, column=1, value=texto)
        celda.font = fuente
        celda.alignment = Alignment(
            horizontal="left",
            vertical="center",
            indent=1,
        )
        ws.row_dimensions[fila].height = 18
        fila += 1

    # Incluye aclaraciones sobre los rangos y la aplicación de los semáforos.
    notas = [
        "Estos son los mismos umbrales que aplica el formato condicional de las hojas Indicadores "
        f"2026 e Indicadores 2025 ({len(FICHAS_SEMAFORO)} indicadores y {TOTAL_REGLAS} reglas). "
        "Al cambiar un rango aquí hay que cambiarlo en el generador (lista FICHAS).",
        "Crecimiento de Captación y Crecimiento de Cartera: verde arriba de 10%, amarillo de "
        "0% a 10%, rojo por debajo de 0%.",
        "Resultado Neto NO tiene semáforo: se muestra como dato de contexto y no se pinta.",
        "NICAP: el amarillo va de 130% a menos de 141% y el rojo por debajo de 130%, para que "
        "ningún valor quede sin color entre los tramos.",
        "Crédito Neto: verde de 70% a 85%, amarillo de 60% a menos de 70% y de más de 85% a 90%, "
        "rojo por debajo de 60% o por arriba de 90%.",
        "% Gastos de Administración: verde por debajo de 60%, amarillo de 60% a 80%, rojo arriba "
        "de 80%. Si el margen financiero es cero o negativo, se muestra n.a. y se marca en rojo.",
        "IMOR, ICOR, ROA, ROE, Margen Financiero y % Gastos de Administración toman el límite "
        "superior del tramo como amarillo (10%, 100%, 2%, 15%, 70% y 80% respectivamente).",
    ]
    _notas(ws, ncols, fila + 1, notas)

    # Ajusta el ancho de las columnas y la presentación para impresión.
    for col, ancho in {"A": 40, "B": 24, "C": 24, "D": 24}.items():
        ws.column_dimensions[col].width = ancho

    ws.freeze_panes = f"A{FILA_DATOS}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

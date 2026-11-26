"""Validaciones de estructura, contenido y formato del libro Excel generado."""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .config import ND
from .estilos import FILA_DATOS, FILA_ENC, FMT_FECHA
from .fichas import FICHAS, FICHAS_SEMAFORO, TOTAL_REGLAS


def hojas_esperadas(filas: dict[int, list[dict]]) -> list[str]:
    """Devuelve los nombres de las hojas esperadas y su orden.

    El orden se deriva de los años presentes en ``filas``, de mayor a menor.
    """
    anios = sorted(filas, reverse=True)
    return (
        ["Catálogo SOFIPOS"]
        + [f"Datos Fuente {a}" for a in anios]
        + [f"Indicadores {a}" for a in anios]
        + ["Diccionario", "Semáforos"]
    )


def validar_libro(
    ruta: Path,
    universo: list[dict],
    filas: dict[int, list[dict]],
) -> bool:
    """Valida la estructura, el contenido y el formato del libro Excel.

    Comprueba las hojas esperadas, las filas del catálogo y de cada año, la
    alineación de datos, algunas fórmulas y encabezados, el diccionario, los
    bloques de semáforos y las reglas de formato condicional. Imprime un
    resumen de las comprobaciones.

    Args:
        ruta: Ruta del archivo Excel que se validará.
        universo: Sociedades que deben aparecer en el catálogo. Las filas
            mensuales se cuentan a partir de ``filas`` e incluyen las bajas.
        filas: Registros preparados, agrupados por año.

    Returns:
        ``True`` si pasan todas las comprobaciones; de lo contrario, ``False``.
    """
    ok = True
    wb = load_workbook(ruta)

    print("\n=== Validación del libro generado ===")
    print(f"  hojas: {wb.sheetnames}")

    # Comprueba que las hojas estén presentes y en el orden esperado.
    esperadas = hojas_esperadas(filas)
    if wb.sheetnames != esperadas:
        print(f"  ERROR: se esperaban {esperadas}.")
        ok = False

    cat = wb["Catálogo SOFIPOS"]
    dic = wb["Diccionario"]
    sem = wb["Semáforos"]

    # Comprueba el número de sociedades y el formato de las fechas del catálogo.
    filas_cat = sum(
        1
        for r in range(FILA_DATOS, FILA_DATOS + len(universo))
        if cat.cell(row=r, column=2).value
    )
    print(f"  Catálogo SOFIPOS: {filas_cat} filas (esperado {len(universo)})")
    if filas_cat != len(universo):
        ok = False

    formatos_fecha = {
        cat.cell(row=r, column=7).number_format
        for r in range(FILA_DATOS, FILA_DATOS + len(universo))
    }
    print(
        f"  Catálogo SOFIPOS: formato de fecha {formatos_fecha} "
        f"(esperado {{'{FMT_FECHA}'}})"
    )
    if formatos_fecha != {FMT_FECHA}:
        ok = False

    for anio, lista in filas.items():
        datos = wb[f"Datos Fuente {anio}"]
        ind = wb[f"Indicadores {anio}"]

        # Comprueba que ambas hojas tengan el número de filas esperado.
        filas_datos = sum(
            1
            for r in range(FILA_DATOS, FILA_DATOS + len(lista))
            if datos.cell(row=r, column=1).value is not None
        )
        filas_ind = sum(
            1
            for r in range(FILA_DATOS, FILA_DATOS + len(lista))
            if ind.cell(row=r, column=1).value is not None
        )
        print(
            f"  Datos Fuente {anio}: {filas_datos} filas · "
            f"Indicadores {anio}: {filas_ind} "
            f"(esperado {len(lista)})"
        )
        if filas_datos != len(lista) or filas_ind != len(lista):
            ok = False

        # Verifica que los registros correspondientes estén alineados.
        desalineadas = [
            r
            for r in range(FILA_DATOS, FILA_DATOS + len(lista))
            if (
                datos.cell(row=r, column=2).value
                != ind.cell(row=r, column=2).value
                or datos.cell(row=r, column=4).value
                != ind.cell(row=r, column=4).value
                or datos.cell(row=r, column=5).value
                != ind.cell(row=r, column=5).value
            )
        ]
        print(
            f"    filas desalineadas Datos Fuente vs Indicadores: "
            f"{len(desalineadas)}"
        )
        if desalineadas:
            ok = False

        # Comprueba que haya una fila por sociedad en cada mes. El conteo se
        # deriva de las claves distintas de ``filas`` (incluye las bajas).
        meses = {}
        for r in range(FILA_DATOS, FILA_DATOS + len(lista)):
            mes = datos.cell(row=r, column=4).value
            meses[mes] = meses.get(mes, 0) + 1

        esperado = len({f["clave"] for f in lista})
        malos = {
            mes: cantidad
            for mes, cantidad in meses.items()
            if cantidad != esperado
        }
        print(
            f"    meses: {len(meses)} · "
            f"{'todas con ' + str(esperado) if not malos else malos}"
        )
        if malos:
            ok = False

        # Comprueba algunas fórmulas clave de la hoja de datos fuente.
        r0 = FILA_DATOS
        for col, esperado in [
            (9, f'=IFERROR(G{r0}+H{r0},"{ND}")'),
            (11, f'=IFERROR(I{r0}-ABS(J{r0}),"{ND}")'),
            (14, f'=IFERROR(L{r0}-M{r0},"{ND}")'),
            (22, f'=IFERROR(T{r0}/U{r0},"{ND}")'),
        ]:
            valor = datos.cell(row=r0, column=col).value
            if valor != esperado:
                print(
                    f"      ERROR: Datos Fuente {anio} "
                    f"{get_column_letter(col)}{r0} = {valor!r}"
                )
                ok = False

        # Verifica que las columnas de indicadores contengan fórmulas.
        for col in range(6, 17):
            valor = ind.cell(row=r0, column=col).value
            if not (isinstance(valor, str) and valor.startswith("=")):
                print(
                    f"      ERROR: Indicadores {anio} "
                    f"{get_column_letter(col)}{r0} no es fórmula: {valor!r}"
                )
                ok = False

        # Comprueba que la fórmula de gastos contemple el caso "n.a.".
        formula_gastos = ind.cell(row=r0, column=16).value
        if '"n.a."' not in str(formula_gastos):
            print(
                f"      ERROR: Indicadores {anio} P{r0} no maneja n.a.: "
                f"{formula_gastos!r}"
            )
            ok = False

        # Verifica el encabezado esperado para la columna de resultado neto.
        if (
            ind.cell(row=FILA_ENC, column=14).value
            != "Resultado Neto\n(miles de pesos)"
        ):
            print(
                f"      ERROR: Indicadores {anio} "
                "no tiene el orden de columnas esperado."
            )
            ok = False

    # Comprueba la cantidad de fichas del diccionario.
    filas_dic = sum(
        1
        for r in range(FILA_DATOS, FILA_DATOS + len(FICHAS))
        if dic.cell(row=r, column=1).value
    )
    print(f"  Diccionario: {filas_dic} indicadores (esperado {len(FICHAS)})")
    if filas_dic != len(FICHAS):
        ok = False

    # Cuenta los bloques de indicadores de la hoja de semáforos.
    bloques = sum(
        1
        for r in range(FILA_DATOS, sem.max_row + 1)
        if sem.cell(row=r, column=1).value == "Umbral"
    )
    print(
        f"  Semáforos: {bloques} bloques de indicador "
        f"(esperado {len(FICHAS_SEMAFORO)})"
    )
    if bloques != len(FICHAS_SEMAFORO):
        ok = False

    # Compara la cantidad de reglas de formato condicional por año.
    for anio in filas:
        ind = wb[f"Indicadores {anio}"]
        reglas = sum(
            len(lista_reglas)
            for lista_reglas in ind.conditional_formatting._cf_rules.values()
        )
        print(
            f"  formato condicional {anio}: {reglas} reglas "
            f"(esperado {TOTAL_REGLAS})"
        )
        if reglas != TOTAL_REGLAS:
            ok = False

    return ok

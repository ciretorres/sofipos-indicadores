"""Pruebas de coherencia entre «Semáforos», el formato condicional y FICHAS.

Las pruebas usan un Workbook en memoria y no generan el libro completo.
"""

import re
from types import SimpleNamespace

import pytest
from openpyxl import Workbook

from sofipos.config import ANIOS
from sofipos.estilos import FILA_DATOS
from sofipos.fichas import COND_NA, FICHAS, FICHAS_SEMAFORO, TOTAL_REGLAS
from sofipos.hojas.indicadores import aplicar_formato_condicional
from sofipos.hojas.semaforos import escribir_semaforos

# Patrones para analizar referencias, porcentajes y condiciones de Excel.
REF = re.compile(r"[A-Z]+\d+")
PORC = re.compile(r"(\d+(?:\.\d+)?)%")
NUM = re.compile(r"-?\d+(?:\.\d+)?")
ENVOLTURA = re.compile(r"^AND\(ISNUMBER\([A-Z]+\d+\),(.*)\)$")

PREFIJO_NOTA = "Excel (fracción): "


@pytest.fixture(scope="module")
def libro():
    """Prepara la hoja «Semáforos» y las reglas de formato en memoria."""
    wb = Workbook()
    escribir_semaforos(wb, sorted(ANIOS, reverse=True))
    ws = wb["Semáforos"]

    umbrales: list[list[str]] = []
    conds: list[str] = []

    for fila in ws.iter_rows(min_col=1, max_col=4):
        valor = fila[0].value
        if valor == "Umbral":
            umbrales.append([celda.value for celda in fila[1:]])
        elif isinstance(valor, str) and valor.startswith(PREFIJO_NOTA):
            conds.append(valor[len(PREFIJO_NOTA):])

    ws_ind = wb.create_sheet("Indicadores")
    aplicar_formato_condicional(ws_ind, nfilas=10)
    formulas = [
        regla.formula[0]
        for cf in ws_ind.conditional_formatting
        for regla in cf.rules
    ]

    return SimpleNamespace(ws=ws, umbrales=umbrales, conds=conds, formulas=formulas)


def _sin_refs(texto: str) -> str:
    """Elimina referencias de celda antes de extraer números."""
    return REF.sub("", texto)


def _porcentajes(texto: str) -> set[float]:
    """Convierte los porcentajes del texto a fracciones; por ejemplo, 141% a 1.41."""
    return {float(x) / 100 for x in PORC.findall(_sin_refs(texto))}


def _numeros(texto: str) -> set[float]:
    """Extrae los valores numéricos de una condición, sin referencias de celda."""
    return {float(x) for x in NUM.findall(_sin_refs(texto))}


def _desenvuelve(formula: str) -> str:
    """Retira el envoltorio ``AND(ISNUMBER(<referencia>), ...)`` de una fórmula."""
    return ENVOLTURA.sub(r"\1", formula)


def _etiquetas(umbral: list[str]) -> list[str]:
    """Obtiene las etiquetas de las celdas de umbral, separadas por color."""
    return [
        etiqueta
        for celda in umbral
        if celda and celda != "—"
        for etiqueta in celda.split(" · ")
    ]


def _pares(nota: str) -> list[tuple[str, str]]:
    """Separa una nota en pares de etiqueta y condición."""
    return [tuple(par.split(" → ", 1)) for par in nota.split(" · ")]


def _nota_prosa(ws, prefijo: str) -> str:
    """Busca y devuelve en la columna A la nota que comienza con el prefijo."""
    for fila in ws.iter_rows(min_col=1, max_col=1):
        valor = fila[0].value
        if isinstance(valor, str) and valor.startswith(prefijo):
            return valor
    raise AssertionError(f"No hay nota que empiece con {prefijo!r}")


def test_la_hoja_tiene_un_bloque_por_ficha(libro):
    """Cada indicador con semáforo tiene umbrales y una nota de condiciones."""
    assert len(libro.umbrales) == len(FICHAS_SEMAFORO)
    assert len(libro.conds) == len(FICHAS_SEMAFORO)


def test_etiquetas_conciden_con_condiciones_y_con_formato_condicional(libro):
    """Comprueba que coincidan las etiquetas, las condiciones y las reglas."""
    inicio = 0
    for ficha, umbral, nota in zip(FICHAS_SEMAFORO, libro.umbrales, libro.conds):
        fin = inicio + len(ficha.tramos)
        bloque = libro.formulas[inicio:fin]
        inicio = fin

        etiquetas = _etiquetas(umbral)
        assert set(etiquetas) == {texto for texto, _, _ in ficha.tramos}

        numeros_nota: set[float] = set()
        for _, condicion in _pares(nota):
            numeros_nota |= _numeros(condicion)

        numeros_cf: set[float] = set()
        for formula in bloque:
            numeros_cf |= _numeros(formula)

        assert _porcentajes(" · ".join(etiquetas)) == numeros_nota
        assert numeros_nota == numeros_cf


def test_condiciones_de_la_nota_son_verbatim_a_las_formulas_cf(libro):
    """Las condiciones de la nota coinciden con las fórmulas de formato."""
    inicio = 0
    for ficha, nota in zip(FICHAS_SEMAFORO, libro.conds):
        fin = inicio + len(ficha.tramos)
        condiciones_nota = [condicion for _, condicion in _pares(nota)]
        condiciones_cf = [
            _desenvuelve(formula) for formula in libro.formulas[inicio:fin]
        ]
        inicio = fin
        assert condiciones_nota == condiciones_cf


def test_numero_de_reglas_cf_es_total_reglas(libro):
    """El formato condicional contiene exactamente ``TOTAL_REGLAS`` reglas."""
    assert TOTAL_REGLAS == 31
    assert len(libro.formulas) == TOTAL_REGLAS


def test_tramo_na_no_tiene_numeros(libro):
    """El tramo «n.a.» verifica texto y no incluye límites numéricos."""
    idx = [f.indicador for f in FICHAS_SEMAFORO].index(
        "% Gastos de Administración"
    )
    ficha = FICHAS_SEMAFORO[idx]
    etiqueta, condicion = _pares(libro.conds[idx])[-1]

    assert "n.a." in etiqueta
    assert condicion == COND_NA.format(c=ficha.columna + str(FILA_DATOS))
    assert _numeros(condicion) == set()


def test_notas_en_prosa_coinciden_con_los_tramos(libro):
    """Las notas de NICAP, Crédito Neto y Gastos citan sus límites."""
    for prefijo, indicador in (
        ("NICAP: ", "NICAP"),
        ("Crédito Neto: ", "Crédito Neto"),
        ("% Gastos de Administración: ", "% Gastos de Administración"),
    ):
        ficha = next(f for f in FICHAS if f.indicador == indicador)
        numeros: set[float] = set()
        for _, condicion, _ in ficha.tramos:
            numeros |= _numeros(condicion)

        assert _porcentajes(_nota_prosa(libro.ws, prefijo)) == numeros


def test_limite_superior_amarillo_de_seis_indicadores(libro):
    """La nota enumera en orden los límites superiores de seis tramos amarillos."""
    nota = _nota_prosa(libro.ws, "IMOR, ICOR, ROA, ROE")
    esperado = []

    for indicador in (
        "IMOR",
        "ICOR",
        "ROA",
        "ROE",
        "Margen Financiero",
        "% Gastos de Administración",
    ):
        ficha = next(f for f in FICHAS if f.indicador == indicador)
        amarillo = next(t for t in ficha.tramos if t[2] == "amarillo")
        ultimo = PORC.findall(amarillo[0])[-1]
        esperado.append(float(ultimo) / 100)

    citados = [float(x) / 100 for x in PORC.findall(nota)]
    assert citados == esperado


def test_resultado_neto_no_tiene_semaforo(libro):
    """Comprueba que Resultado Neto no tenga tramos y que la hoja lo indique."""
    ficha = next(f for f in FICHAS if f.indicador == "Resultado Neto")

    assert ficha.tramos == ()
    assert ficha not in FICHAS_SEMAFORO
    assert "Resultado Neto NO tiene semáforo" in _nota_prosa(
        libro.ws, "Resultado Neto NO"
    )

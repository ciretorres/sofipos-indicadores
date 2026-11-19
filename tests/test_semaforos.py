"""Pruebas de coherencia entre la hoja «Semáforos», el formato condicional y FICHAS.

Todo se construye en un ``Workbook`` en memoria: no se genera el libro completo.
"""

import re
from types import SimpleNamespace

import pytest
from openpyxl import Workbook

from sofipos.estilos import FILA_DATOS
from sofipos.fichas import COND_NA, FICHAS, FICHAS_SEMAFORO, TOTAL_REGLAS
from sofipos.hojas.indicadores import aplicar_formato_condicional
from sofipos.hojas.semaforos import escribir_semaforos

# Referencias de celda (F5): hay que quitarlas antes de extraer números.
REF = re.compile(r"[A-Z]+\d+")
PORC = re.compile(r"(\d+(?:\.\d+)?)%")
NUM = re.compile(r"-?\d+(?:\.\d+)?")
ENVOLTURA = re.compile(r"^AND\(ISNUMBER\([A-Z]+\d+\),(.*)\)$")

PREFIJO_NOTA = "Excel (fracción): "


@pytest.fixture(scope="module")
def libro():
    """Hoja «Semáforos» y reglas de formato condicional en memoria."""
    wb = Workbook()
    escribir_semaforos(wb)
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
    """Elimina las referencias de celda para no envenenar los números."""
    return REF.sub("", texto)


def _porcentajes(texto: str) -> set[float]:
    """Porcentajes citados en un texto, como fracciones (141% → 1.41)."""
    return {float(x) / 100 for x in PORC.findall(_sin_refs(texto))}


def _numeros(texto: str) -> set[float]:
    """Números de una condición de Excel, sin las referencias de celda."""
    return {float(x) for x in NUM.findall(_sin_refs(texto))}


def _desenvuelve(formula: str) -> str:
    """Quita el envoltorio ``AND(ISNUMBER(<ref>), …)`` de una fórmula CF."""
    return ENVOLTURA.sub(r"\1", formula)


def _etiquetas(umbral: list[str]) -> list[str]:
    """Etiquetas de una fila «Umbral» (celdas verde/amarillo/rojo)."""
    return [
        etiqueta
        for celda in umbral
        if celda and celda != "—"
        for etiqueta in celda.split(" · ")
    ]


def _pares(nota: str) -> list[tuple[str, str]]:
    """Pares (etiqueta, condición) de una nota «Excel (fracción)»."""
    return [par.split(" → ", 1) for par in nota.split(" · ")]


def _nota_prosa(ws, prefijo: str) -> str:
    """Devuelve la nota en prosa de la columna A que empieza por el prefijo."""
    for fila in ws.iter_rows(min_col=1, max_col=1):
        valor = fila[0].value
        if isinstance(valor, str) and valor.startswith(prefijo):
            return valor
    raise AssertionError(f"No hay nota que empiece con {prefijo!r}")


def test_la_hoja_tiene_un_bloque_por_ficha(libro):
    """Hay una fila de umbrales y una nota de condiciones por indicador."""
    assert len(libro.umbrales) == len(FICHAS_SEMAFORO)
    assert len(libro.conds) == len(FICHAS_SEMAFORO)


def test_etiquetas_conciden_con_condiciones_y_con_formato_condicional(libro):
    """Por indicador: etiquetas ÷ 100 == números de la condición == fórmulas CF."""
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
    """La nota «Excel (fracción)» es idéntica a la fórmula CF desenvuelta."""
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
    """El formato condicional define exactamente ``TOTAL_REGLAS`` reglas."""
    assert TOTAL_REGLAS == 31
    assert len(libro.formulas) == TOTAL_REGLAS


def test_tramo_na_no_tiene_numeros(libro):
    """El tramo n.a. solo exige el texto «n.a.»; su etiqueta lo menciona."""
    idx = [f.indicador for f in FICHAS_SEMAFORO].index(
        "% Gastos de Administración"
    )
    ficha = FICHAS_SEMAFORO[idx]
    etiqueta, condicion = _pares(libro.conds[idx])[-1]
    assert "n.a." in etiqueta
    assert condicion == COND_NA.format(c=ficha.columna + str(FILA_DATOS))
    assert _numeros(condicion) == set()


def test_notas_en_prosa_coinciden_con_los_tramos(libro):
    """NICAP, Crédito Neto y % Gastos citan los números de sus fichas."""
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
    """La nota cita en orden los límites superiores de seis tramos amarillos."""
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
    """La ficha «Resultado Neto» no define tramos y la hoja lo advierte."""
    ficha = next(f for f in FICHAS if f.indicador == "Resultado Neto")
    assert ficha.tramos == ()
    assert ficha not in FICHAS_SEMAFORO
    assert "Resultado Neto NO tiene semáforo" in _nota_prosa(
        libro.ws, "Resultado Neto NO"
    )

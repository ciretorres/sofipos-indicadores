"""Definiciones de indicadores y rangos para los semáforos del libro."""

from __future__ import annotations

from dataclasses import dataclass

from .estilos import (
    AMARILLO_FILL,
    AMARILLO_FONT,
    ROJO_FILL,
    ROJO_FONT,
    VERDE_FILL,
    VERDE_FONT,
)


# Estilos disponibles para cada nivel del semáforo.
COLORES = {
    "verde": (VERDE_FILL, VERDE_FONT),
    "amarillo": (AMARILLO_FILL, AMARILLO_FONT),
    "rojo": (ROJO_FILL, ROJO_FONT),
}


@dataclass(frozen=True)
class Ficha:
    """Contiene la definición de un indicador y sus rangos de evaluación.

    Attributes:
        indicador: Nombre del indicador.
        columna: Columna en la que se muestra en la hoja de indicadores.
        formula: Fórmula o referencia usada para obtener el valor.
        grupo: Categoría financiera a la que pertenece.
        descripcion: Explicación del indicador.
        tramos: Rangos del semáforo como tuplas de
            (etiqueta, condición de Excel, color). Puede estar vacío
            cuando el indicador no tiene semáforo.
    """

    indicador: str
    columna: str
    formula: str
    grupo: str
    descripcion: str
    tramos: tuple[tuple[str, str, str], ...] = ()


# Las condiciones se evalúan como fracciones porque Excel muestra las celdas
# con formato de porcentaje. La condición ISNUMBER se usa para evitar que los
# valores de texto, como «n.d.», se clasifiquen como si fueran números.
# COND_NA es una excepción: comprueba específicamente el texto «n.a.».
COND_NA = '{c}="n.a."'


# Definiciones de los indicadores incluidos en la hoja.
FICHAS: list[Ficha] = [
    Ficha(
        indicador="NICAP",
        columna="F",
        formula="Capital Neto / Requerimiento de capital",
        grupo="Solvencia",
        descripcion=(
            "Mide si el capital neto cubre los riesgos asumidos. "
            "El mínimo regulatorio es 131%; por debajo de eso la sociedad incumple."
        ),
        tramos=(
            ("≥ 141%", "{c}>=1.41", "verde"),
            ("130% – < 141%", "AND({c}>=1.3,{c}<1.41)", "amarillo"),
            ("< 130%", "{c}<1.3", "rojo"),
        ),
    ),
    Ficha(
        indicador="IMOR",
        columna="G",
        formula="Cartera E3 / Cartera Total",
        grupo="Calidad del activo",
        descripcion=(
            "Proporción de la cartera que ya no genera pago. Mide el deterioro "
            "de la cartera: a mayor IMOR, peor calidad."
        ),
        tramos=(
            ("< 5%", "{c}<0.05", "verde"),
            ("5% – 10%", "AND({c}>=0.05,{c}<=0.1)", "amarillo"),
            ("> 10%", "{c}>0.1", "rojo"),
        ),
    ),
    Ficha(
        indicador="ICOR",
        columna="H",
        formula="ABS(Estimaciones preventivas) / Cartera E3",
        grupo="Cobertura financiera",
        descripcion=(
            "Mide cuánto cubren las estimaciones preventivas a la cartera en "
            "etapa 3. Por encima de 100%, la cartera vencida está cubierta por completo."
        ),
        tramos=(
            ("> 100%", "{c}>1", "verde"),
            ("80% – 100%", "AND({c}>=0.8,{c}<=1)", "amarillo"),
            ("< 80%", "{c}<0.8", "rojo"),
        ),
    ),
    Ficha(
        indicador="ROA",
        columna="I",
        formula="Resultado Neto 12M / Activo promedio 12M",
        grupo="Rentabilidad",
        descripcion=(
            "Rendimiento anualizado del activo. Se calcula con el resultado neto "
            "de los últimos 12 meses y el activo promedio de esos 12 meses; reproduce "
            "el indicador ROA (27200001) que publica la CNBV."
        ),
        tramos=(
            ("> 2%", "{c}>0.02", "verde"),
            ("0% – 2%", "AND({c}>=0,{c}<=0.02)", "amarillo"),
            ("< 0%", "{c}<0", "rojo"),
        ),
    ),
    Ficha(
        indicador="ROE",
        columna="J",
        formula="Resultado Neto 12M / Capital contable promedio 12M",
        grupo="Rentabilidad",
        descripcion=(
            "Rendimiento anualizado sobre el patrimonio. Se calcula con el resultado "
            "neto y el capital contable promedio de los últimos 12 meses; reproduce "
            "el indicador ROE (27200002) que publica la CNBV."
        ),
        tramos=(
            ("> 15%", "{c}>0.15", "verde"),
            ("5% – 15%", "AND({c}>=0.05,{c}<=0.15)", "amarillo"),
            ("< 5%", "{c}<0.05", "rojo"),
        ),
    ),
    Ficha(
        indicador="Margen Financiero",
        columna="K",
        formula="Margen Financiero / Ingresos por intereses",
        grupo="Rentabilidad",
        descripcion=(
            "Parte del ingreso por intereses que queda después del costo de fondeo. "
            "Muestra la rentabilidad del negocio crediticio. El margen financiero "
            "se toma de la columna N de Datos Fuente (Ingresos por intereses − "
            "Gastos por intereses)."
        ),
        tramos=(
            ("> 70%", "{c}>0.7", "verde"),
            ("50% – 70%", "AND({c}>=0.5,{c}<=0.7)", "amarillo"),
            ("< 50%", "{c}<0.5", "rojo"),
        ),
    ),
    Ficha(
        indicador="Crecimiento de Captación",
        columna="L",
        formula="(Captación actual / Captación del mes anterior) − 1",
        grupo="Crecimiento",
        descripcion=(
            "Variación mensual de la captación tradicional (depósitos del público). "
            "Mide si la SOFIPO está ganando o perdiendo clientes y depósitos respecto "
            "del mes inmediato anterior."
        ),
        tramos=(
            ("> 10%", "{c}>0.1", "verde"),
            ("0% – 10%", "AND({c}>=0,{c}<=0.1)", "amarillo"),
            ("< 0%", "{c}<0", "rojo"),
        ),
    ),
    Ficha(
        indicador="Crecimiento de Cartera",
        columna="M",
        formula="(Cartera Total actual / Cartera Total del mes anterior) − 1",
        grupo="Crecimiento",
        descripcion=(
            "Variación mensual de la cartera de crédito total (etapas 1, 2 y 3). "
            "Mide la expansión o contracción del negocio crediticio; una caída "
            "señala contracción de la cartera."
        ),
        tramos=(
            ("> 10%", "{c}>0.1", "verde"),
            ("0% – 10%", "AND({c}>=0,{c}<=0.1)", "amarillo"),
            ("< 0%", "{c}<0", "rojo"),
        ),
    ),
    Ficha(
        indicador="Resultado Neto",
        columna="N",
        formula="Parámetro: columna P de Datos Fuente (resultado neto YTD)",
        grupo="Rentabilidad",
        descripcion=(
            "Resultado neto acumulado del año (YTD) en miles de pesos. No es un "
            "cociente ni tiene semáforo: se traslada como parámetro desde Datos "
            "Fuente para leerlo junto con los indicadores."
        ),
        tramos=(),
    ),
    Ficha(
        indicador="Crédito Neto",
        columna="O",
        formula="Cartera Neta / Activos Totales",
        grupo="Estructura financiera",
        descripcion=(
            "Peso de la cartera neta dentro del activo total. Indica qué tan "
            "concentrado está el balance en el negocio de crédito."
        ),
        tramos=(
            ("70% – 85%", "AND({c}>=0.7,{c}<=0.85)", "verde"),
            (
                "60% – < 70% y > 85% – 90%",
                "OR(AND({c}>=0.6,{c}<0.7),AND({c}>0.85,{c}<=0.9))",
                "amarillo",
            ),
            ("< 60% o > 90%", "OR({c}<0.6,{c}>0.9)", "rojo"),
        ),
    ),
    Ficha(
        indicador="% Gastos de Administración",
        columna="P",
        formula="Gastos de administración y promoción / Margen Financiero",
        grupo="Eficiencia operativa",
        descripcion=(
            "Costo operativo por cada peso de margen financiero. A menor "
            "porcentaje, más eficiente es la estructura de gastos. El margen "
            "financiero se toma de la columna N de Datos Fuente. Cuando el margen "
            "es cero o negativo, la razón no aplica y se muestra n.a."
        ),
        tramos=(
            ("< 60%", "{c}<0.6", "verde"),
            ("60% – 80%", "AND({c}>=0.6,{c}<=0.8)", "amarillo"),
            ("> 80%", "{c}>0.8", "rojo"),
            ("Margen ≤ 0 (n.a.)", COND_NA, "rojo"),
        ),
    ),
]

# Indicadores que tienen rangos definidos para el semáforo.
FICHAS_SEMAFORO = [f for f in FICHAS if f.tramos]

# Cantidad total de reglas de color definidas en esos rangos.
TOTAL_REGLAS = sum(len(f.tramos) for f in FICHAS_SEMAFORO)

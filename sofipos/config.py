"""Configuración de rutas, periodos y campos usados por el proyecto."""

from __future__ import annotations

from pathlib import Path


# Rutas principales del repositorio y archivos de entrada y salida.
AQUI = Path(__file__).resolve().parent.parent  # Este archivo está en sofipos/.
DIR_ENTRADAS = AQUI / "entradas"
DIR_SALIDAS = AQUI / "salidas"
SALIDA_DEFECTO = DIR_SALIDAS / "SOFIPOS_2026_Datos.xlsx"

# Archivos que se leen durante la generación y validación del libro.
BE_UNIVERSO = "BE_SOFIPOS_202606.xlsx"
CSV_TABLERO = Path("tablero") / "sh_datos_csv_27_i" / "027_datos.csv"
BE_VALIDACION = ["BE_SOFIPOS_202603.xlsx", "BE_SOFIPOS_202606.xlsx"]

# Nombres de las hojas y columna con el total del sector en el boletín.
HOJA_BE = "Estados_Financieros"
HOJA_UNIVERSO = "Universo"
COL_TOTAL_SECTOR = "E"

# Parámetros para interpretar los datos del tablero.
TIPO_SALDO = "130"  # Saldos sin consolidar.
CLAVE_SECTOR = "27"  # Fila agregada del sistema; no representa una SOFIPO.
MILES = 1000.0

# Relación entre el número del mes y su nombre, y su inversa.
MESES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}
MES_NUM = {nombre: numero for numero, nombre in MESES.items()}

# Años y meses que se incluirán en el libro.
ANIOS = {
    2025: tuple(range(1, 13)),
    2026: tuple(range(1, 9)),
}

# Relaciona las claves CNBV del tablero con los nombres internos de los campos.
CONCEPTOS = {
    "100000000000": "activos",
    "272724": "cartera_e12",  # Cartera de crédito: etapas 1 y 2.
    "100800104001": "cartera_e1",
    "100800104002": "cartera_e2",
    "100800104003": "cartera_e3",
    "130800103001": "cartera_total_cruda",
    "100800103003": "estimaciones",
    "500200101001": "ingresos_int",
    "600400101002": "gastos_int",
    "601800301009": "gastos_adm",
    "330000000006": "resultado_neto",  # Acumulado del año (YTD).
    "200200001001": "captacion",  # Captación tradicional.
    "911300000000": "capital_neto",
    "910700000000": "requerimiento",
    "27200043": "rn_12m",  # Resultado neto de los últimos 12 meses.
    "27200044": "activo_prom_12m",
    "27200051": "capital_prom_12m",
    # Indicadores publicados, utilizados únicamente para validación.
    "27200093": "nicap_pub",
    "27200094": "icap_pub",
    "27200005": "imor_pub",
    "27200013": "icor_pub",
    "27200001": "roa_pub",
    "27200002": "roe_pub",
}

# Campos para los que se calculan o leen los totales del sector.
CAMPOS_SECTOR = [
    "activos",
    "cartera_e12",
    "cartera_e3",
    "cartera_total_cruda",
    "estimaciones",
    "ingresos_int",
    "gastos_int",
    "gastos_adm",
    "resultado_neto",
    "captacion",
    "capital_neto",
    "requerimiento",
    "rn_12m",
    "activo_prom_12m",
    "capital_prom_12m",
]

# Texto que se muestra cuando un dato no está disponible.
ND = "n.d."

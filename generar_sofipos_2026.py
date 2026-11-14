#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera SOFIPOS_2026_Datos.xlsx con la información mensual de las 35 SOFIPOS:
enero-diciembre de 2025 y enero-agosto de 2026.

Fuente de las cifras
--------------------
Tablero de información financiera de la CNBV, descarga "sh_datos_csv_27_i"
(serie mensual 201601-202608, saldos sin consolidar, tipo_saldo = 130):

  entradas/tablero/sh_datos_csv_27_i/027_datos.csv
  entradas/tablero/sh_datos_csv_27_i/catalogo_conceptos_027.xlsx
  entradas/tablero/sh_datos_csv_27_i/catalogo_instituciones_027.xlsx

El único dato que se sigue tomando del boletín es la hoja "Universo"
(Federación, Nivel de operación, Fecha de autorización y Estatus):

  entradas/BE_SOFIPOS_202606.xlsx, hoja Universo

El CSV viene en pesos; el libro se publica en MILES de pesos.

Hojas del libro
---------------
  1. Catálogo SOFIPOS   -> 35 entidades autorizadas en operación
  2. Datos Fuente 2026   -> enero-agosto 2026, 35 entidades por mes (280 filas)
  3. Datos Fuente 2025   -> enero-diciembre 2025, 35 entidades por mes (420 filas)
  4. Indicadores 2026    -> 280 filas, fórmulas contra Datos Fuente 2026
  5. Indicadores 2025    -> 420 filas, fórmulas contra Datos Fuente 2025
  6. Diccionario         -> 11 indicadores, fórmula, grupo y descripción
  7. Semáforos           -> 10 bloques y 31 reglas de formato condicional

Método
------
- La fila con institución "27" es el total del sector: se excluye de las sumas.
- "Periodo anterior" es el mes inmediato anterior. Enero de 2025 se compara
  contra diciembre de 2024; enero de 2026 contra diciembre de 2025.
- Las columnas de periodo anterior (captación y cartera total) van como DATO,
  no como fórmula, para que cada hoja sea autocontenida.
- Los estados de resultados (ingresos, gastos, gastos de administración y
  resultado neto) se publican acumulados del año (YTD), igual que el tablero
  y que el libro anterior.
- ROA y ROE se calculan con los insumos de 12 meses del propio tablero:
    ROA = Resultado Neto 12M / Activo promedio 12M
    ROE = Resultado Neto 12M / Capital contable promedio 12M
  Reproducen exactamente los indicadores publicados 27200001 / 27200002.
- NICAP se calcula como Capital Neto / Requerimiento (coincide con 27200093).
- Cuando el margen financiero es cero o negativo, % Gastos de Administración
  se muestra como "n.a." (no aplicable) y se marca en rojo.
- Cuando la fuente no publica un dato se escribe "n.d."; no se imputan ceros.

Uso:
    .venv/bin/python generar_sofipos_2026.py
    .venv/bin/python generar_sofipos_2026.py --entradas otra --salidas otra
    .venv/bin/python generar_sofipos_2026.py --out otra_ruta.xlsx
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import posixpath
import re
import sys
import tempfile
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Configuración
# --------------------------------------------------------------------------

AQUI = Path(__file__).resolve().parent
DIR_ENTRADAS = AQUI / "entradas"
DIR_SALIDAS = AQUI / "salidas"
SALIDA_DEFECTO = DIR_SALIDAS / "SOFIPOS_2026_Datos.xlsx"

BE_UNIVERSO = "BE_SOFIPOS_202606.xlsx"
CSV_TABLERO = Path("tablero") / "sh_datos_csv_27_i" / "027_datos.csv"
BE_VALIDACION = ["BE_SOFIPOS_202603.xlsx", "BE_SOFIPOS_202606.xlsx"]

HOJA_BE = "Estados_Financieros"
HOJA_UNIVERSO = "Universo"
COL_TOTAL_SECTOR = "E"

TIPO_SALDO = "130"          # saldos sin consolidar
CLAVE_SECTOR = "27"         # fila agregada del sistema: no es una SOFIPO
MILES = 1000.0

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}
MES_NUM = {v: k for k, v in MESES.items()}

# Años y meses que entran al libro
ANIOS = {2025: tuple(range(1, 13)), 2026: tuple(range(1, 9))}

# Conceptos del tablero (clave CNBV -> campo interno)
CONCEPTOS = {
    "100000000000": "activos",
    "272724": "cartera_e12",             # Cartera (E1 + E2)
    "100800104001": "cartera_e1",
    "100800104002": "cartera_e2",
    "100800104003": "cartera_e3",
    "130800103001": "cartera_total_cruda",
    "100800103003": "estimaciones",
    "500200101001": "ingresos_int",
    "600400101002": "gastos_int",
    "601800301009": "gastos_adm",
    "330000000006": "resultado_neto",    # acumulado del año (YTD)
    "200200001001": "captacion",         # captación tradicional
    "911300000000": "capital_neto",
    "910700000000": "requerimiento",
    "27200043": "rn_12m",                # resultado neto 12 meses
    "27200044": "activo_prom_12m",
    "27200051": "capital_prom_12m",
    # Indicadores publicados: solo se usan para validar
    "27200093": "nicap_pub",
    "27200094": "icap_pub",
    "27200005": "imor_pub",
    "27200013": "icor_pub",
    "27200001": "roa_pub",
    "27200002": "roe_pub",
}

CAMPOS_SECTOR = [
    "activos", "cartera_e12", "cartera_e3", "cartera_total_cruda",
    "estimaciones", "ingresos_int", "gastos_int", "gastos_adm",
    "resultado_neto", "captacion", "capital_neto", "requerimiento",
    "rn_12m", "activo_prom_12m", "capital_prom_12m",
]

ND = "n.d."

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


# --------------------------------------------------------------------------
# Catálogo: denominación social registral + sitio web (por Clave CASFIM)
# Investigado el 11/nov/2026 en CONDUSEF/SIPRES (clave registral = CASFIM)
# y, para Nu México, en el oficio de SEGOB/DOF.
# --------------------------------------------------------------------------

CATALOGO = {
    "027001": ("Fincomún, Servicios Financieros, S.A. de C.V., Sociedad Financiera Popular", "https://www.fincomun.com.mx"),
    "027003": ("UNAGRA, S.A. de C.V., Sociedad Financiera Popular", "https://www.unagra.com.mx"),
    "027004": ("Financiera Mexicana para el Desarrollo Rural, S.A. de C.V., S. F. P.", "https://www.finamigo.com.mx"),
    "027005": ("Financiera Tamazula, S.A. de C.V., S.F.P.", "https://www.fitaserviciosfinancieros.com.mx"),
    "027008": ("Financiera Monte de Piedad, S.A. de C.V., S.F.P.", "https://financieramontedepiedad.com.mx"),
    "027009": ("Klar Technologies, S.A. de C.V., S.F.P.", "https://www.klar.mx"),
    "027012": ("Ictineo Plataforma, S.A. de C.V., Sociedad Financiera Popular", "https://bbdcdigital.com.mx"),
    "027013": ("Solución Asea, S.A. de C.V., S.F.P.", "https://www.credicapital.com.mx"),
    "027014": ("NU MÉXICO FINANCIERA, S.A. de C.V., S.F.P.", "https://www.nu.com.mx"),
    "027015": ("Operaciones de Tu Lado, S.A. de C.V., Sociedad Financiera Popular", "https://avantthya.com.mx"),
    "027018": ("Únete Financiera de Allende, S.A. de C.V., S.F.P.", "https://www.unetefinanciera.com.mx"),
    "027019": ("Sociedad de Alternativas Económicas, S.A. de C.V., S.F.P.", "https://levainvierte.com/"),
    "027020": ("Multiplica México, S.A. de C.V., S.F.P.", "http://www.multiplicamexico.com"),
    "027025": ("Financiera del Sector Social, S.A. de C.V., S.F.P.", "https://www.finsocial.mx"),
    "027026": ("Administradora de Caja Bienestar, S.A. de C.V., S.F.P.", "https://www.ahorrosbienestar.com"),
    "027028": ("Acción y Evolución, S.A. de C.V., S.F.P.", "https://www.realfinanciera.com"),
    "027029": ("Opciones Empresariales del Noreste, S.A. de C.V., S.F.P.", "https://www.aspintegraopciones.com"),
    "027030": ("Financiera Broxel, S.A. de C.V., S.F.P.", "https://broxel.com/financiera/"),
    "027031": ("Caja Progressa, S.A. de C.V., S.F.P.", "https://www.progressa.com.mx"),
    "027032": ("Stori México, S.A. de C.V. S.F.P.", "https://www.storicard.com"),
    "027033": ("Libertad Servicios Financieros, S.A. de C.V., S.F.P.", "https://www.libertad.com.mx"),
    "027034": ("Capital Activo, S.A. de C.V., S.F.P.", "https://www.capitalactivo.com.mx"),
    "027035": ("J.P. Sofiexpress, S.A. de C.V., S.F.P.", "http://www.sofiexpress.com.mx"),
    "027036": ("Fondea Technologies, S.A. de C.V., S.F.P.", "https://www.fondeadora.com"),
    "027037": ("Financiera Más, S.A. de C.V., S.F.P.", "https://www.financieramas.com"),
    "027040": ("Impulso para el Desarrollo de México, S.A. de C.V., S.F.P.", "https://www.impulsomexico.net"),
    "027042": ("Paso Seguro Creando Futuro, S.A. de C.V., S.F.P.", "http://www.pasoseguro.com.mx"),
    "027044": ("Financiera Súmate, S.A. de C.V., S.F.P.", "https://www.financierasumate.com.mx"),
    "027045": ("Ku-Bo Financiero, S.A. de C.V., S.F.P.", "https://www.kubofinanciero.com"),
    "027046": ("Financiera Sustentable de México, S.A. de C.V., S.F.P.", "https://finsus.mx"),
    "027047": ("Crediclub, S.A. de C.V., S.F.P.", "https://www.crediclub.com"),
    "027049": ("SFP Porvenir, S.A. de C.V., S.F.P.", "https://www.porvenirsofipo.mx"),
    "027050": ("Préstamo Empresarial Oportuno, S.A. de C.V., S.F.P.", "https://www.premo.mx"),
    "027051": ("Grupo Regional de Negocios, S.A. de C.V., S.F.P.", "https://grensa.com"),
    "027052": ("Trafalgar SFP, S.A. de C.V., Sociedad Financiera Popular", "https://trafalgar.mx"),
}


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------

def _col_letra(ref: str) -> str:
    m = re.match(r"[A-Z]+", ref or "")
    return m.group(0) if m else ""


def _a_numero(texto):
    if texto is None:
        return None
    if isinstance(texto, (int, float)):
        return float(texto)
    try:
        return float(str(texto).replace(",", "").replace("$", "").strip())
    except ValueError:
        return None


def _norm(texto) -> str:
    s = unicodedata.normalize("NFKD", str(texto or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def normalizar_clave(valor) -> str:
    digitos = re.sub(r"\D", "", str(valor))
    return digitos.zfill(6)


def normalizar_entidad_legal(nombre: str) -> str:
    """Conserva el nombre registral y unifica solo el sufijo societario."""
    s = re.sub(r"\s+", " ", str(nombre)).strip().rstrip(",")
    s = re.sub(r",?\s*Sociedad\s+Financiera\s+Popular\.?$", "", s, flags=re.I)
    s = re.sub(r",?\s*S\.\s*F\.\s*P\.?$", "", s, flags=re.I)
    s = s.strip().rstrip(",")
    if re.search(r"S\.A\.\s*de\s*C\.\s*V\.?$", s, re.I):
        s = re.sub(r"S\.A\.\s*de\s*C\.\s*V\.?$", "S.A. de C.V.", s, flags=re.I)
    else:
        s = f"{s}, S.A. de C.V."
    return f"{s}, S.F.P."


def trimestre(mes: int) -> str:
    return f"T{((mes - 1) // 3) + 1}"


def fecha_mes(anio: int, mes: int) -> str:
    return f"{anio}{mes:02d}"


def fecha_anterior(fecha: str) -> str:
    anio, mes = int(fecha[:4]), int(fecha[4:])
    mes -= 1
    if mes == 0:
        anio, mes = anio - 1, 12
    return f"{anio}{mes:02d}"


def a_miles(valor):
    """Pesos -> miles de pesos, redondeado a 2 decimales."""
    if valor is None:
        return None
    return round(float(valor) / MILES, 2)


# --------------------------------------------------------------------------
# Lectura del .xlsx (solo biblioteca estándar)
# --------------------------------------------------------------------------

def _leer_shared_strings(z: zipfile.ZipFile) -> list[str]:
    try:
        raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(t.text or "" for t in si.iter(NS + "t")) for si in raiz.findall(NS + "si")]


def _rutas_de_hojas(z: zipfile.ZipFile) -> dict[str, str]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    mapa = {r.get("Id"): r.get("Target") for r in rels.findall(NS_REL + "Relationship")}
    rutas = {}
    for hoja in wb.find(NS + "sheets"):
        destino = mapa[hoja.get(RNS + "id")]
        if destino.startswith("/"):
            ruta = destino.lstrip("/")
        else:
            ruta = posixpath.normpath(posixpath.join("xl", destino))
        rutas[hoja.get("name")] = ruta
    return rutas


def _leer_hoja(z: zipfile.ZipFile, ruta: str, shared: list[str]) -> dict[int, dict[str, object]]:
    raiz = ET.fromstring(z.read(ruta))
    filas: dict[int, dict[str, object]] = {}
    for fila in raiz.iter(NS + "row"):
        numero = int(fila.get("r"))
        celdas: dict[str, object] = {}
        for celda in fila.findall(NS + "c"):
            col = _col_letra(celda.get("r"))
            tipo = celda.get("t")
            v = celda.find(NS + "v")
            if tipo == "inlineStr":
                valor = "".join(t.text or "" for t in celda.iter(NS + "t"))
            elif v is None:
                continue
            elif tipo == "s":
                valor = shared[int(v.text)]
            elif tipo in ("str", "e"):
                valor = v.text
            else:
                num = _a_numero(v.text)
                valor = num if num is not None else v.text
            celdas[col] = valor
        filas[numero] = celdas
    return filas


def _abrir_hoja(ruta: Path, nombre: str) -> dict[int, dict[str, object]]:
    with zipfile.ZipFile(ruta) as z:
        shared = _leer_shared_strings(z)
        rutas = _rutas_de_hojas(z)
        if nombre not in rutas:
            raise SystemExit(f"No se encontró la hoja '{nombre}' en {ruta.name}. Hojas: {sorted(rutas)}")
        return _leer_hoja(z, rutas[nombre], shared)


def _fila_encabezado(hoja: dict[int, dict[str, object]], texto: str) -> int:
    objetivo = _norm(texto)
    for num in sorted(hoja):
        if _norm(hoja[num].get("D", "")).startswith(objetivo):
            return num
    raise SystemExit(f"No se localizó el encabezado '{texto}' en la columna D.")


def _columnas_sofipo(codigos: dict[str, object]) -> list[str]:
    columnas = []
    for col, valor in codigos.items():
        if col in (COL_TOTAL_SECTOR, "D"):
            continue
        if isinstance(valor, float):
            valor = int(valor)
        if re.fullmatch(r"\d{4,6}", str(valor).strip()):
            columnas.append(col)
    columnas.sort(key=lambda c: (len(c), c))
    return columnas


# --------------------------------------------------------------------------
# Lectura del boletín BE (solo para el Universo y para validar)
# --------------------------------------------------------------------------

ETIQUETAS_BE = {
    "activos": "activo",
    "cartera_1_2": "cartera de credito etapa 1 + etapa 2",
    "cartera_3": "cartera de credito etapa 3",
    "estimaciones": "estimaciones preventivas para riesgos crediticios",
    "ingresos_int": "ingresos por intereses",
    "gastos_int": "gastos por intereses",
    "gastos_adm": "gastos de administracion y promocion",
}


@dataclass
class Boletin:
    registros: list[dict]
    totales: dict
    filas: dict[str, int] = field(default_factory=dict)


def leer_be_sofipos(ruta: Path) -> Boletin:
    """Lee Estados_Financieros localizando cada concepto por su etiqueta."""
    hoja = _abrir_hoja(ruta, HOJA_BE)
    fila_enc = _fila_encabezado(hoja, "Clave CASFIM")

    codigos = hoja[fila_enc]
    nombres = hoja.get(fila_enc + 1, {})
    federaciones = hoja.get(fila_enc + 2, {})
    niveles = hoja.get(fila_enc + 3, {})

    filas: dict[str, int] = {}
    for campo, etiqueta in ETIQUETAS_BE.items():
        for num in sorted(hoja):
            if num <= fila_enc + 3:
                continue
            if _norm(hoja[num].get("D", "")) == etiqueta:
                filas[campo] = num
                break
        if campo not in filas:
            raise SystemExit(f"No se localizó la etiqueta '{etiqueta}' en {ruta.name}.")

    columnas = _columnas_sofipo(codigos)
    registros = []
    for col in columnas:
        crudo = codigos[col]
        clave = normalizar_clave(int(crudo) if isinstance(crudo, float) else crudo)
        reg = {
            "clave": clave,
            "nombre": str(nombres.get(col, "")).strip(),
            "federacion": str(federaciones.get(col, "")).strip(),
            "nivel": str(niveles.get(col, "")).strip(),
        }
        for campo, fila in filas.items():
            valor = hoja.get(fila, {}).get(col)
            reg[campo] = valor.strip() if isinstance(valor, str) else valor
        registros.append(reg)

    totales = {campo: hoja.get(fila, {}).get(COL_TOTAL_SECTOR) for campo, fila in filas.items()}
    return Boletin(registros=registros, totales=totales, filas=filas)


def leer_universo(ruta: Path) -> list[dict]:
    hoja = _abrir_hoja(ruta, HOJA_UNIVERSO)
    fila_enc = _fila_encabezado(hoja, "Clave CASFIM")

    registros = []
    for num in sorted(hoja):
        if num <= fila_enc:
            continue
        fila = hoja[num]
        consecutivo = fila.get("C")
        clave = fila.get("D")
        estatus = _norm(fila.get("J", ""))
        # Solo la primera sección: sociedades autorizadas EN OPERACIÓN.
        if estatus != "en operacion":
            continue
        if _a_numero(consecutivo) is None or not re.fullmatch(r"\d{4,6}", str(clave).strip().split(".")[0]):
            continue
        serial = _a_numero(fila.get("F"))
        fecha = dt.date(1899, 12, 30) + dt.timedelta(days=int(serial)) if serial else None
        registros.append({
            "consecutivo": int(_a_numero(consecutivo)),
            "clave": normalizar_clave(int(clave) if isinstance(clave, float) else clave),
            "nombre": str(fila.get("E", "")).strip(),
            "fecha_autorizacion": fecha,
            "nivel": str(fila.get("G", "")).strip(),
            "federacion": str(fila.get("I", "")).strip(),
            "estatus": str(fila.get("J", "")).strip(),
        })
    return registros


# --------------------------------------------------------------------------
# Lectura del tablero (CSV mensual en pesos)
# --------------------------------------------------------------------------

@dataclass
class Tablero:
    valores: dict[tuple[str, str], dict[str, float]]   # (clave, fecha) -> campo -> pesos
    sector: dict[tuple[str, str], float]               # (fecha, campo) -> pesos

    def dato(self, clave: str, fecha: str, campo: str):
        return self.valores.get((clave, fecha), {}).get(campo)


def leer_tablero(ruta: Path, fechas: set[str]) -> Tablero:
    """Lee el CSV una sola vez y conserva solo los conceptos y meses que usa el libro."""
    valores: dict[tuple[str, str], dict[str, float]] = {}
    sector: dict[tuple[str, str], float] = {}
    leidas = 0
    with open(ruta, newline="", encoding="utf-8-sig") as f:
        for x in csv.DictReader(f):
            if x.get("tipo_saldo") != TIPO_SALDO:
                continue
            fecha = x.get("fecha", "")
            if fecha not in fechas:
                continue
            campo = CONCEPTOS.get(x.get("concepto", ""))
            if not campo:
                continue
            valor = _a_numero(x.get("saldo_se"))
            if valor is None:
                continue
            institucion = str(x.get("institucion", "")).strip()
            leidas += 1
            if institucion == CLAVE_SECTOR:
                sector[(fecha, campo)] = valor
                continue
            clave = normalizar_clave(institucion)
            valores.setdefault((clave, fecha), {})[campo] = valor
    print(f"  tablero: {leidas:,} registros útiles de {ruta.name} ({len(fechas)} meses)")
    return Tablero(valores=valores, sector=sector)


# --------------------------------------------------------------------------
# Preparación de filas por año
# --------------------------------------------------------------------------

def _cartera_e12(v: dict) -> float | None:
    """Cartera etapa 1 + etapa 2, con respaldo en sus componentes o en el total."""
    if v.get("cartera_e12") is not None:
        return v["cartera_e12"]
    if v.get("cartera_e1") is not None and v.get("cartera_e2") is not None:
        return v["cartera_e1"] + v["cartera_e2"]
    if v.get("cartera_total_cruda") is not None and v.get("cartera_e3") is not None:
        return v["cartera_total_cruda"] - v["cartera_e3"]
    return None


def _cartera_e3(v: dict, e12: float | None) -> float | None:
    """Cartera etapa 3; si no viene desglosada, total de cartera menos E1+E2."""
    if v.get("cartera_e3") is not None:
        return v["cartera_e3"]
    if v.get("cartera_total_cruda") is not None and e12 is not None:
        return v["cartera_total_cruda"] - e12
    return None


def preparar_filas(tablero: Tablero, universo: list[dict], nombres: dict[str, str],
                   anio: int) -> list[dict]:
    filas: list[dict] = []
    for mes in ANIOS[anio]:
        fecha = fecha_mes(anio, mes)
        previa = fecha_anterior(fecha)
        for reg in universo:
            clave = reg["clave"]
            v = tablero.valores.get((clave, fecha), {})
            p = tablero.valores.get((clave, previa), {})

            e12 = _cartera_e12(v)
            e3 = _cartera_e3(v, e12)
            e12_prev = _cartera_e12(p)
            e3_prev = _cartera_e3(p, e12_prev)
            cartera_prev = (e12_prev + e3_prev
                            if e12_prev is not None and e3_prev is not None else None)

            filas.append({
                "no": reg["consecutivo"],
                "clave": clave,
                "sofipo": nombres.get(clave) or reg["nombre"],
                "federacion": reg.get("federacion") or "Sin federación",
                "nivel": reg.get("nivel") or "Sin nivel",
                "anio": anio,
                "mes": MESES[mes],
                "trimestre": trimestre(mes),
                "activos": a_miles(v.get("activos")),
                "cartera_e12": a_miles(e12),
                "cartera_e3": a_miles(e3),
                "estimaciones": a_miles(v.get("estimaciones")),
                "ingresos_int": a_miles(v.get("ingresos_int")),
                "gastos_int": a_miles(v.get("gastos_int")),
                "gastos_adm": a_miles(v.get("gastos_adm")),
                "resultado_neto": a_miles(v.get("resultado_neto")),
                "captacion": a_miles(v.get("captacion")),
                "captacion_prev": a_miles(p.get("captacion")),
                "cartera_total_prev": a_miles(cartera_prev),
                "capital_neto": a_miles(v.get("capital_neto")),
                "requerimiento": a_miles(v.get("requerimiento")),
                "rn_12m": a_miles(v.get("rn_12m")),
                "activo_prom_12m": a_miles(v.get("activo_prom_12m")),
                "capital_prom_12m": a_miles(v.get("capital_prom_12m")),
                # Indicadores publicados (por ciento), solo para validar
                "nicap_pub": v.get("nicap_pub"),
                "imor_pub": v.get("imor_pub"),
                "icor_pub": v.get("icor_pub"),
                "roa_pub": v.get("roa_pub"),
                "roe_pub": v.get("roe_pub"),
                "cartera_total_cruda": a_miles(v.get("cartera_total_cruda")),
                "tiene_datos": bool(v),
            })
    return filas


# --------------------------------------------------------------------------
# Validaciones de datos (antes de escribir el libro)
# --------------------------------------------------------------------------

def _num(x) -> bool:
    return isinstance(x, (int, float))


def validar_datos(tablero: Tablero, filas: dict[int, list[dict]], universo: list[dict],
                  be_por_archivo: dict[str, Boletin]) -> bool:
    ok = True
    print("\n=== Validación de datos ===")
    claves = [u["clave"] for u in universo]

    # 1. 35 entidades en cada mes
    print("\n  Cobertura por mes:")
    for anio, lista in filas.items():
        por_mes: dict[str, int] = {}
        for f in lista:
            por_mes[f["mes"]] = por_mes.get(f["mes"], 0) + 1
        malos = {m: n for m, n in por_mes.items() if n != len(universo)}
        print(f"    {anio}: {len(lista)} filas · {len(por_mes)} meses · "
              f"{'todas con ' + str(len(universo)) if not malos else malos}")
        if malos:
            ok = False

    # 2. Consistencia del propio tablero: suma de todas las entidades vs fila "27"
    print("\n  Consistencia del tablero: suma de entidades vs fila 27 (total del sector):")
    campos = ["activos", "cartera_e12", "cartera_e3", "cartera_total_cruda",
              "captacion", "capital_neto", "requerimiento"]
    etiquetas = {
        "activos": "Activos", "cartera_e12": "Cartera E1+E2", "cartera_e3": "Cartera E3",
        "cartera_total_cruda": "Cartera total", "captacion": "Captación",
        "capital_neto": "Capital Neto", "requerimiento": "Requerimiento",
    }
    for campo in campos:
        comparaciones = 0
        peor = 0.0
        for anio in ANIOS:
            for mes in ANIOS[anio]:
                fecha = fecha_mes(anio, mes)
                esperado = tablero.sector.get((fecha, campo))
                if esperado is None:
                    continue
                suma = sum(v.get(campo, 0.0) for (clave, f), v in tablero.valores.items()
                           if f == fecha)
                comparaciones += 1
                peor = max(peor, abs(suma - esperado))
        estado = "OK" if peor <= 1.0 else f"DIF {peor:,.2f}"
        if peor > 1.0:
            ok = False
        print(f"    {etiquetas[campo]:<18} {comparaciones:>2} meses · peor diferencia {peor:,.2f} pesos · {estado}")

    # 3. Contra el boletín BE de marzo y junio de 2026 (7 variables)
    print("\n  Tablero vs boletín BE (marzo y junio 2026, miles de pesos):")
    pares = [
        ("activos", "activos", "Activos"),
        ("cartera_e12", "cartera_1_2", "Cartera E1+E2"),
        ("cartera_e3", "cartera_3", "Cartera E3"),
        ("estimaciones", "estimaciones", "Estimaciones"),
        ("ingresos_int", "ingresos_int", "Ingresos por intereses"),
        ("gastos_int", "gastos_int", "Gastos por intereses"),
        ("gastos_adm", "gastos_adm", "Gastos de administración"),
    ]
    excepciones_be = {("202606", "027019", "gastos_adm")}
    for archivo, boletin in be_por_archivo.items():
        fecha = archivo.split("_")[-1].replace(".xlsx", "")  # 202603 / 202606
        anio, mes = int(fecha[:4]), int(fecha[4:])
        por_clave = {r["clave"]: r for r in boletin.registros}
        filas_mes = {f["clave"]: f for f in filas[anio]
                     if f["mes"] == MESES[mes]}
        print(f"    {MESES[mes]} {anio} ({'serie reexpresada' if fecha == '202603' else 'corte definitivo'}):")
        for campo_tab, campo_be, etiqueta in pares:
            comparaciones = 0
            peor = 0.0
            peor_clave = None
            discrepancias = []
            for clave, f in filas_mes.items():
                be = por_clave.get(clave)
                if not be or not _num(f.get(campo_tab)) or not _num(be.get(campo_be)):
                    continue
                comparaciones += 1
                dif = abs(f[campo_tab] - be[campo_be])
                if dif > 0.02:
                    discrepancias.append(clave)
                if dif > peor:
                    peor, peor_clave = dif, clave
            if fecha == "202606":
                no_documentadas = [c for c in discrepancias
                                   if (fecha, c, campo_tab) not in excepciones_be]
                if no_documentadas:
                    ok = False
                if not discrepancias:
                    estado = "OK"
                elif no_documentadas:
                    estado = f"REVISAR {len(no_documentadas)} entidades"
                else:
                    estado = "OK · excepción documentada 027019 SAE"
            else:
                estado = ("OK" if not discrepancias
                          else f"{len(discrepancias)} diferencias por reexpresión")
            extra = f" · peor {peor:,.4f} ({peor_clave})" if peor_clave else ""
            print(f"      {etiqueta:<24} {comparaciones:>2} entidades{extra} · {estado}")

    # 4. ROA / ROE calculados contra los indicadores publicados
    #    Se usan los importes originales en pesos para no arrastrar el
    #    redondeo a miles; el único margen es el redondeo del indicador (2 dec.).
    def crudo(f):
        return tablero.valores.get(
            (f["clave"], fecha_mes(f["anio"], MES_NUM[f["mes"]])), {}
        )

    print("\n  ROA / ROE calculados (RN 12M / promedio 12M) vs publicados:")
    for nombre, num, den, pub in [
        ("ROA", "rn_12m", "activo_prom_12m", "roa_pub"),
        ("ROE", "rn_12m", "capital_prom_12m", "roe_pub"),
    ]:
        comparaciones = 0
        peor = 0.0
        peor_fila = None
        for anio, lista in filas.items():
            for f in lista:
                v = crudo(f)
                if not (_num(v.get(num)) and _num(v.get(den)) and _num(f.get(pub))):
                    continue
                if v.get(den) == 0:
                    continue
                calculado = v[num] / v[den] * 100
                dif = abs(calculado - f[pub])
                comparaciones += 1
                if dif > peor:
                    peor, peor_fila = dif, f
        estado = "OK" if peor <= 0.02 else f"DIF {peor:,.4f}"
        if peor > 0.02:
            ok = False
        extra = f" ({peor_fila['clave']} {peor_fila['mes']} {peor_fila['anio']})" if peor_fila else ""
        print(f"    {nombre}: {comparaciones:>3} comparaciones · peor diferencia {peor:,.4f} pts · {estado}{extra}")

    # 5. NICAP, IMOR e ICOR calculados contra los publicados
    print("\n  NICAP / IMOR / ICOR calculados vs publicados:")
    pruebas = [
        ("NICAP", lambda v: v["capital_neto"] / v["requerimiento"] * 100,
         "nicap_pub", ("capital_neto", "requerimiento")),
        ("IMOR", lambda v: v["cartera_e3"] / (v["cartera_e12"] + v["cartera_e3"]) * 100,
         "imor_pub", ("cartera_e3", "cartera_e12")),
        ("ICOR", lambda v: abs(v["estimaciones"]) / v["cartera_e3"] * 100,
         "icor_pub", ("estimaciones", "cartera_e3")),
    ]
    for nombre, calculo, campo_pub, requeridos in pruebas:
        comparaciones = 0
        peor = 0.0
        peor_fila = None
        for anio, lista in filas.items():
            for f in lista:
                v = crudo(f)
                if not all(_num(v.get(c)) for c in requeridos) or not _num(f.get(campo_pub)):
                    continue
                if nombre == "NICAP" and not v["requerimiento"]:
                    continue
                if nombre in ("IMOR", "ICOR") and (
                    v["cartera_e3"] == 0
                    or (nombre == "IMOR" and (v["cartera_e12"] + v["cartera_e3"]) == 0)
                ):
                    continue
                dif = abs(calculo(v) - f[campo_pub])
                comparaciones += 1
                if dif > peor:
                    peor, peor_fila = dif, f
        estado = "OK" if peor <= 0.02 else f"DIF {peor:,.4f}"
        if peor > 0.02:
            ok = False
        extra = f" ({peor_fila['clave']} {peor_fila['mes']} {peor_fila['anio']})" if peor_fila else ""
        print(f"    {nombre}: {comparaciones:>3} comparaciones · peor diferencia {peor:,.4f} pts · {estado}{extra}")

    # 6. Cartera total = E1+E2+E3
    print("\n  Cartera total vs boletín (E1+E2+E3):")
    for anio, lista in filas.items():
        comparaciones = 0
        peor = 0.0
        for f in lista:
            if all(_num(f.get(c)) for c in ("cartera_e12", "cartera_e3", "cartera_total_cruda")):
                comparaciones += 1
                peor = max(peor, abs((f["cartera_e12"] + f["cartera_e3"]) -
                                     f["cartera_total_cruda"]))
        print(f"    {anio}: {comparaciones} entidades-mes · peor diferencia "
              f"{peor:,.4f} miles · {'OK' if peor <= 0.02 else 'REVISAR'}")
        if peor > 0.02:
            ok = False

    _ = claves
    return ok


# --------------------------------------------------------------------------
# Estilos
# --------------------------------------------------------------------------

AZUL = "1F3864"
GRIS = "F2F2F2"
AMBAR = "FFF2CC"
BORDE = Side(style="thin", color="BFBFBF")

FMT_MILES = '#,##0.00;[Red]-#,##0.00'
FMT_PORC = "0.00%"
FMT_ANIO = "0"
FMT_TEXTO = "@"
FMT_FECHA = "dd-mmm-yyyy"

# Paleta de semáforos con contraste AAA de WCAG 2.2 (>= 7:1 texto/fondo).
# Verde 7.21:1 · Amarillo 7.21:1 · Rojo 7.08:1.
VERDE_FILL = PatternFill("solid", fgColor="D1E7DD")
AMARILLO_FILL = PatternFill("solid", fgColor="FFF3CD")
ROJO_FILL = PatternFill("solid", fgColor="F8D7DA")
VERDE_FONT = Font(color="0F5132", bold=True)
AMARILLO_FONT = Font(color="664D03", bold=True)
ROJO_FONT = Font(color="842029", bold=True)

COLORES = {
    "verde": (VERDE_FILL, VERDE_FONT),
    "amarillo": (AMARILLO_FILL, AMARILLO_FONT),
    "rojo": (ROJO_FILL, ROJO_FONT),
}


@dataclass(frozen=True)
class Ficha:
    """Definición única de cada indicador: diccionario + semáforo."""
    indicador: str
    columna: str
    formula: str
    grupo: str
    descripcion: str
    tramos: tuple[tuple[str, str, str], ...] = ()   # (texto, condición Excel, color)


# Las condiciones se evalúan sobre fracciones porque las celdas llevan formato %.
# Van envueltas en ISNUMBER: en Excel cualquier texto es mayor que cualquier
# número, así que sin esa guarda los "n.d." se pintarían de verde. La única
# excepción es COND_NA, que aplica justamente a texto y se evalúa sin la guarda.
COND_NA = '{c}="n.a."'

FICHAS: list[Ficha] = [
    Ficha(
        indicador="NICAP",
        columna="F",
        formula="Capital Neto / Requerimiento de capital",
        grupo="Solvencia",
        descripcion="Mide si el capital neto cubre los riesgos asumidos. El mínimo "
                    "regulatorio es 131%; por debajo de eso la sociedad incumple.",
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
        descripcion="Proporción de la cartera que ya no genera pago. Mide el "
                    "deterioro de la cartera: a mayor IMOR, peor calidad.",
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
        descripcion="Qué tanto cubren las estimaciones preventivas a la cartera "
                    "en etapa 3. Arriba de 100% la cartera vencida está cubierta por completo.",
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
        descripcion="Rendimiento anualizado del activo. Se calcula con el resultado "
                    "neto de los últimos 12 meses y el activo promedio de esos 12 meses; "
                    "reproduce el indicador ROA (27200001) que publica la CNBV.",
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
        descripcion="Rendimiento anualizado sobre el patrimonio. Se calcula con el "
                    "resultado neto y el capital contable promedio de los últimos 12 "
                    "meses; reproduce el indicador ROE (27200002) que publica la CNBV.",
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
        descripcion="Parte del ingreso por intereses que queda después del costo "
                    "de fondeo. Muestra la rentabilidad del negocio crediticio. "
                    "El margen financiero se toma de la columna N de Datos Fuente "
                    "(Ingresos por intereses − Gastos por intereses).",
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
        descripcion="Variación mensual de la captación tradicional (depósitos del "
                    "público). Mide si la SOFIPO está ganando o perdiendo clientes y "
                    "depósitos respecto del mes inmediato anterior.",
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
        descripcion="Variación mensual de la cartera de crédito total (etapas 1, 2 y 3). "
                    "Mide la expansión o contracción del negocio crediticio; una caída "
                    "señala contracción de la cartera.",
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
        descripcion="Resultado neto acumulado del año (YTD) en miles de pesos. "
                    "No es un cociente ni tiene semáforo: se traslada como parámetro "
                    "desde Datos Fuente para leerlo junto con los indicadores.",
        tramos=(),
    ),
    Ficha(
        indicador="Crédito Neto",
        columna="O",
        formula="Cartera Neta / Activos Totales",
        grupo="Estructura financiera",
        descripcion="Peso de la cartera neta dentro del activo total. Indica qué "
                    "tan concentrado está el balance en el negocio de crédito.",
        tramos=(
            ("70% – 85%", "AND({c}>=0.7,{c}<=0.85)", "verde"),
            ("60% – < 70% y > 85% – 90%",
             "OR(AND({c}>=0.6,{c}<0.7),AND({c}>0.85,{c}<=0.9))", "amarillo"),
            ("< 60% o > 90%", "OR({c}<0.6,{c}>0.9)", "rojo"),
        ),
    ),
    Ficha(
        indicador="% Gastos de Administración",
        columna="P",
        formula="Gastos de administración y promoción / Margen Financiero",
        grupo="Eficiencia operativa",
        descripcion="Costo operativo por cada peso de margen financiero. A menor "
                    "porcentaje, más eficiente es la estructura de gastos. El margen "
                    "financiero se toma de la columna N de Datos Fuente. Cuando el "
                    "margen es cero o negativo, la razón no aplica y se muestra n.a.",
        tramos=(
            ("< 60%", "{c}<0.6", "verde"),
            ("60% – 80%", "AND({c}>=0.6,{c}<=0.8)", "amarillo"),
            ("> 80%", "{c}>0.8", "rojo"),
            ("Margen ≤ 0 (n.a.)", COND_NA, "rojo"),
        ),
    ),
]

FICHAS_SEMAFORO = [f for f in FICHAS if f.tramos]
TOTAL_REGLAS = sum(len(f.tramos) for f in FICHAS_SEMAFORO)


# --------------------------------------------------------------------------
# Escritura del .xlsx
# --------------------------------------------------------------------------

FILA_ENC = 4
FILA_DATOS = 5

ENC_CATALOGO = ["No.", "Clave CASFIM", "SOFIPO", "Federación", "Entidad Legal",
                "Sitio Web", "Fecha de Autorización", "Nivel", "Estatus"]

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
    for j in range(1, columnas + 1):
        celda = ws.cell(row=fila, column=j)
        celda.font = Font(bold=True, color="FFFFFF", size=10)
        celda.fill = PatternFill("solid", fgColor=AZUL)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = Border(left=BORDE, right=BORDE, top=BORDE,
                              bottom=Side(style="medium", color=AZUL))
    ws.row_dimensions[fila].height = alto


def _titulo(ws, columnas: int, titulo: str, subtitulo: str, nota: str) -> None:
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
    for texto in textos:
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=columna_final)
        celda = ws.cell(row=fila, column=1, value=texto)
        celda.font = Font(size=8, color="595959")
        celda.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[fila].height = 26
        fila += 1


def _texto(ws, fila: int, col: int, valor, negrita=False, alineacion="left", color=None) -> None:
    celda = ws.cell(row=fila, column=col, value=valor)
    celda.number_format = FMT_TEXTO
    celda.font = Font(bold=negrita, italic=color is not None and color != "000000",
                      color=color) if color else Font(bold=negrita)
    celda.alignment = Alignment(horizontal=alineacion, vertical="center",
                                wrap_text=False, indent=1 if alineacion == "left" else 0)


def _celda_numero(ws, fila: int, col: int, valor, formato: str) -> None:
    """Escribe un número o, si no hay dato, el texto 'n.d.' en gris."""
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
    celda = ws.cell(row=fila, column=col, value=formula)
    celda.number_format = formato
    celda.font = Font(bold=True, color=AZUL)
    celda.alignment = Alignment(horizontal="right", vertical="center")
    celda.fill = PatternFill("solid", fgColor=AMBAR)


def _borde_y_banda(ws, fila: int, columnas: int, banda: bool) -> None:
    for j in range(1, columnas + 1):
        celda = ws.cell(row=fila, column=j)
        celda.border = Border(left=BORDE, right=BORDE, top=BORDE, bottom=BORDE)
        if banda and celda.fill.fgColor.rgb in (None, "00000000"):
            celda.fill = PatternFill("solid", fgColor=GRIS)


def _texto_envuelto(ws, fila: int, col: int, valor, alto: int = 30) -> None:
    celda = ws.cell(row=fila, column=col, value=valor)
    celda.number_format = FMT_TEXTO
    celda.font = Font(size=10)
    celda.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    ws.row_dimensions[fila].height = alto


# --------------------------------------------------------------------------
# Hoja: Catálogo SOFIPOS
# --------------------------------------------------------------------------

def escribir_catalogo(wb: Workbook, universo: list[dict], nombres: dict[str, str]) -> None:
    ws = wb.create_sheet("Catálogo SOFIPOS")
    ncols = len(ENC_CATALOGO)

    _titulo(
        ws, ncols,
        "SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Catálogo",
        "35 sociedades autorizadas en operación · Información al 30 de junio de 2026",
        "Autorización, nivel y estatus: BE_SOFIPOS_202606.xlsx, hoja Universo. "
        "Entidad Legal y Sitio Web: investigación en CONDUSEF/SIPRES y SEGOB/DOF.",
    )

    for j, encabezado in enumerate(ENC_CATALOGO, start=1):
        ws.cell(row=FILA_ENC, column=j, value=encabezado)
    _estilo_encabezado(ws, FILA_ENC, ncols, alto=40)

    for i, reg in enumerate(universo):
        r = FILA_DATOS + i
        crudo, sitio = CATALOGO[reg["clave"]]

        _celda_numero(ws, r, 1, reg["consecutivo"], "0")
        _texto(ws, r, 2, reg["clave"], alineacion="center")
        _texto(ws, r, 3, nombres.get(reg["clave"], reg["nombre"]), negrita=True)
        _texto(ws, r, 4, reg["federacion"])
        _texto(ws, r, 5, normalizar_entidad_legal(crudo))
        _texto(ws, r, 6, sitio, color="0563C1")

        fecha = ws.cell(row=r, column=7, value=reg["fecha_autorizacion"])
        fecha.number_format = FMT_FECHA
        fecha.alignment = Alignment(horizontal="center", vertical="center")
        fecha.font = Font()

        _texto(ws, r, 8, reg["nivel"], alineacion="center")
        _texto(ws, r, 9, reg["estatus"], alineacion="center")

        enlace = ws.cell(row=r, column=6)
        enlace.hyperlink = sitio
        enlace.font = Font(color="0563C1", underline="single", size=9)

        _borde_y_banda(ws, r, ncols, i % 2 == 1)

    ultima = FILA_DATOS + len(universo) - 1

    notas = [
        "Fuente del catálogo: BE_SOFIPOS_202606.xlsx, hoja Universo (No., Clave CASFIM, SOFIPO, "
        "Federación, Fecha de autorización, Nivel de operación y Estatus).",
        "Las cifras financieras NO se toman del boletín: provienen del tablero mensual de la CNBV "
        "(entradas/tablero/sh_datos_csv_27_i/027_datos.csv), que se describe en las hojas de datos.",
        "Entidad Legal y Sitio Web investigados el 11 de noviembre de 2026 en el registro de la "
        "CONDUSEF/SIPRES, cuya clave registral es la misma Clave CASFIM; para Nu México se usó el "
        "oficio de SEGOB/DOF y su sitio oficial.",
        "Los sufijos de las denominaciones sociales se normalizaron a ', S.A. de C.V., S.F.P.'; "
        "no se alteró ninguna otra parte del nombre registral.",
        "Sitios activos que rechazan peticiones automáticas (HTTP 403): nu.com.mx y premo.mx.",
        "La Federación corresponde a la clasificación vigente en el boletín de junio de 2026.",
    ]
    _notas(ws, ncols, ultima + 2, notas)

    anchos = {"A": 6, "B": 13, "C": 24, "D": 20, "E": 58, "F": 48, "G": 15, "H": 7, "I": 14}
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho

    ws.freeze_panes = f"C{FILA_DATOS}"
    ws.auto_filter.ref = f"A{FILA_ENC}:{get_column_letter(ncols)}{ultima}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{FILA_ENC}:{FILA_ENC}"


# --------------------------------------------------------------------------
# Hojas: Datos Fuente 2026 / 2025
# --------------------------------------------------------------------------

def escribir_datos_fuente(wb: Workbook, filas: list[dict], anio: int, n_entidades: int) -> None:
    nombre_hoja = f"Datos Fuente {anio}"
    ws = wb.create_sheet(nombre_hoja)
    ncols = len(ENC_DATOS)

    meses = ANIOS[anio]
    rango = f"{MESES[min(meses)]}–{MESES[max(meses)]} {anio}"
    _titulo(
        ws, ncols,
        f"SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Datos Fuente {anio}",
        f"{rango} · {n_entidades} sociedades por mes · {len(filas)} filas",
        "Todas las cifras están en MILES DE PESOS. Fuente: tablero mensual de la CNBV "
        "(027_datos.csv), saldos sin consolidar (tipo_saldo 130). "
        "Los estados de resultados son acumulados del año (YTD).",
    )

    for j, encabezado in enumerate(ENC_DATOS, start=1):
        ws.cell(row=FILA_ENC, column=j, value=encabezado)
    _estilo_encabezado(ws, FILA_ENC, ncols, alto=62)

    for i, f in enumerate(filas):
        r = FILA_DATOS + i

        _celda_numero(ws, r, 1, f["no"], "0")
        _texto(ws, r, 2, f["sofipo"], negrita=True)
        _celda_numero(ws, r, 3, f["anio"], FMT_ANIO)
        _texto(ws, r, 4, f["mes"], alineacion="center")
        _texto(ws, r, 5, f["trimestre"], alineacion="center")

        # F a H: datos del tablero
        _celda_numero(ws, r, 6, f["activos"], FMT_MILES)
        _celda_numero(ws, r, 7, f["cartera_e12"], FMT_MILES)
        _celda_numero(ws, r, 8, f["cartera_e3"], FMT_MILES)

        # I: Cartera Total = G + H
        _celda_formula(ws, r, 9, f'=IFERROR(G{r}+H{r},"{ND}")', FMT_MILES)

        # J: Estimaciones
        _celda_numero(ws, r, 10, f["estimaciones"], FMT_MILES)

        # K: Cartera Neta = I - ABS(J)
        _celda_formula(ws, r, 11, f'=IFERROR(I{r}-ABS(J{r}),"{ND}")', FMT_MILES)

        # L y M: datos del tablero
        _celda_numero(ws, r, 12, f["ingresos_int"], FMT_MILES)
        _celda_numero(ws, r, 13, f["gastos_int"], FMT_MILES)

        # N: Margen Financiero = L - M
        _celda_formula(ws, r, 14, f'=IFERROR(L{r}-M{r},"{ND}")', FMT_MILES)

        # O a U: datos del tablero
        _celda_numero(ws, r, 15, f["gastos_adm"], FMT_MILES)
        _celda_numero(ws, r, 16, f["resultado_neto"], FMT_MILES)
        _celda_numero(ws, r, 17, f["captacion"], FMT_MILES)
        _celda_numero(ws, r, 18, f["captacion_prev"], FMT_MILES)
        _celda_numero(ws, r, 19, f["cartera_total_prev"], FMT_MILES)
        _celda_numero(ws, r, 20, f["capital_neto"], FMT_MILES)
        _celda_numero(ws, r, 21, f["requerimiento"], FMT_MILES)

        # V: NICAP = T / U
        _celda_formula(ws, r, 22, f'=IFERROR(T{r}/U{r},"{ND}")', FMT_PORC)

        # W a Y: insumos de ROA / ROE
        _celda_numero(ws, r, 23, f["rn_12m"], FMT_MILES)
        _celda_numero(ws, r, 24, f["activo_prom_12m"], FMT_MILES)
        _celda_numero(ws, r, 25, f["capital_prom_12m"], FMT_MILES)

        _borde_y_banda(ws, r, ncols, i % 2 == 1)

    ultima = FILA_DATOS + len(filas) - 1

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
    _notas(ws, ncols, ultima + 2, notas_2026 if anio == 2026 else notas_2025)

    anchos = {"A": 6, "B": 22, "C": 8, "D": 11, "E": 10}
    for j in range(6, ncols + 1):
        anchos[get_column_letter(j)] = 21
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho

    ws.freeze_panes = f"C{FILA_DATOS}"
    ws.auto_filter.ref = f"A{FILA_ENC}:{get_column_letter(ncols)}{ultima}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{FILA_ENC}:{FILA_ENC}"


# --------------------------------------------------------------------------
# Hojas: Indicadores 2026 / 2025
# --------------------------------------------------------------------------

def aplicar_formato_condicional(ws, nfilas: int) -> None:
    """Pinta cada indicador según los tramos definidos en FICHAS."""
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
                    c=celda, cond=condicion.format(c=celda)
                )
            ws.conditional_formatting.add(
                rango,
                FormulaRule(formula=[regla], fill=relleno, font=fuente, stopIfTrue=False),
            )


def escribir_indicadores(wb: Workbook, filas: list[dict], anio: int) -> None:
    nombre_hoja = f"Indicadores {anio}"
    ws = wb.create_sheet(nombre_hoja)
    ncols = len(ENC_INDICADORES)
    ref = f"'Datos Fuente {anio}'"

    meses = ANIOS[anio]
    rango = f"{MESES[min(meses)]}–{MESES[max(meses)]} {anio}"
    _titulo(
        ws, ncols,
        f"SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Indicadores {anio}",
        f"{rango} · {len(filas)} filas",
        "Todas las celdas son fórmulas que apuntan a la hoja Datos Fuente del mismo año, en la "
        "misma fila. NICAP y Resultado Neto se toman como parámetro; ROA y ROE se calculan con "
        "los insumos de 12 meses.",
    )

    for j, encabezado in enumerate(ENC_INDICADORES, start=1):
        ws.cell(row=FILA_ENC, column=j, value=encabezado)
    _estilo_encabezado(ws, FILA_ENC, ncols, alto=50)

    for i, f in enumerate(filas):
        r = FILA_DATOS + i

        _celda_numero(ws, r, 1, f["no"], "0")
        _texto(ws, r, 2, f["sofipo"], negrita=True)
        _celda_numero(ws, r, 3, f["anio"], FMT_ANIO)
        _texto(ws, r, 4, f["mes"], alineacion="center")
        _texto(ws, r, 5, f["trimestre"], alineacion="center")

        formulas_pct = {
            6: f"=IFERROR({ref}!V{r},\"{ND}\")",                                    # NICAP
            7: f"=IFERROR({ref}!H{r}/{ref}!I{r},\"{ND}\")",                          # IMOR
            8: f"=IFERROR(ABS({ref}!J{r})/{ref}!H{r},\"{ND}\")",                     # ICOR
            9: f"=IFERROR({ref}!W{r}/{ref}!X{r},\"{ND}\")",                          # ROA
            10: f"=IFERROR({ref}!W{r}/{ref}!Y{r},\"{ND}\")",                         # ROE
            11: f"=IFERROR({ref}!N{r}/{ref}!L{r},\"{ND}\")",                         # Margen
            12: f"=IFERROR({ref}!Q{r}/{ref}!R{r}-1,\"{ND}\")",                       # Crec. captación
            13: f"=IFERROR({ref}!I{r}/{ref}!S{r}-1,\"{ND}\")",                       # Crec. cartera
            15: f"=IFERROR({ref}!K{r}/{ref}!F{r},\"{ND}\")",                         # Crédito neto
            16: (f"=IFERROR(IF({ref}!N{r}<=0,\"n.a.\","
                 f"{ref}!O{r}/{ref}!N{r}),\"{ND}\")"),                               # % Gastos
        }
        for col, formula in formulas_pct.items():
            _celda_formula(ws, r, col, formula, FMT_PORC)

        # N (14): Resultado Neto como parámetro, en miles de pesos
        _celda_formula(ws, r, 14, f"=IFERROR({ref}!P{r},\"{ND}\")", FMT_MILES)

        _borde_y_banda(ws, r, ncols, i % 2 == 1)

    ultima = FILA_DATOS + len(filas) - 1

    aplicar_formato_condicional(ws, len(filas))

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

    anchos = {"A": 6, "B": 22, "C": 8, "D": 11, "E": 10}
    for j in range(6, ncols + 1):
        anchos[get_column_letter(j)] = 20
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho

    ws.freeze_panes = f"C{FILA_DATOS}"
    ws.auto_filter.ref = f"A{FILA_ENC}:{get_column_letter(ncols)}{ultima}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{FILA_ENC}:{FILA_ENC}"


# --------------------------------------------------------------------------
# Hoja: Diccionario
# --------------------------------------------------------------------------

def escribir_diccionario(wb: Workbook) -> None:
    ws = wb.create_sheet("Diccionario")
    ncols = len(ENC_DICCIONARIO)

    _titulo(
        ws, ncols,
        "SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Diccionario de indicadores",
        "Qué mide cada indicador, cómo se calcula y a qué grupo de razón financiera pertenece",
        "Cada renglón corresponde a una columna de las hojas Indicadores 2026 e Indicadores 2025. "
        "Las fórmulas se expresan en función de las columnas de Datos Fuente.",
    )

    for j, encabezado in enumerate(ENC_DICCIONARIO, start=1):
        ws.cell(row=FILA_ENC, column=j, value=encabezado)
    _estilo_encabezado(ws, FILA_ENC, ncols, alto=32)

    for i, ficha in enumerate(FICHAS):
        r = FILA_DATOS + i
        _texto(ws, r, 1, ficha.indicador, negrita=True)
        _texto_envuelto(ws, r, 2, ficha.formula, alto=34)
        _texto(ws, r, 3, ficha.grupo, alineacion="center")
        _texto_envuelto(ws, r, 4, ficha.descripcion, alto=40)
        _borde_y_banda(ws, r, ncols, i % 2 == 1)

    ultima = FILA_DATOS + len(FICHAS) - 1

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

    for col, ancho in {"A": 30, "B": 62, "C": 28, "D": 88}.items():
        ws.column_dimensions[col].width = ancho

    ws.freeze_panes = f"A{FILA_DATOS}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{FILA_ENC}:{FILA_ENC}"


# --------------------------------------------------------------------------
# Hoja: Semáforos
# --------------------------------------------------------------------------

def escribir_semaforos(wb: Workbook) -> None:
    ws = wb.create_sheet("Semáforos")
    ncols = len(ENC_SEMAFOROS)

    _titulo(
        ws, ncols,
        "SOCIEDADES FINANCIERAS POPULARES (SOFIPO) · Semáforos",
        "Umbrales de referencia por indicador · se aplican como formato condicional en Indicadores",
        "Verde = dentro del rango objetivo · Amarillo = vigilancia · "
        "Rojo = fuera de rango o n.a. (margen ≤ 0). Las condiciones se evalúan sobre "
        "fracciones (141% = 1.41) y solo si la celda es numérica.",
    )

    for j, encabezado in enumerate(ENC_SEMAFOROS, start=1):
        ws.cell(row=FILA_ENC, column=j, value=encabezado)
    _estilo_encabezado(ws, FILA_ENC, ncols, alto=24)

    fila = FILA_DATOS
    for ficha in FICHAS_SEMAFORO:
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
        titulo = ws.cell(
            row=fila, column=1,
            value=f"{ficha.indicador}  ·  {ficha.grupo}  ·  columna {ficha.columna} de Indicadores",
        )
        titulo.font = Font(bold=True, color="FFFFFF", size=10)
        titulo.fill = PatternFill("solid", fgColor="2F5597")
        titulo.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[fila].height = 20
        fila += 1

        etiqueta = ws.cell(row=fila, column=1, value="Umbral")
        etiqueta.font = Font(bold=True, size=10)
        etiqueta.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        etiqueta.border = Border(left=BORDE, right=BORDE, top=BORDE, bottom=BORDE)

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
                celda.alignment = Alignment(horizontal="center", vertical="center")
            celda.alignment = Alignment(horizontal="center", vertical="center")
            celda.border = Border(left=BORDE, right=BORDE, top=BORDE, bottom=BORDE)
        ws.row_dimensions[fila].height = 20
        fila += 1

        celda = ficha.columna + str(FILA_DATOS)
        detalles = []
        for texto, condicion, _ in ficha.tramos:
            detalles.append(f"{texto} → {condicion.format(c=celda)}")
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
        nota = ws.cell(row=fila, column=1, value="Excel (fracción): " + " · ".join(detalles))
        nota.font = Font(size=8, italic=True, color="595959")
        nota.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[fila].height = 16
        fila += 2

    leyenda_fila = fila
    ws.merge_cells(start_row=leyenda_fila, start_column=1, end_row=leyenda_fila, end_column=ncols)
    leyenda = ws.cell(row=leyenda_fila, column=1, value="Leyenda")
    leyenda.font = Font(bold=True, size=10, color=AZUL)
    leyenda.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    fila += 1
    for texto, relleno, fuente in (
        ("Verde: comportamiento dentro del rango objetivo", VERDE_FILL, VERDE_FONT),
        ("Amarillo: vigilancia, el indicador se acerca al límite", AMARILLO_FILL, AMARILLO_FONT),
        ("Rojo: fuera de rango o n.a. (margen ≤ 0), requiere seguimiento", ROJO_FILL, ROJO_FONT),
    ):
        for col in range(1, ncols + 1):
            ws.cell(row=fila, column=col).fill = relleno
            ws.cell(row=fila, column=col).border = Border(
                left=BORDE, right=BORDE, top=BORDE, bottom=BORDE
            )
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
        celda = ws.cell(row=fila, column=1, value=texto)
        celda.font = fuente
        celda.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[fila].height = 18
        fila += 1

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

    for col, ancho in {"A": 40, "B": 24, "C": 24, "D": 24}.items():
        ws.column_dimensions[col].width = ancho

    ws.freeze_panes = f"A{FILA_DATOS}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# --------------------------------------------------------------------------
# Ensamblado del libro
# --------------------------------------------------------------------------

def escribir_xlsx(universo: list[dict], nombres: dict[str, str],
                  filas: dict[int, list[dict]], ruta: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    escribir_catalogo(wb, universo, nombres)
    escribir_datos_fuente(wb, filas[2026], 2026, len(universo))
    escribir_datos_fuente(wb, filas[2025], 2025, len(universo))
    escribir_indicadores(wb, filas[2026], 2026)
    escribir_indicadores(wb, filas[2025], 2025)
    escribir_diccionario(wb)
    escribir_semaforos(wb)

    wb.properties.title = ("SOFIPOS - Catálogo, datos fuente, indicadores, diccionario "
                           "y semáforos 2025-2026")
    wb.properties.subject = "Cartera, captación, resultados, capitalización e indicadores mensuales"
    wb.properties.creator = "Generador automático (tablero CNBV + Universo BE)"

    wb.calculation.fullCalcOnLoad = True

    ruta.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile("wb", suffix=".xlsx", dir=ruta.parent, delete=False) as tmp:
        wb.save(tmp.name)
        temporal = tmp.name
    os.replace(temporal, ruta)


# --------------------------------------------------------------------------
# Validación del libro generado
# --------------------------------------------------------------------------

def validar_libro(ruta: Path, universo: list[dict], filas: dict[int, list[dict]]) -> bool:
    from openpyxl import load_workbook

    ok = True
    wb = load_workbook(ruta)
    print("\n=== Validación del libro generado ===")
    print(f"  hojas: {wb.sheetnames}")

    esperadas = ["Catálogo SOFIPOS", "Datos Fuente 2026", "Datos Fuente 2025",
                 "Indicadores 2026", "Indicadores 2025", "Diccionario", "Semáforos"]
    if wb.sheetnames != esperadas:
        print(f"  ERROR: se esperaban {esperadas}.")
        ok = False

    cat = wb["Catálogo SOFIPOS"]
    dic = wb["Diccionario"]
    sem = wb["Semáforos"]

    filas_cat = sum(1 for r in range(FILA_DATOS, FILA_DATOS + len(universo))
                    if cat.cell(row=r, column=2).value)
    print(f"  Catálogo SOFIPOS: {filas_cat} filas (esperado {len(universo)})")
    if filas_cat != len(universo):
        ok = False

    formatos_fecha = {
        cat.cell(row=r, column=7).number_format
        for r in range(FILA_DATOS, FILA_DATOS + len(universo))
    }
    print(f"  Catálogo SOFIPOS: formato de fecha {formatos_fecha} (esperado {{'{FMT_FECHA}'}})")
    if formatos_fecha != {FMT_FECHA}:
        ok = False

    for anio, lista in filas.items():
        datos = wb[f"Datos Fuente {anio}"]
        ind = wb[f"Indicadores {anio}"]

        filas_datos = sum(1 for r in range(FILA_DATOS, FILA_DATOS + len(lista))
                          if datos.cell(row=r, column=1).value is not None)
        filas_ind = sum(1 for r in range(FILA_DATOS, FILA_DATOS + len(lista))
                        if ind.cell(row=r, column=1).value is not None)
        print(f"  Datos Fuente {anio}: {filas_datos} filas · Indicadores {anio}: {filas_ind} "
              f"(esperado {len(lista)})")
        if filas_datos != len(lista) or filas_ind != len(lista):
            ok = False

        desalineadas = [
            r for r in range(FILA_DATOS, FILA_DATOS + len(lista))
            if datos.cell(row=r, column=2).value != ind.cell(row=r, column=2).value
            or datos.cell(row=r, column=4).value != ind.cell(row=r, column=4).value
            or datos.cell(row=r, column=5).value != ind.cell(row=r, column=5).value
        ]
        print(f"    filas desalineadas Datos Fuente vs Indicadores: {len(desalineadas)}")
        if desalineadas:
            ok = False

        # 35 entidades por mes
        meses = {}
        for r in range(FILA_DATOS, FILA_DATOS + len(lista)):
            meses[datos.cell(row=r, column=4).value] = meses.get(datos.cell(row=r, column=4).value, 0) + 1
        malos = {m: n for m, n in meses.items() if n != len(universo)}
        print(f"    meses: {len(meses)} · {'35 entidades por mes' if not malos else malos}")
        if malos:
            ok = False

        # Fórmulas visibles
        r0 = FILA_DATOS
        for col, esperado in [
            (9, f'=IFERROR(G{r0}+H{r0},"{ND}")'),
            (11, f'=IFERROR(I{r0}-ABS(J{r0}),"{ND}")'),
            (14, f'=IFERROR(L{r0}-M{r0},"{ND}")'),
            (22, f'=IFERROR(T{r0}/U{r0},"{ND}")'),
        ]:
            valor = datos.cell(row=r0, column=col).value
            if valor != esperado:
                print(f"      ERROR: Datos Fuente {anio} {get_column_letter(col)}{r0} = {valor!r}")
                ok = False
        for col in range(6, 17):
            valor = ind.cell(row=r0, column=col).value
            if not (isinstance(valor, str) and valor.startswith("=")):
                print(f"      ERROR: Indicadores {anio} {get_column_letter(col)}{r0} no es fórmula: {valor!r}")
                ok = False

        formula_gastos = ind.cell(row=r0, column=16).value
        if '"n.a."' not in str(formula_gastos):
            print(f"      ERROR: Indicadores {anio} P{r0} no maneja n.a.: {formula_gastos!r}")
            ok = False
        if ind.cell(row=FILA_ENC, column=14).value != "Resultado Neto\n(miles de pesos)":
            print(f"      ERROR: Indicadores {anio} no tiene el orden de columnas esperado.")
            ok = False

    filas_dic = sum(1 for r in range(FILA_DATOS, FILA_DATOS + len(FICHAS))
                    if dic.cell(row=r, column=1).value)
    print(f"  Diccionario: {filas_dic} indicadores (esperado {len(FICHAS)})")
    if filas_dic != len(FICHAS):
        ok = False

    bloques = sum(
        1 for r in range(FILA_DATOS, sem.max_row + 1)
        if sem.cell(row=r, column=1).value == "Umbral"
    )
    print(f"  Semáforos: {bloques} bloques de indicador (esperado {len(FICHAS_SEMAFORO)})")
    if bloques != len(FICHAS_SEMAFORO):
        ok = False

    for anio in filas:
        ind = wb[f"Indicadores {anio}"]
        reglas = sum(len(lista) for lista in ind.conditional_formatting._cf_rules.values())
        print(f"  formato condicional {anio}: {reglas} reglas (esperado {TOTAL_REGLAS})")
        if reglas != TOTAL_REGLAS:
            ok = False

    return ok


# --------------------------------------------------------------------------
# Principal
# --------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Genera el Excel mensual de las 35 SOFIPOS (2025 completo y 2026 enero-agosto)."
    )
    parser.add_argument("--entradas", type=Path, default=DIR_ENTRADAS,
                        help="Carpeta con el BE del Universo y la carpeta tablero/")
    parser.add_argument("--salidas", type=Path, default=DIR_SALIDAS,
                        help="Carpeta donde se escribe el .xlsx")
    parser.add_argument("--out", type=Path, default=None,
                        help="Ruta exacta del .xlsx (por defecto salidas/SOFIPOS_2026_Datos.xlsx)")
    parser.add_argument("--tablero", type=Path, default=None,
                        help="Ruta exacta del CSV 027_datos.csv")
    args = parser.parse_args()

    dir_entradas = args.entradas.expanduser().resolve()
    if not dir_entradas.is_dir():
        raise SystemExit(f"No existe la carpeta de entradas: {dir_entradas}")

    ruta_csv = (args.tablero.expanduser().resolve() if args.tablero
                else dir_entradas / CSV_TABLERO)
    ruta_universo = dir_entradas / BE_UNIVERSO

    faltantes = [str(p) for p in (ruta_csv, ruta_universo) if not p.exists()]
    if faltantes:
        disponibles = sorted(x.name for x in dir_entradas.iterdir() if x.is_file())
        raise SystemExit(
            "Faltan archivos:\n  - " + "\n  - ".join(faltantes)
            + "\nContenido actual de entradas/:\n  - " + ("\n  - ".join(disponibles) or "(vacía)")
        )

    salida = args.out or (args.salidas.expanduser().resolve() / SALIDA_DEFECTO.name)

    universo = leer_universo(ruta_universo)
    be_principal = leer_be_sofipos(ruta_universo)
    nombres = {r["clave"]: r["nombre"] for r in be_principal.registros}

    print(f"Entradas: {dir_entradas}")
    print(f"Universo: {len(universo)} sociedades · {ruta_universo.name}")

    sin_catalogo = [u["clave"] for u in universo if u["clave"] not in CATALOGO]
    if sin_catalogo:
        raise SystemExit(f"Faltan datos de catálogo para: {sin_catalogo}")

    # Fechas necesarias: todos los meses del libro + el mes anterior de enero
    fechas: set[str] = set()
    for anio, meses in ANIOS.items():
        for mes in meses:
            f = fecha_mes(anio, mes)
            fechas.add(f)
            fechas.add(fecha_anterior(f))

    print(f"\nLeyendo tablero: {ruta_csv.name} ...")
    tablero = leer_tablero(ruta_csv, fechas)

    filas = {anio: preparar_filas(tablero, universo, nombres, anio) for anio in ANIOS}
    print(f"  filas preparadas: " + " · ".join(f"{a}: {len(l)}" for a, l in filas.items()))

    be_por_archivo = {}
    for nombre in BE_VALIDACION:
        ruta = dir_entradas / nombre
        if ruta.exists():
            be_por_archivo[nombre] = leer_be_sofipos(ruta)

    todo_ok = validar_datos(tablero, filas, universo, be_por_archivo)

    escribir_xlsx(universo, nombres, filas, salida)
    todo_ok &= validar_libro(salida, universo, filas)

    print(f"\nArchivo generado: {salida}")
    print("Validación general:", "OK" if todo_ok else "CON OBSERVACIONES")
    return 0 if todo_ok else 2


if __name__ == "__main__":
    sys.exit(main())

"""Utilidades para leer hojas de cálculo XLSX directamente desde su contenido XML."""

from __future__ import annotations

import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from .config import COL_TOTAL_SECTOR
from .utilidades import _a_numero, _norm

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _col_letra(ref: str) -> str:
    """Extrae las letras de columna de una referencia de celda.

    Por ejemplo, convierte ``"BC12"`` en ``"BC"``.

    Args:
        ref: Referencia de celda de Excel.

    Returns:
        Letras de la columna, o una cadena vacía si no hay coincidencia.
    """
    m = re.match(r"[A-Z]+", ref or "")
    return m.group(0) if m else ""


def _leer_shared_strings(z: zipfile.ZipFile) -> list[str]:
    """Lee la tabla de cadenas compartidas del archivo XLSX.

    Args:
        z: Archivo XLSX abierto como archivo ZIP.

    Returns:
        Cadenas compartidas en el orden en que aparecen. Devuelve una lista
        vacía si el archivo no contiene ``sharedStrings.xml``.
    """
    try:
        raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    return [
        "".join(t.text or "" for t in si.iter(NS + "t"))
        for si in raiz.findall(NS + "si")
    ]


def _rutas_de_hojas(z: zipfile.ZipFile) -> dict[str, str]:
    """Obtiene la ruta interna de cada hoja del archivo XLSX.

    Args:
        z: Archivo XLSX abierto como archivo ZIP.

    Returns:
        Diccionario que relaciona el nombre de cada hoja con su ruta interna.
    """
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    mapa = {
        r.get("Id"): r.get("Target")
        for r in rels.findall(NS_REL + "Relationship")
    }

    rutas = {}
    for hoja in wb.find(NS + "sheets"):
        destino = mapa[hoja.get(RNS + "id")]
        if destino.startswith("/"):
            ruta = destino.lstrip("/")
        else:
            ruta = posixpath.normpath(posixpath.join("xl", destino))
        rutas[hoja.get("name")] = ruta

    return rutas


def _leer_hoja(
    z: zipfile.ZipFile,
    ruta: str,
    shared: list[str],
) -> dict[int, dict[str, object]]:
    """Lee las celdas de una hoja y las organiza por fila y columna.

    Args:
        z: Archivo XLSX abierto como archivo ZIP.
        ruta: Ruta interna de la hoja dentro del archivo.
        shared: Lista de cadenas compartidas del libro.

    Returns:
        Diccionario con el número de fila como clave y, para cada fila,
        un diccionario de letras de columna y valores.
    """
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
    """Abre un XLSX y lee la hoja indicada por nombre.

    Args:
        ruta: Ruta del archivo XLSX.
        nombre: Nombre de la hoja que se desea leer.

    Returns:
        Celdas de la hoja organizadas por fila y columna.

    Raises:
        SystemExit: Si el libro no contiene una hoja con el nombre indicado.
    """
    with zipfile.ZipFile(ruta) as z:
        shared = _leer_shared_strings(z)
        rutas = _rutas_de_hojas(z)

        if nombre not in rutas:
            raise SystemExit(
                f"No se encontró la hoja '{nombre}' en {ruta.name}. "
                f"Hojas: {sorted(rutas)}"
            )

        return _leer_hoja(z, rutas[nombre], shared)


def _fila_encabezado(hoja: dict[int, dict[str, object]], texto: str) -> int:
    """Busca en la columna D la fila que comienza con el encabezado indicado.

    La comparación normaliza el texto para ignorar diferencias de formato.

    Args:
        hoja: Celdas de la hoja organizadas por fila y columna.
        texto: Texto esperado al inicio del encabezado.

    Returns:
        Número de la fila donde se encontró el encabezado.

    Raises:
        SystemExit: Si no se encuentra el encabezado en la columna D.
    """
    objetivo = _norm(texto)

    for num in sorted(hoja):
        if _norm(hoja[num].get("D", "")).startswith(objetivo):
            return num

    raise SystemExit(
        f"No se localizó el encabezado '{texto}' en la columna D."
    )


def _columnas_sofipo(codigos: dict[str, object]) -> list[str]:
    """Identifica las columnas cuyo valor corresponde a un código de SOFIPO.

    Excluye la columna del total del sector y la columna D. Los códigos deben
    contener entre cuatro y seis dígitos.

    Args:
        codigos: Diccionario con letras de columna y valores de sus celdas.

    Returns:
        Letras de las columnas con códigos válidos, ordenadas por posición.
    """
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

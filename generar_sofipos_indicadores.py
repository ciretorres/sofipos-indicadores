#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera SOFIPOS_Indicadores_Datos.xlsx con la información mensual de las SOFIPOS:
2024 y 2025 completos (enero-diciembre) y 2026 de enero a septiembre (corte
30 de septiembre de 2026).

Fuente de las cifras
--------------------
Tablero de información financiera de la CNBV, descarga "sh_datos_csv_27_i"
(serie mensual 201601-202609, saldos sin consolidar, tipo_saldo = 130):

  entradas/tablero/sh_datos_csv_27_i/027_datos.csv
  entradas/tablero/sh_datos_csv_27_i/catalogo_conceptos_027.xlsx
  entradas/tablero/sh_datos_csv_27_i/catalogo_instituciones_027.xlsx

El único dato que se sigue tomando del boletín es la hoja "Universo"
(Federación, Nivel de operación, Fecha de autorización y Estatus):

  entradas/BE_SOFIPOS_202609.xlsx, hoja Universo

El CSV viene en pesos; el libro se publica en MILES de pesos.

Hojas del libro
---------------
Estructura dinámica: 3 hojas fijas (Catálogo, Diccionario y Semáforos) y 2
por cada año de ANIOS (Datos Fuente e Indicadores). Con 2024-2026 son 9 hojas.
  1. Catálogo SOFIPOS   -> 33 sociedades en operación + apartado de bajas
  2. Datos Fuente 2026   -> enero-septiembre 2026, 35 entidades por mes (315 filas)
  3. Datos Fuente 2025   -> enero-diciembre 2025, 35 entidades por mes (420 filas)
  4. Datos Fuente 2024   -> enero-diciembre 2024, 35 entidades por mes (420 filas)
  5. Indicadores 2026    -> 315 filas, fórmulas contra Datos Fuente 2026
  6. Indicadores 2025    -> 420 filas, fórmulas contra Datos Fuente 2025
  7. Indicadores 2024    -> 420 filas, fórmulas contra Datos Fuente 2024
  8. Diccionario         -> 11 indicadores, fórmula, grupo y descripción
  9. Semáforos           -> 10 bloques y 31 reglas de formato condicional

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

Uso (ambas formas son equivalentes y se corren desde la raíz del repo):
    .venv/bin/python generar_sofipos_indicadores.py
    .venv/bin/python -m sofipos
    .venv/bin/python generar_sofipos_indicadores.py --entradas otra --salidas otra
    .venv/bin/python generar_sofipos_indicadores.py --out otra_ruta.xlsx
    .venv/bin/python generar_sofipos_indicadores.py --anios 2026

El código vive en el paquete sofipos/; este archivo solo lo invoca.
"""

from __future__ import annotations

import sys

from sofipos.main import main

if __name__ == "__main__":
    sys.exit(main())

# SOFIPOS · Generador de datos e indicadores

Programa que arma un solo libro de Excel con las 35 sociedades financieras
populares (SOFIPO) que lo componen —33 activas + 2 bajas con histórico—:
catálogo, cifras mensuales por año, indicadores, diccionario y semáforos.

Todas las cifras financieras provienen del **tablero mensual de la CNBV**. Del
boletín electrónico (BE) se toma únicamente la hoja `Universo`, que no existe
en el tablero.

## Estructura

```
sofipos-indicadores/
├── entradas/                                 datos de la CNBV (no versionado)
│   ├── tablero/
│   │   └── sh_datos_csv_27_i/
│   │       ├── 027_datos.csv                 fuente mensual (en pesos)
│   │       ├── catalogo_conceptos_027.xlsx   1,502 conceptos
│   │       └── catalogo_instituciones_027.xlsx
│   └── BE_SOFIPOS_AAAAMM.xlsx                hoja Universo 
├── salidas/                                  libro generado (no versionado)
│   └── SOFIPOS_Indicadores_Datos.xlsx        3 hojas fijas + 2 por año
├── generar_sofipos_indicadores.py            punto de entrada (wrapper)
├── pyproject.toml                            metadatos, dependencias, CLI y pytest
├── sofipos/                                  todo el código del generador
│   ├── __main__.py                           permite `python -m sofipos`
│   ├── main.py                               CLI (argparse) y orquestación
│   ├── config.py                             rutas, ANIOS, MESES, CONCEPTOS, ND
│   ├── catalogo.py                           CATALOGO: 35 entidades
│   ├── utilidades.py                         fechas, números, normalización
│   ├── xlsx_lector.py                        lee .xlsx con zipfile + ElementTree
│   ├── be.py                                 boletín BE: Universo y validación
│   ├── tablero.py                            lee el CSV mensual (027_datos.csv)
│   ├── filas.py                              arma las filas de cada año
│   ├── validacion_datos.py                   las 6 validaciones de datos
│   ├── estilos.py                            paleta, formatos y helpers de celda
│   ├── fichas.py                             FICHAS: diccionario y semáforos
│   ├── hojas/                                escritura de las hojas
│   │   ├── __init__.py                       escribir_xlsx (orden de hojas)
│   │   ├── catalogo.py                       Catálogo SOFIPOS
│   │   ├── datos_fuente.py                   Datos Fuente {año}
│   │   ├── indicadores.py                    Indicadores {año} + CF
│   │   ├── diccionario.py                    Diccionario
│   │   └── semaforos.py                      Semáforos
│   └── validacion_libro.py                   valida el .xlsx generado
├── scripts/
│   └── comparar.py                           compara dos libros y dos stdout
├── tests/
│   ├── fixtures/
│   │   └── tablero_mini.csv                  CSV mini versionado (no usa entradas/)
│   ├── test_config.py
│   ├── test_cli.py
│   ├── test_filas.py
│   ├── test_semaforos.py
│   ├── test_tablero.py
│   ├── test_utilidades.py
│   └── test_validacion_datos.py
└── README.md
```

Cada archivo tiene una sola responsabilidad y solo importa de capas
anteriores (no hay imports circulares):

```
config ← utilidades ← xlsx_lector ← be / tablero ← filas ← validacion_datos
   ↑                                                             ↑
estilos ← fichas ← hojas/* ← validacion_libro ← main ← generar_sofipos_indicadores.py
```

`entradas/` y `salidas/` están en `.gitignore` porque pesan decenas de MB y
los datos los publica la CNBV. Para reconstruir la carpeta de entrada, sigue
las instrucciones de [`entradas/README.md`](entradas/README.md).

## Requisitos

- Python 3.11 o superior. En Windows usa `python`; en macOS y Linux, `python3`.
- [`openpyxl`](https://openpyxl.readthedocs.io/) 3.1+ (las dependencias viven
  en `pyproject.toml`)

## Instalación

```bash
git clone https://github.com/ciretorres/sofipos-indicadores.git
cd sofipos-indicadores
```

### Crear el entorno virtual

**Windows** (PowerShell o CMD):

```powershell
python -m venv .venv
```

**macOS / Linux**:

```bash
python3 -m venv .venv
```

### Instalar dependencias

**Windows**:

```powershell
.venv\Scripts\python.exe -m pip install -e ".[dev]"
```

**macOS / Linux**:

```bash
.venv/bin/pip install -e ".[dev]"
```

### MCP de Excel (opcional)

Para que OpenCode o Claude Code lean los `.xlsx` con el MCP de Excel se
instala `uv` (comando en `AGENTS.md` § MCP de Excel); el generador no lo
necesita.

## Pruebas

```bash
.venv/bin/python -m pytest
```

Los tests (en `tests/`) usan fixtures CSV versionados y no requieren las
carpetas `entradas/` ni `salidas/`.

## Uso

Descarga primero los datos de entrada (ver
[`entradas/README.md`](entradas/README.md)) y después ejecuta:

**Windows**:

```powershell
.venv\Scripts\python.exe generar_sofipos_indicadores.py
```

**macOS / Linux**:

```bash
.venv/bin/python generar_sofipos_indicadores.py
```

Sale con código 0 cuando todas las validaciones pasan. Hay dos formas de
invocarlo (hacen exactamente lo mismo); ambas se ejecutan desde la raíz del
repositorio:

**Windows**:

```powershell
.venv\Scripts\python.exe -m sofipos
```

**macOS / Linux**:

```bash
.venv/bin/python -m sofipos
```

Opciones (iguales en las dos formas):

**Windows**:

```powershell
.venv\Scripts\python.exe generar_sofipos_indicadores.py --entradas otra_carpeta
.venv\Scripts\python.exe generar_sofipos_indicadores.py --salidas otra_carpeta
.venv\Scripts\python.exe generar_sofipos_indicadores.py --out ruta/exacta.xlsx
.venv\Scripts\python.exe generar_sofipos_indicadores.py --tablero ruta/027_datos.csv
.venv\Scripts\python.exe generar_sofipos_indicadores.py --anios 2026
```

**macOS / Linux**:

```bash
.venv/bin/python generar_sofipos_indicadores.py --entradas otra_carpeta
.venv/bin/python generar_sofipos_indicadores.py --salidas otra_carpeta
.venv/bin/python generar_sofipos_indicadores.py --out ruta/exacta.xlsx
.venv/bin/python generar_sofipos_indicadores.py --tablero ruta/027_datos.csv
.venv/bin/python generar_sofipos_indicadores.py --anios 2026
```

`--anios` acepta un año (`2026`) o un rango (`2024-2026`); sin él se generan
todos los años configurados en `ANIOS`. Las hojas de un año fuera de `ANIOS`
terminan con un mensaje de error.

## Hojas del libro

La estructura es dinámica: **3 hojas fijas** (`Catálogo SOFIPOS`, `Diccionario`
y `Semáforos`) y **2 por cada año generado** (`Datos Fuente AAAA` e
`Indicadores AAAA`). Con 2024–2026 el libro trae 9 hojas; con `--anios 2026`,
5.

| Hoja | Contenido |
|---|---|
| `Catálogo SOFIPOS` | 33 sociedades en operación: No., Clave CASFIM, SOFIPO, Federación, Entidad Legal, Sitio Web, Fecha de Autorización, Nivel, Estatus. Al final, un apartado con las sociedades transformadas o revocadas (clave, nombre, motivo y fecha) |
| `Datos Fuente AAAA` | Activos Totales, Cartera Actual, Cartera Vencida, Cartera Total, Estimaciones Preventivas, Cartera Neta, Ingresos por Intereses, Gastos por Intereses, Margen Financiero, Gastos de Administración, Resultado Neto, Captación Tradicional, Captación del mes anterior, Captación total del mes anterior, Capital Neto, Requerimiento de Capital, NICAP, Resultado Neto 12M, Activo Promedio 12M, Capital Contable Promedio 12M |
| `Indicadores AAAA` | Todas por fórmula contra `Datos Fuente AAAA`: NICAP, IMOR, ICOR, ROA, ROE, Margen Financiero, Crecimiento de Captación, Crecimiento de Cartera, Resultado Neto, Crédito Neto, % Gastos de Administración |
| `Diccionario` | 11 indicadores: fórmula, grupo de razón financiera y descripción |
| `Semáforos` | 10 bloques y 31 reglas de formato condicional |

## Fuente: tablero de la CNBV

`entradas/tablero/sh_datos_csv_27_i/027_datos.csv` es una serie larga
(`concepto, institucion, fecha, saldo_se, tipo_saldo`):

- **905 conceptos**, de **201601 a 202609**, frecuencia mensual.
- Importes en **pesos**.
- Solo se usan saldos **sin consolidar** (`tipo_saldo = 130`).
- La fila `institucion = 27` es el **total del sector**: no es una SOFIPO y se
  excluye de las sumas.

Conceptos utilizados:

| Dato | Concepto CNBV |
|---|---|
| Activos totales | `100000000000` |
| Cartera E1 + E2 | `272724` (respaldo: `100800104001` + `100800104002`) |
| Cartera E3 | `100800104003` (si falta: `130800103001` − E1+E2) |
| Cartera de crédito total | `130800103001` (solo para derivar y validar) |
| Estimaciones preventivas | `100800103003` |
| Ingresos por intereses | `500200101001` |
| Gastos por intereses | `600400101002` |
| Gastos de administración y promoción | `601800301009` |
| Resultado neto YTD | `330000000006` |
| Captación tradicional | `200200001001` |
| Capital Neto | `911300000000` |
| Requerimiento de capital | `910700000000` |
| Resultado neto 12M (ROA/ROE) | `27200043` |
| Activo promedio 12M | `27200044` |
| Capital contable promedio 12M | `27200051` |
| NICAP publicado (validación) | `27200093` |
| IMOR publicado (validación) | `27200005` |
| ICOR publicado (validación) | `27200013` |
| ROA / ROE publicados (validación) | `27200001` / `27200002` |

Del BE solo se lee `BE_SOFIPOS_AAAAMM.xlsx`, hoja `Universo`: Federación,
Nivel de operación, Fecha de autorización y Estatus. El cruce es siempre por
**Clave CASFIM**.

## Unidades

**Todo el libro está en miles de pesos.** El tablero publica en pesos; el
generador divide cada importe entre 1,000 antes de escribirlo. Las columnas
monetarias usan el formato `#,##0.00`. Las notas al pie de cada hoja de datos
lo indican.

## Método

- **Periodicidad mensual.** 2024 y 2025 completos (enero–diciembre) y 2026 de
  enero a septiembre.
- **Fecha de autorización.** La celda guarda una fecha real con formato
  `dd-mmm-yyyy`, de modo que ordenar y filtrar siguen funcionando. El nombre
  del mes se muestra según el idioma de Excel (en español, `18-dic-2008`).
- **Periodo anterior = mes inmediato anterior.** Enero 2025 se compara contra
  diciembre 2024; enero 2026 contra diciembre 2025. Las columnas de captación y
  cartera del mes anterior van como **dato**, no como fórmula, para que cada
  hoja sea autocontenida.
- **35 entidades en cada mes** (33 en operación + 2 transformadas o revocadas
  con histórico). Si una entidad aún no operaba, dejó de operar o no reportó el
  mes, sus celdas dicen `n.d.`; no se imputan ceros.
- **Estados de resultados en YTD.** Ingresos, gastos, gastos de administración
  y resultado neto son acumulados del año, igual que los publica el tablero.
- **Cartera Total = G + H**, **Cartera Neta = I − ABS(J)** y
  **Margen Financiero = L − M** son fórmulas visibles. **NICAP = T / U**
  también.
- **ROA y ROE se calculan, no se copian:**
  - `ROA = Resultado Neto 12M / Activo promedio 12M` (W / X)
  - `ROE = Resultado Neto 12M / Capital contable promedio 12M` (W / Y)
  Reproducen exactamente los indicadores `27200001` / `27200002`.
- **NICAP** se calcula como `Capital Neto / Requerimiento` y coincide con
  `27200093`.
- Las estimaciones preventivas conservan el signo reportado por la CNBV
  (registro de activo con saldo acreedor).

### Columnas de `Datos Fuente AAAA`

| Col | Campo | Tipo |
|---|---|---|
| A–E | No., SOFIPO, Año, Mes, Trimestre | dato |
| F | Activos Totales | dato |
| G | Cartera de crédito etapa 1 + etapa 2 | dato |
| H | Cartera de crédito etapa 3 | dato |
| I | Cartera Total = G + H | fórmula |
| J | Estimaciones preventivas | dato |
| K | Cartera Neta = I − ABS(J) | fórmula |
| L | Ingresos por intereses (YTD) | dato |
| M | Gastos por intereses (YTD) | dato |
| N | Margen Financiero = L − M | fórmula |
| O | Gastos de administración y promoción (YTD) | dato |
| P | Resultado Neto (YTD) | dato |
| Q | Captación Tradicional | dato |
| R | Captación del mes anterior | dato |
| S | Cartera Total del mes anterior | dato |
| T | Capital Neto | dato |
| U | Requerimiento de capital | dato |
| V | NICAP = T / U | fórmula |
| W | Resultado Neto 12M | dato |
| X | Activo promedio 12M | dato |
| Y | Capital contable promedio 12M | dato |

### Columnas de `Indicadores AAAA`

| Col | Indicador | Fórmula |
|---|---|---|
| F | 🏦 NICAP | `Datos Fuente!V` (parámetro) |
| G | 📉 IMOR | `H / I` |
| H | 📉 ICOR | `ABS(J) / H` |
| I | 💰 ROA | `W / X` |
| J | 💰 ROE | `W / Y` |
| K | 💰 Margen Financiero | `N / L` |
| L | 📈 Crecimiento de Captación | `Q / R − 1` |
| M | 📈 Crecimiento de Cartera | `I / S − 1` |
| N | 💰 Resultado Neto | `Datos Fuente!P` (parámetro, miles) |
| O | 📊 Crédito Neto | `K / F` |
| P | ⚙️ % Gastos de Administración | `O / N`, con `n.a.` si el margen es ≤ 0 |

Todas son fórmulas que apuntan a la hoja de datos del mismo año, en la misma
fila.

## Semáforos

| Indicador | Verde 🟢 | Amarillo 🟡 | Rojo 🔴 |
|---|---|---|---|
| NICAP | ≥ 141% | 130% – < 141% | < 130% |
| IMOR | < 5% | 5% – 10% | > 10% |
| ICOR | > 100% | 80% – 100% | < 80% |
| ROA | > 2% | 0% – 2% | < 0% |
| ROE | > 15% | 5% – 15% | < 5% |
| Margen Financiero | > 70% | 50% – 70% | < 50% |
| Crecimiento de Captación | > 10% | 0% – 10% | < 0% |
| Crecimiento de Cartera | > 10% | 0% – 10% | < 0% |
| Crédito Neto | 70% – 85% | 60% – < 70% y > 85% – 90% | < 60% o > 90% |
| % Gastos de Administración | < 60% | 60% – 80% | > 80% o `n.a.` |

- **Resultado Neto no tiene semáforo.**
- `% Gastos de Administración` muestra `n.a.` cuando el margen financiero es
  cero o negativo; esa celda se marca en rojo porque la razón no aplica.
- Paleta con contraste **AAA de WCAG 2.2** (≥ 7:1 entre texto y fondo):
  verde `D1E7DD` / `0F5132` (7.21:1), amarillo `FFF3CD` / `664D03`
  (7.21:1) y rojo `F8D7DA` / `842029` (7.08:1).
- Las condiciones se evalúan sobre fracciones (141% = 1.41) y solo si la celda
  es numérica, para que los `n.d.` no se pinten. La regla de `n.a.` es la
  excepción: aplica a texto.
- Las reglas son exclusivas: 10 indicadores × 3 + la regla de `n.a.` = **31
  reglas por hoja de indicadores**.

## Validaciones que corre el script

1. **35 entidades en cada mes** de 2024, 2025 y 2026 (enero–septiembre).
2. **Consistencia interna del tablero:** suma de todas las entidades contra la
   fila `27` en 33 meses × 7 conceptos. Diferencia máxima aceptada: 1 peso.
3. **Contra los BE de marzo, junio y septiembre de 2026:** 7 variables por
   entidad.
   - Junio y septiembre coinciden al centavo. En junio la única excepción
     documentada es el gasto de administración de SAE (`027019`).
   - Marzo presenta diferencias por **reexpresión** de la serie (por ejemplo,
     Klar `027009`): el tablero es la versión revisada.
4. **ROA y ROE calculados** contra `27200001` / `27200002` en 1,065
   comparaciones. Tolerancia: 0.02 puntos porcentuales.
5. **NICAP, IMOR e ICOR calculados** contra `27200093`, `27200005` y
   `27200013` (1,031, 933 y 933 comparaciones). Tolerancia: 0.02 puntos o
   `1e-5 × |publicado|` (valores publicados muy grandes, caso Stori `027032`).
6. **Cartera Total = E1 + E2 + E3** contra `130800103001` en 1,046
   entidades-mes.
7. **Estructura del libro:** nombres y orden de hojas, 420/420/315 filas, 33
   filas en el catálogo principal, 35 entidades por mes, fórmulas visibles, 11
   indicadores en el diccionario, 10 bloques de semáforo y 31 reglas de
   formato condicional por hoja.

## Cobertura y excepciones conocidas

- **027028 Acción y Evolución:** revocada el 7 de septiembre de 2026. No
  aparece en el tablero (el BE también la reporta como `n.d.`), pero conserva
  su lugar en las hojas de datos con toda la serie en `n.d.` hasta cerrar 2026.
  En el catálogo aparece en el apartado de bajas.
- **027014 Nu México:** transformada el 6 de agosto de 2026. Sigue en las hojas
  de datos con su histórico y `n.d.` en los meses sin reporte (desde agosto de
  2026); en el catálogo aparece en el apartado de bajas.
- **027040 Impulso:** no opera hasta marzo 2026 y no completa 12 meses de
  historia, así que ROA y ROE son `n.d.` aunque sus saldos y flujos existan.
- **027052 Trafalgar:** empieza a reportar en octubre 2025; sin 12 meses de
  historia, ROA y ROE son `n.d.`
- **Septiembre 2026:** es el corte más reciente y está **incompleto**. Capital
  Neto y Requerimiento no se publican ese mes, así que NICAP (y los
  indicadores que dependen de él) aparecen como `n.d.` desde septiembre.
- **027038 CAME:** el BE la reporta como "Revocada para operar el 19 de
  septiembre de 2025; última información agosto 2025", por lo que no entra en
  las 35. El tablero todavía conserva algunos meses suyos; por eso las sumas de
  las 35 no igualan la fila `27` en esos meses. La validación de la fila `27`
  se hace, por eso, con todas las entidades del tablero, no solo con las 35.
- **SAE (`027019`), gastos de administración:** el subtotal del BE de junio
  2026 no coincide con `601800301009` del tablero. El tablero es
  autoconsistente: su flujo de 12 meses (`27200039`) se reconstruye
  exactamente con la serie YTD de `601800301009`. Se conserva el tablero.
- **Stori (`027032`), NICAP de mayo 2024:** el requerimiento de capital del
  boletín viene en miles en vez de millones y la diferencia calculada es de
  125.75 puntos. La validación de NICAP/IMOR/ICOR la acepta por la tolerancia
  relativa (`1e-5 × |publicado|`), porque el error relativo es de 7.5e-6. Los
  datos del tablero no se reexpresan.
- **Reexpresión:** la serie del tablero está revisada. Las cifras de 2025
  pueden diferir de los boletines BE de 2025 que se usaron antes. Se conserva
  la serie revisada de la CNBV.

## Cómo actualizar

### Un mes nuevo de 2026

1. Descarga el tablero actualizado y reemplaza
   `entradas/tablero/sh_datos_csv_27_i/027_datos.csv`.
2. Agrega el mes a `ANIOS`, en `sofipos/config.py`:

```python
ANIOS = {
    2024: tuple(range(1, 13)),
    2025: tuple(range(1, 13)),
    2026: tuple(range(1, 10)),   # 9 = septiembre
}
```

3. Corre el programa. No hay que tocar fórmulas ni conceptos: el lector toma
   automáticamente el mes anterior.

Para probar un año sin generar el libro completo usa `--anios 2026` (o
`--anios 2024-2026`); el Diccionario y los Semáforos solo citan las hojas de
los años incluidos.

### Un año nuevo

Agrega el año a `ANIOS` (en `sofipos/config.py`) y listo: las hojas de datos e
indicadores, la validación del libro y los textos (título, notas del
catálogo) se derivan de `ANIOS`. Si el año tiene notas específicas,
agrégalas a `notas_por_anio` en `sofipos/hojas/datos_fuente.py`; si no, el
año corre solo con las notas comunes.

```python
ANIOS = {
    2024: tuple(range(1, 13)),
    2025: tuple(range(1, 13)),
    2026: tuple(range(1, 10)),   # 9 = septiembre
    2027: tuple(range(1, 13)),
}
```

### Una SOFIPO nueva

1. Debe aparecer en la hoja `Universo` del BE vigente con estatus "en
   operación".
2. Agrega su denominación social y sitio web al diccionario `CATALOGO` de
   `sofipos/catalogo.py`.

## Verificar que un cambio no altera el libro

`scripts/comparar.py` compara dos libros (y opcionalmente dos salidas de
`stdout`) celda por celda: valores, fórmulas, formatos, notas, anchos y
formato condicional. Sale con código 0 si son idénticos.

**Windows** (compara solo el .xlsx):

```powershell
.venv\Scripts\python.exe generar_sofipos_indicadores.py --out $env:TEMP\sofi.xlsx
Copy-Item $env:TEMP\sofi.xlsx $env:TEMP\referencia.xlsx

# ... hacer el cambio ...

.venv\Scripts\python.exe generar_sofipos_indicadores.py --out $env:TEMP\sofi.xlsx
.venv\Scripts\python.exe scripts\comparar.py --xlsx-a $env:TEMP\referencia.xlsx --xlsx-b $env:TEMP\sofi.xlsx
```

**macOS / Linux** (compara el .xlsx y el stdout):

```bash
# 1. generar una referencia con el código actual
.venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx > /tmp/salida1.txt
cp /tmp/sofi.xlsx /tmp/referencia.xlsx

# 2. hacer el cambio y volver a generar (misma ruta)
.venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx > /tmp/salida2.txt

# 3. comparar
.venv/bin/python scripts/comparar.py \
    --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
    --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt
```

Si también comparas el `stdout`, ambas corridas deben escribir en la **misma
ruta**: la última línea (`Archivo generado: …`) imprime la ruta exacta.

Sirve tanto para cambios grandes (como el refactor en `sofipos/`) como para
ajustes chicos: un `IDÉNTICO` significa que el libro no cambió en nada.

## Notas de mantenimiento

- Las reglas del semáforo viven en la lista `FICHAS` (`sofipos/fichas.py`).
  Si cambia un umbral, hay que cambiarlo ahí para que el libro se vuelva a
  producir igual; la hoja `Semáforos` se genera a partir de esa misma lista.
- Las estimaciones preventivas conservan el signo de la CNBV; por eso ICOR usa
  `ABS(...)`.

## Fuentes

- [Porfatolio Informativo de SOFIPOs](https://portafolioinfo.cnbv.gob.mx/Paginas/Inicio.aspx)
- [Tablero Power BI de entidades supervisadas por CNBV](https://app.powerbi.com/view?r=eyJrIjoiMTlhN2U5Y2QtYTcxNi00MzQ3LTg2MGYtNWI5ZjBhNDk2MmYyIiwidCI6IjVlMmM0OTc3LTEwN2QtNDBhMy04YWY3LTcwMDc0ODFhNjBkNCIsImMiOjR9)
- [Boletines de SOFIPOs](https://portafolioinfo.cnbv.gob.mx/Paginas/Contenidos.aspx?ID=27&Contenido=Boletines&Titulo=Sociedades%20Financieras%20Populares)
- [Nivel de Capitalización y Alertas Tempranas de SOFIPOs](https://portafolioinfo.cnbv.gob.mx/Paginas/Contenidos.aspx?ID=27&Contenido=Nivel%20de%20Capitalizaci%C3%B3n%20y%20Alertas%20Tempranas&Titulo=Sociedades%20Financieras%20Populares)
- [Series Históricas de SOFIPOs](https://app.powerbi.com/view?r=eyJrIjoiYjEwY2RhODYtYjg2ZS00YTQ3LWJmNzItMGVjZWY4MmNkNDY4IiwidCI6IjVlMmM0OTc3LTEwN2QtNDBhMy04YWY3LTcwMDc0ODFhNjBkNCIsImMiOjR9)

### Youtube

- [KUBO Financiero sigue en caída libre, Nu Bank pierde más de 1,300 millones. Análisis 2T 2025 Pt. 1](https://www.youtube.com/watch?v=M8wACwOYcU4)
- [Klar vs FINSUS vs Supertasas: ¿Cuál es más sólida? - Análisis 2T 2025 pt. 2](https://www.youtube.com/watch?v=Fgc_-SM46n4)
- [¿Está seguro tu dinero en estás SOFIPOS? Análisis 1T Pt. 1 (Nu, Finamigo, Klar, Fondeadora, Stori)](https://www.youtube.com/watch?v=zBY6bdW7HeU)
- [CAME intervenida y Kubo en la mira: lo que revelan los números ⚠️ Análisis 1T 2025 Pt.2](https://www.youtube.com/watch?v=51CWhLGsBhA)
- [¿Cómo cerraron las SOFIPOS en 2024? ➡️ Ojo con Kubo Financiero y Update de CAME 🚩](https://www.youtube.com/watch?v=cLkwQoLcRvU)
- [SOFIPOS Análisis 3T 2024 Pt. 1: Klar Sorprende y NuBank PIERDE más de 1,500 Millones ⚠️](https://www.youtube.com/watch?v=kHtTuFekOa8)
- [SOFIPOS Análisis 3T Pt. 2⚠️Cuidado con CAME➡️Super Tasas vuelve a ser Rentable✅FINSUS crece cartera!](https://www.youtube.com/watch?v=EHS6s0hlX2Q)
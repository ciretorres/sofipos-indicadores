# SOFIPOS · Generador de datos e indicadores

Programa que arma un solo libro de Excel con las 35 sociedades financieras
populares (SOFIPO) en operación: catálogo, cifras mensuales de 2025 y 2026,
indicadores, diccionario y semáforos.

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
│   └── SOFIPOS_2026_Datos.xlsx               7 hojas
├── generar_sofipos_2026.py                   generador
├── requirements.txt
└── README.md
```

`entradas/` y `salidas/` están en `.gitignore` porque pesan decenas de MB y
los datos los publica la CNBV. Para reconstruir la carpeta de entrada, sigue
las instrucciones de [`entradas/README.md`](entradas/README.md).

## Requisitos

- Python 3.9 o superior
- [`openpyxl`](https://openpyxl.readthedocs.io/) 3.1+ (ver `requirements.txt`)

## Instalación

```bash
git clone https://github.com/<tu-usuario>/sofipos-indicadores.git
cd sofipos-indicadores
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## Uso

Descarga primero los datos de entrada (ver
[`entradas/README.md`](entradas/README.md)) y después ejecuta:

```bash
.venv/bin/python generar_sofipos_2026.py
```

Sale con código 0 cuando todas las validaciones pasan. Opciones:

```bash
.venv/bin/python generar_sofipos_2026.py --entradas otra_carpeta
.venv/bin/python generar_sofipos_2026.py --salidas otra_carpeta
.venv/bin/python generar_sofipos_2026.py --out ruta/exacta.xlsx
.venv/bin/python generar_sofipos_2026.py --tablero ruta/027_datos.csv
```

## Hojas del libro

| Hoja | Contenido |
|---|---|
| `Catálogo SOFIPOS` | 35 entidades: No., Clave CASFIM, SOFIPO, Federación, Entidad Legal, Sitio Web, Fecha de Autorización, Nivel, Estatus |
| `Datos Fuente 2026` | Enero–agosto 2026 · 8 meses × 35 = 280 filas |
| `Datos Fuente 2025` | Enero–diciembre 2025 · 12 meses × 35 = 420 filas |
| `Indicadores 2026` | 280 filas, todas por fórmula contra `Datos Fuente 2026` |
| `Indicadores 2025` | 420 filas, todas por fórmula contra `Datos Fuente 2025` |
| `Diccionario` | 11 indicadores: fórmula, grupo de razón financiera y descripción |
| `Semáforos` | 10 bloques y 31 reglas de formato condicional |

Las dos hojas de datos tienen exactamente las mismas columnas, y las dos de
indicadores también: el año cambia, la estructura no.

## Fuente: tablero de la CNBV

`entradas/tablero/sh_datos_csv_27_i/027_datos.csv` es una serie larga
(`concepto, institucion, fecha, saldo_se, tipo_saldo`):

- **903 conceptos**, de **201601 a 202608**, frecuencia mensual.
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

- **Periodicidad mensual.** 2025 completo (enero–diciembre) y 2026 de enero a
  agosto.
- **Fecha de autorización.** La celda guarda una fecha real con formato
  `dd-mmm-yyyy`, de modo que ordenar y filtrar siguen funcionando. El nombre
  del mes se muestra según el idioma de Excel (en español, `18-dic-2008`).
- **Periodo anterior = mes inmediato anterior.** Enero 2025 se compara contra
  diciembre 2024; enero 2026 contra diciembre 2025. Las columnas de captación y
  cartera del mes anterior van como **dato**, no como fórmula, para que cada
  hoja sea autocontenida.
- **35 entidades en cada mes.** Si una entidad aún no operaba o no reportó el
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

### Columnas de `Datos Fuente 2026` y `Datos Fuente 2025`

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

### Columnas de `Indicadores 2026` y `Indicadores 2025`

| Col | Indicador | Fórmula |
|---|---|---|
| F | NICAP | `Datos Fuente!V` (parámetro) |
| G | IMOR | `H / I` |
| H | ICOR | `ABS(J) / H` |
| I | ROA | `W / X` |
| J | ROE | `W / Y` |
| K | Margen Financiero | `N / L` |
| L | Crecimiento de Captación | `Q / R − 1` |
| M | Crecimiento de Cartera | `I / S − 1` |
| N | Resultado Neto | `Datos Fuente!P` (parámetro, miles) |
| O | Crédito Neto | `K / F` |
| P | % Gastos de Administración | `O / N`, con `n.a.` si el margen es ≤ 0 |

Todas son fórmulas que apuntan a la hoja de datos del mismo año, en la misma
fila.

## Semáforos

| Indicador | Verde | Amarillo | Rojo |
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

1. **35 entidades en cada mes** de 2025 y 2026.
2. **Consistencia interna del tablero:** suma de todas las entidades contra la
   fila `27` en 20 meses × 7 conceptos. Diferencia máxima aceptada: 1 peso.
3. **Contra el BE de marzo y junio 2026:** 7 variables por entidad.
   - Junio coincide al centavo; la única excepción documentada es el gasto de
     administración de SAE (`027019`).
   - Marzo presenta diferencias por **reexpresión** de la serie (por ejemplo,
     Klar `027009`): el tablero es la versión revisada.
4. **ROA y ROE calculados** contra `27200001` / `27200002` en 639
   entidades-mes. Tolerancia: 0.02 puntos porcentuales.
5. **NICAP, IMOR e ICOR calculados** contra `27200093`, `27200005` y
   `27200013` (606, 568 y 568 comparaciones). Tolerancia: 0.02 puntos.
6. **Cartera Total = E1 + E2 + E3** contra `130800103001`.
7. **Estructura del libro:** nombres y orden de hojas, 280/420 filas, 35
   entidades por mes, fórmulas visibles, 11 indicadores en el diccionario, 10
   bloques de semáforo y 31 reglas de formato condicional por hoja.

## Cobertura y excepciones conocidas

- **027028 Acción y Evolución:** no aparece en el tablero (el BE también la
  reporta como `n.d.`). Toda la serie va en `n.d.`.
- **027040 Impulso:** no opera hasta marzo 2026 y no completa 12 meses de
  historia, así que ROA y ROE son `n.d.` aunque sus saldos y flujos existan.
- **027052 Trafalgar:** empieza a reportar en octubre 2025; sin 12 meses de
  historia, ROA y ROE son `n.d.`
- **Agosto 2026:** es el corte más reciente y está **incompleto**. Nu México
  (`027014`) todavía no reporta ese mes y Capital Neto / Requerimiento no se
  publican; por eso NICAP y los indicadores que dependen de ellos aparecen como
  `n.d.`
- **027038 CAME:** el BE la reporta como "Revocada para operar el 19 de
  septiembre de 2025; última información agosto 2025", por lo que no entra en
  las 35. El tablero todavía conserva algunos meses suyos; por eso las sumas de
  las 35 no igualan la fila `27` en esos meses. La validación de la fila `27`
  se hace, por eso, con todas las entidades del tablero, no solo con las 35.
- **SAE (`027019`), gastos de administración:** el subtotal del BE de junio
  2026 no coincide con `601800301009` del tablero. El tablero es
  autoconsistente: su flujo de 12 meses (`27200039`) se reconstruye
  exactamente con la serie YTD de `601800301009`. Se conserva el tablero.
- **Reexpresión:** la serie del tablero está revisada. Las cifras de 2025
  pueden diferir de los boletines BE de 2025 que se usaron antes. Se conserva
  la serie revisada de la CNBV.

## Cómo actualizar

### Un mes nuevo de 2026

1. Descarga el tablero actualizado y reemplaza
   `entradas/tablero/sh_datos_csv_27_i/027_datos.csv`.
2. Agrega el mes a `ANIOS` en el generador:

```python
ANIOS = {2025: tuple(range(1, 13)), 2026: tuple(range(1, 9))}   # 8 = agosto
```

3. Corre el programa. No hay que tocar fórmulas ni conceptos: el lector toma
   automáticamente el mes anterior.

### Un año nuevo

Agrega el año a `ANIOS` y crea la hoja correspondiente. El generador arma las
hojas de datos e indicadores con el mismo código (`escribir_datos_fuente` y
`escribir_indicadores` reciben el año). Habrá que agregar la llamada en
`escribir_xlsx` y el nombre esperado en `validar_libro`.

### Una SOFIPO nueva

1. Debe aparecer en la hoja `Universo` del BE vigente con estatus "en
   operación".
2. Agrega su denominación social y sitio web al diccionario `CATALOGO` del
   generador.

## Notas de mantenimiento

- Las reglas del semáforo viven en la lista `FICHAS`. Si cambia un umbral, hay
  que cambiarlo ahí para que el libro se vuelva a producir igual; la hoja
  `Semáforos` se genera a partir de esa misma lista.
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
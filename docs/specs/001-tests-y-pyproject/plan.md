# Plan: Tests y pyproject — 001

Fuente: `spec.md` (aprobado). Las 4 decisiones del gate están al final de
`spec.md`. El detalle de tareas está en `task.md`.

## Enfoque técnico

### 1. `pyproject.toml` (T001)

- Contenido exacto = el esqueleto de la spec: setuptools (`>=61`) como build
  backend, layout plano, `name = "sofipos-indicadores"`, `version = "0.1.0"`,
  `requires-python = ">=3.11"`, `dependencies = ["openpyxl>=3.1,<4"]`,
  extra `dev = ["pytest", "pyflakes"]`, entry point
  `sofipos = "sofipos.main:main"` y `[tool.pytest.ini_options]
  testpaths = ["tests"]`.
- Instalación: `.venv/bin/pip install -e ".[dev]"`. Smoke: `.venv/bin/sofipos
  --help` (argparse) y `python -c "import sofipos.hojas.semaforos"` para
  confirmar que `sofipos.hojas` queda empaquetado.
- Si setuptools se queja en el layout plano (auto-discovery), se agrega
  explícitamente:
  ```toml
  [tool.setuptools.packages.find]
  include = ["sofipos*"]
  ```

### 2. Fixture CSV (`tests/fixtures/tablero_mini.csv`) (T003)

Columnas exactas del `csv.DictReader` de `tablero.py`:
`concepto,institucion,fecha,saldo_se,tipo_saldo`. Una sola prueba reutiliza
este archivo (`test_tablero.py` y `test_filas.py`). ~25–30 filas:

| Grupo | Filas | Detalle |
|---|---|---|
| activos `027001` | 9 | fechas `202512`–`202608`, `tipo_saldo=130` |
| activos `027009` | 1 | solo `202601` → el resto de sus meses queda sin datos |
| sector `27` | 2 | activos `202601`, `202602` (va a `Tablero.sector`) |
| captación `027001` | 2 | `202512` y `202601` (para `captacion_prev`) |
| cartera directa `027001` `202601` | 1 | concepto `272724` (fallback directo) |
| cartera `027001` `202512` | 1 | `272724` (para `cartera_total_prev`) |
| cartera total−E3 `027001` `202602` | 2 | `130800103001` + `100800104003`, sin `272724` |
| cartera E1+E2 `027009` `202601` | 2 | `100800104001` + `100800104002` |
| gastos adm. `027001` `202601` | 1 | concepto `601800301009` |
| descartes | 4 | `tipo_saldo=140`; fecha `202501`; concepto `999999`; `saldo_se=n.d.` |

Notas:

- `institucion=27` se compara **en crudo** (`CLAVE_SECTOR`), por eso esa fila
  es `27` sin ceros; las sociedades pasan por `normalizar_clave` (6 dígitos).
- Con `fechas = {fecha_mes(2026, m) for m in ANIOS[2026]} | {"202512"}` se
  cubren los 8 meses de 2026 y el mes anterior de enero.

### 3. Universo y nombres en el test (T004)

Se construyen en el propio test, sin leer el BE:

```python
UNIVERSO = [
    {"clave": "027001", "consecutivo": 1, "nombre": "Sociedad Uno"},
    {"clave": "027009", "consecutivo": 2, "nombre": "Sociedad Dos"},
]
NOMBRES = {"027001": "Sociedad Uno Renombrada"}   # manda sobre reg["nombre"]
```

Casos que debe cubrir `test_filas.py`:

- nº de filas = `len(ANIOS[2026]) × 2` = 16.
- `a_miles`: `1_000_000` → `1000.0`.
- Tres fallbacks de cartera: E12 directo (`027001` `202601`), E1+E2
  (`027009` `202601`), total−E3 (`027001` `202602`).
- `tiene_datos`: `True` en meses con filas, `False` en los meses de `027009`
  sin registros.
- `nombres` manda: `027001` → «Sociedad Uno Renombrada»; `027009` (fuera de
  `NOMBRES`) → `reg["nombre"]`.
- Mes anterior: en la fila `202601`, `captacion_prev` = captación de `202512`
  en miles y `cartera_total_prev` = (E12+E3) de `202512` en miles; `mes` =
  «Enero», `trimestre` = «T1».

### 4. `test_semaforos.py` — leer la hoja y las reglas CF en memoria (T005)

Todo en un `Workbook` en memoria; no se genera el libro completo:

```python
wb = Workbook()
escribir_semaforos(wb)
ws = wb["Semáforos"]
```

Lecturas:

1. **Fila «Umbral»**: recorrer filas donde la columna A vale `"Umbral"`; B/C/D
   contienen las etiquetas unidas con `" · "` (verde/amarillo/rojo). Se
   emparejan **por orden** con `FICHAS_SEMAFORO` (la hoja itera esa misma
   lista).
2. **Nota de condiciones**: filas cuya columna A empieza con
   `"Excel (fracción): "`; se parten por `" · "` en pares `etiqueta →
   condición`.
3. **Reglas CF**: `aplicar_formato_condicional(wb.create_sheet("Indicadores"),
   nfilas=10)` y luego:
   ```python
   for cf in ws_ind.conditional_formatting:
       for regla in cf.rules:
           formulas.append(regla.formula[0])   # "AND(ISNUMBER(F5),F5>=1.41)"
   ```
   Des-envolver: quitar el prefijo `AND(ISNUMBER(<ref>),` y el `)` final; la
   regla `COND_NA` (`F5="n.a."`) no se envuelve. Recuento == `TOTAL_REGLAS`
   (31).
4. **Extracción de números**: primero eliminar las referencias de celda
   (`[A-Z]+\d+`) del texto — sin eso, el `5` de `F5` envenena los resultados —
   y luego aplicar `(\d+(?:\.\d+)?)%` a las etiquetas (÷ 100) y
   `-?\d+(?:\.\d+)?` a las condiciones.

Aserciones:

- Por indicador: `set(porcentajes de etiquetas ÷ 100) == set(números de la
  condición) == set(números de la fórmula CF desenvuelta)`.
- La condición de la nota es **idéntica (verbatim)** a la fórmula CF ya
  desenvuelta (usan la misma referencia `columna + str(FILA_DATOS)` → `F5`…).
- `len(formulas) == TOTAL_REGLAS`.
- Tramo `n.a.`: su condición (`COND_NA`) no tiene números; solo se valida que
  la etiqueta mencione `n.a.`.
- Notas en prosa (columna A, celdas que empiezan por `NICAP:`,
  `Crédito Neto:`, `% Gastos de Administración:`,
  `IMOR, ICOR, ROA, ROE…`): los porcentajes citados coinciden con los números
  de las condiciones de las fichas de `FICHAS` correspondientes. Para la nota
  de límites superiores, los 6 porcentajes `(10%, 100%, 2%, 15%, 70%, 80%)`
  se comparan **en orden** con los tramos amarillos de IMOR, ICOR, ROA, ROE,
  Margen Financiero y % Gastos.
- `FICHAS` de «Resultado Neto» tiene `tramos == ()` y la nota lo dice.

### 5. README y `requirements.txt` (T006)

- Las 4 referencias: línea 48 (árbol), línea 68 (Requisitos), líneas 96 y
  102 (bloques de instalación Windows/macOS) → `pip install -e ".[dev]"`.
- Línea 67: «Python 3.9 o superior» → «Python 3.11 o superior».
- Árbol: agregar `pyproject.toml` y `tests/`; quitar `requirements.txt`.
- Sección «Pruebas» (3–4 líneas) después de «Instalación»:
  ```bash
  .venv/bin/python -m pytest
  ```
  diciendo que los tests usan fixtures versionados y no requieren `entradas/`.
- `git rm requirements.txt` (borrado del archivo; sin commit).
- Índice: fila `001-tests-y-pyproject` con estado `borrador` en
  `docs/specs/README.md`.

### 6. Verificación con `comparar.py`

No se duplica el bloque de comandos: está en **AGENTS.md § «Verificar un
cambio»**. Al **inicio** de la implementación se corre el paso 1 (generar
`/tmp/referencia.xlsx` + `/tmp/salida1.txt`); en el gate final (T007) se corren
los pasos 2 y 3 → `IDÉNTICO`.

## Archivos

| Acción | Archivo |
|---|---|
| Crear | `pyproject.toml` |
| Crear | `tests/fixtures/tablero_mini.csv` |
| Crear | `tests/test_utilidades.py`, `tests/test_config.py`, `tests/test_tablero.py`, `tests/test_filas.py`, `tests/test_semaforos.py` |
| Modificar | `README.md` (4 refs + línea Python + árbol + sección «Pruebas») |
| Modificar | `docs/specs/README.md` (fila del índice) |
| Modificar | `spec.md` (solo el bloque «Decisiones del gate») |
| Borrar | `requirements.txt` |
| Sin cambios | `sofipos/`, `generar_sofipos_2026.py`, `scripts/comparar.py` |

## Orden de cambios

```
T001 pyproject + instalación      (base: todo lo demás se instala y corre)
T002 test_utilidades + test_config (sin fixture, más rápido)
T003 fixture CSV + test_tablero
T004 test_filas                    (depende del fixture de T003)
T005 test_semaforos                (independiente; puede ir tras T002)
T006 README + borrar requirements  (después de probar el install de T001)
T007 gate final                    (todo lo anterior)
```

Checkpoint tras T004: `pytest` en verde con 4 archivos y `pyflakes` limpio.

## Riesgos

| Riesgo | Impacto | Mitigación |
|---|---|---|
| setuptools no descubre el paquete en layout plano | Alto | `[tool.setuptools.packages.find] include = ["sofipos*"]`; verificar con `import sofipos.hojas.semaforos` |
| `pip install -e` pide red para el build backend | Medio | Reintentar con `--no-build-isolation` |
| Los dígitos de `F5` contaminan la extracción de números | Medio | Eliminar referencias de celda antes del regex (ver §4) |
| `leer_tablero` imprime a stdout | Bajo | pytest lo captura; sin impacto |
| `comparar.py` necesita `entradas/` completo | Medio | Ya está en el workspace; si falta, ver `entradas/README.md` |
| Leer reglas CF cambia entre versiones de openpyxl | Bajo | Fijado `openpyxl>=3.1,<4`; API `ws.conditional_formatting` estable |

## Verificación final (gate T007)

```bash
# 1. Instalación editable con extras dev
.venv/bin/pip install -e ".[dev]"

# 2. Pruebas (los 5 archivos, testpaths=tests)
.venv/bin/python -m pytest

# 3. Lint — 4 rutas
.venv/bin/pyflakes sofipos tests generar_sofipos_2026.py scripts

# 4. CLI instalada
.venv/bin/sofipos --help

# 5. Regresión del libro → IDÉNTICO (pasos 2 y 3 de AGENTS.md §
#    «Verificar un cambio»; el paso 1 se corrió al inicio)
.venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
.venv/bin/python scripts/comparar.py \
    --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
    --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt

# 6. Sin residuos de requirements.txt (grep sin coincidencias = exit 1)
grep -n "requirements.txt" README.md
test ! -e requirements.txt
```

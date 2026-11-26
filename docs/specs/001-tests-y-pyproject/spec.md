# Spec: Tests y pyproject — 001

Estado: terminada

## Objective

Hoy el repo no tiene pruebas, ni `pyproject.toml`, ni CI: solo `requirements.txt`
(1 línea: `openpyxl>=3.1,<4`) y la verificación manual del libro con
`scripts/comparar.py`. Esta spec:

1. Agrega **`pyproject.toml`** como fuente única de metadatos, dependencias,
   entry point y configuración de pytest.
2. Agrega **`tests/`** con fixtures CSV mini **versionados** (no dependen de
   `entradas/`, que está gitignorado) que cubran utilidades, config, semáforos,
   tablero y filas.
3. Actualiza **`README.md`**: las 4 referencias a `requirements.txt` pasan a
   `pip install -e ".[dev]"` y el árbol de estructura refleja los cambios.
4. **Elimina `requirements.txt`** (el pyproject queda como fuente única).

Nada de esto cambia el libro generado: el paquete `sofipos/` **no se toca**.

Éxito central: pytest en verde, `pyflakes` limpio, instalación editable con CLI
funcionando y `scripts/comparar.py` → `IDÉNTICO`.

## Tech Stack

- Python ≥ 3.11 (venv local: 3.15; el MCP de Excel pide ≥ 3.11 — ver Open Question 2).
- [`openpyxl`](https://openpyxl.readthedocs.io/) `>=3.1,<4` (única dependencia de runtime).
- `pytest` y `pyflakes` como dependencias extra `dev`.
- `setuptools` como build backend; layout plano (el paquete `sofipos/` se descubre solo).
- Sin cambios en `sofipos/`: los tests importan el código existente tal cual.

Esqueleto de `pyproject.toml` (todo lo demás es default de setuptools):

```toml
[project]
name = "sofipos-indicadores"
version = "0.1.0"
requires-python = ">=3.11"
dependencies = ["openpyxl>=3.1,<4"]

[project.optional-dependencies]
dev = ["pytest", "pyflakes"]

[project.scripts]
sofipos = "sofipos.main:main"

[build-system]
requires = ["setuptools>=61"]
build-backend = "setuptools.build_meta"

[tool.pytest.ini_options]
testpaths = ["tests"]
```

## Commands

```bash
# Instalar (una vez; reemplaza a pip install -r requirements.txt)
.venv/bin/pip install -e ".[dev]"
# Windows: .venv\Scripts\python.exe -m pip install -e ".[dev]"

# Pruebas
.venv/bin/python -m pytest

# Lint
.venv/bin/pyflakes sofipos tests

# CLI instalada
.venv/bin/sofipos --help

# Regresión obligatoria del libro (ver AGENTS.md; referencia y cambio)
.venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
.venv/bin/python scripts/comparar.py \
    --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
    --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt   # → IDÉNTICO
```

## Project Structure

```
sofipos-indicadores/
├── pyproject.toml              # NUEVO: metadatos, deps, entry point, pytest
├── sofipos/                    # SIN CAMBIOS (2,880 líneas)
├── tests/                      # NUEVO
│   ├── fixtures/
│   │   └── tablero_mini.csv    # CSV mini versionado (columnas del tablero)
│   ├── test_config.py
│   ├── test_filas.py
│   ├── test_semaforos.py
│   ├── test_tablero.py
│   └── test_utilidades.py
├── scripts/comparar.py         # SIN CAMBIOS
├── README.md                   # ACTUALIZA 4 refs a requirements.txt + árbol
├── requirements.txt            # ELIMINADO
└── docs/specs/001-tests-y-pyproject/   # esta spec + plan.md + task.md
```

## Code Style

Español en docstrings y mensajes, como el resto del repo; nombres descriptivos;
sin relleno. Ejemplo del estilo esperado en los tests:

```python
def test_fecha_anterior_cruza_el_anio():
    """Enero regresa a diciembre del año anterior."""
    assert fecha_anterior("202601") == "202512"
    assert fecha_mes(2026, 1) == "202601"
```

- Un test = una aserción clara por comportamiento; nombres en español.
- Los tests importan de `sofipos.*` directamente; no duplican lógica del producto.
- Los fixtures viven en `tests/fixtures/` y se leen con `Path(__file__).parent`.

## Testing Strategy

Framework: **pytest**, `testpaths = ["tests"]`. Correr con
`.venv/bin/python -m pytest` desde la raíz (eso mete el CWD en `sys.path`, además
de la instalación editable). Cobertura: las 5 áreas siguientes; sin thresholds
de cobertura en esta versión.

**Seguridad de imports (investigado):** importar módulos de `sofipos/` no tiene
efectos secundarios — a nivel de módulo solo se definen constantes y funciones;
`sofipos/__main__.py` solo corre al ejecutarse como programa; `config.py` solo
calcula rutas en memoria (no toca disco). `leer_tablero` imprime a stdout, que
no afecta a los tests.

| Archivo | Qué cubre |
|---|---|
| `test_utilidades.py` | `fecha_mes`, `fecha_anterior` con borde enero → diciembre del año anterior |
| `test_config.py` | `CONCEPTOS`: claves sin duplicados, claves solo dígitos, valores únicos; `ANIOS`: años enteros, meses en 1–12 sin repetir y ordenados |
| `test_semaforos.py` | consistencia de umbrales (ver más abajo) |
| `test_tablero.py` | `leer_tablero` con CSV fixture: filtros de `tipo_saldo`, fecha, concepto y `saldo_se`; fila `institucion=27` → `sector`; claves normalizadas |
| `test_filas.py` | `preparar_filas` con el mismo CSV: nº de filas = meses × universo; `a_miles`; fallbacks de cartera (E12 directo / E1+E2 / total−E3); `tiene_datos`; `nombres` manda sobre `reg["nombre"]`; columnas del mes anterior |

**Fixture CSV** (`tests/fixtures/tablero_mini.csv`): columnas
`concepto, institucion, fecha, saldo_se, tipo_saldo` (las que lee
`csv.DictReader` en `tablero.py`). Dimensión: ~30 filas — 2 sociedades + 1 fila
de sector (`institucion=27`), fechas `202512`–`202608` (cubre `ANIOS[2026]` y
su mes anterior), 4–6 conceptos de `CONCEPTOS`, más filas que deben descartarse
(`tipo_saldo` ≠ 130, fecha fuera del set, concepto desconocido, `saldo_se` no
numérico). El universo y los nombres para `preparar_filas` se construyen como
`list[dict]`/`dict` en el propio test (claves `clave`, `consecutivo`, `nombre`),
sin leer el BE.

### `test_semaforos.py` — dónde viven los umbrales (hallazgo)

Hoy los números viven en **tres** referencias y nada los cruza:

1. `sofipos/fichas.py` — `FICHAS[*].tramos` guarda por tramo la **etiqueta**
   (`"≥ 141%"`) y la **condición de Excel** (`"{c}>=1.41"`). Una puede cambiar
   sin la otra y nada falla.
2. `sofipos/hojas/semaforos.py` — la hoja «Semáforos» escribe las etiquetas en
   la fila «Umbral» y las condiciones en la nota «Excel (fracción)», leyendo
   `ficha.tramos`; además incluye **notas en prosa con los números escritos a
   mano** (NICAP 130/141, Crédito Neto 60/70/85/90, % Gastos 60/80, límites
   superiores de IMOR/ICOR/ROA/ROE/Margen/Gastos).
3. `sofipos/hojas/indicadores.py` — `aplicar_formato_condicional` arma los
   `FormulaRule` a partir de la **condición** de `tramos`, envuelta en
   `AND(ISNUMBER(...), …)` (o `COND_NA` para `n.a.`).

El test construye un `Workbook` en memoria (no genera el libro completo):

1. `escribir_semaforos(wb)` → lee de la hoja «Semáforos» las etiquetas de la
   fila «Umbral» y las condiciones de la nota «Excel (fracción)».
2. `aplicar_formato_condicional(ws, nfilas)` sobre una hoja «Indicadores»
   vacía → lee las fórmulas de las reglas de formato condicional.

Aserciones:

- Por indicador: los porcentajes de las etiquetas ÷ 100 == los números de las
  condiciones == los números de las fórmulas de formato condicional.
- Las condiciones de la nota son idénticas (verbatim) a las fórmulas ya
  des-envueltas de `AND(ISNUMBER(...))`.
- Número de reglas de formato condicional == `TOTAL_REGLAS`.
- Excepción `COND_NA` (tramo `n.a.`): la condición no tiene números; solo se
  valida que la etiqueta mencione `n.a.`.
- Notas en prosa: los porcentajes citados para NICAP, Crédito Neto, % Gastos y
  los límites superiores (IMOR 10, ICOR 100, ROA 2, ROE 15, Margen 70,
  Gastos 80) coinciden con los tramos de `FICHAS` correspondientes.

## Boundaries

- **Always:** pytest verde + `pyflakes` limpio antes de cerrar tarea; correr la
  verificación `comparar.py` → `IDÉNTICO`; tests sin depender de `entradas/` ni
  `salidas/`; estilo español del repo.
- **Ask first:** cambiar `requires-python` o las dependencias del pyproject;
  tocar README fuera de las 4 referencias + árbol; eliminar `requirements.txt`
  (Open Question 1); añadir tests que generen el libro completo.
- **Never:** modificar `sofipos/`; hacer `git commit` / `git push`; tests que
  lean `entradas/`; bajar o quitar aserciones para que pasen.

## Success Criteria

1. `.venv/bin/python -m pytest` → exit 0; existen y pasan los 5 archivos de test.
2. `.venv/bin/pyflakes sofipos tests` → sin salida (exit 0).
3. `.venv/bin/pip install -e ".[dev]"` completa sin error y `.venv/bin/sofipos --help`
   → exit 0 con el uso de argparse.
4. **`scripts/comparar.py` → `IDÉNTICO`** (exit 0): el libro generado no cambia.
5. `README.md` no contiene `requirements.txt` (0 coincidencias) y su árbol
   muestra `pyproject.toml` y `tests/`; `requirements.txt` ya no existe.

## Out of Scope (Deferred)

- CI / GitHub Actions — se plantea después de tener pytest estable.
- `mypy` / tipado estricto.
- Golden end-to-end completo (libro generado comparado contra un snapshot).
- Cualquier cambio en `sofipos/` (si un test revela un bug, va en otra spec).
- Tests de `be.py`, `xlsx_lector.py`, validaciones y `catalogo.py`.
- Umbrales de cobertura (`--cov`) en pytest.

## Open Questions

1. **Confirmar la eliminación de `requirements.txt`**: el pyproject pasa a ser
   fuente única. ¿Se borra en esta spec?
2. **Confirmar `requires-python = ">=3.11"`**: el venv local es 3.15 y el MCP
   de Excel pide ≥ 3.11, pero el README (línea 67) dice «Python 3.9 o superior».
   Si se confirma 3.11, ¿actualizamos esa línea del README también?
3. ¿Quién actualiza el índice de `docs/specs/README.md` (estado `borrador`)?
   Esta spec solo escribe en su carpeta.
4. ¿Añadir una sección «Pruebas» al README (cómo correr pytest)? No está en el
   alcance acordado.

## Decisiones del gate

Respuestas del usuario al aprobar la spec (cierran los Open Questions):

1. **OQ1 — Sí, eliminar `requirements.txt`:** el `pyproject.toml` queda como
   fuente única de metadatos y dependencias.
2. **OQ2 — Sí, `requires-python = ">=3.11"`:** y actualizar la línea ~67 del
   README de «Python 3.9 o superior» a «Python 3.11 o superior».
3. **OQ3 — Índice:** lo actualiza el planner en `docs/specs/README.md` con
   estado `borrador`; el usuario lo pasa a `aprobada` al aprobar el plan.
4. **OQ4 — Sí, añadir sección «Pruebas» al README:** 3–4 líneas con cómo
   correr pytest.

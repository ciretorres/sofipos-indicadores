# Tareas: Tests y pyproject — 001

Ordenadas por dependencia; cada tarea se marca `[x]` solo cuando su
verificación pasa. Convención de comandos: desde la raíz del repo, con el
`.venv` activo.

## T001 — `pyproject.toml` + instalación editable + smoke de CLI

- [x] **Primero (antes de cualquier cambio):** generar la referencia de
  regresión según AGENTS.md § «Verificar un cambio» paso 1 — deja
  `/tmp/referencia.xlsx` y `/tmp/salida1.txt` listos para T007.
- **Archivos:** crear `pyproject.toml` (contenido = esqueleto de `spec.md`
  § Tech Stack: setuptools, `requires-python = ">=3.11"`, extra `dev`,
  entry point `sofipos = "sofipos.main:main"`, `testpaths = ["tests"]`).
- **Aceptación:**
  - `.venv/bin/pip install -e ".[dev]"` completa sin error.
  - `.venv/bin/sofipos --help` sale con código 0 y muestra el uso de argparse.
  - `.venv/bin/python -c "import sofipos.hojas.semaforos"` no falla (el
    paquete `sofipos.hojas` queda empaquetado).
- **Verificación:**
  ```bash
  .venv/bin/pip install -e ".[dev]"
  .venv/bin/sofipos --help
  .venv/bin/python -c "import sofipos.hojas.semaforos"
  ```
- **Depende de:** ninguna. Si el install falla por auto-discovery, agregar
  `[tool.setuptools.packages.find] include = ["sofipos*"]` (ver plan.md).

## T002 — Tests de `utilidades` y `config`

- [x] **Archivos:** crear `tests/test_utilidades.py`, `tests/test_config.py`.
- **Aceptación:**
  - `test_utilidades.py`: `fecha_mes` (`fecha_mes(2026, 1) == "202601"`) y
    `fecha_anterior` con el borde enero → diciembre del año anterior
    (`fecha_anterior("202601") == "202512"`).
  - `test_config.py`: `CONCEPTOS` — claves sin duplicadas, claves solo
    dígitos y valores únicos; `ANIOS` — años enteros y, en cada uno, meses en
    1–12 ordenados y sin repetir.
  - Nombres de test en español, sin dependencias de `entradas/`.
- **Verificación:**
  ```bash
  .venv/bin/python -m pytest tests/test_utilidades.py tests/test_config.py
  .venv/bin/pyflakes tests/test_utilidades.py tests/test_config.py
  ```
- **Revisión (post-review):** faltaba la 3.ª aserción del spec § Testing
  Strategy (`CONCEPTOS` claves sin duplicadas); añadida como
  `test_conceptos_claves_sin_duplicadas` → 4 tests en verde.

## T003 — Fixture CSV mini + `test_tablero.py`

- [x] **Archivos:** crear `tests/fixtures/tablero_mini.csv` y
  `tests/test_tablero.py`.
- **Aceptación:**
  - El CSV tiene las columnas `concepto,institucion,fecha,saldo_se,tipo_saldo`
    y la composición de la tabla de `plan.md` §2 (2 sociedades + fila `27`,
    fechas `202512`–`202608`, 4+ conceptos, 4 filas que deben descartarse).
  - `leer_tablero(ruta, fechas)` descarta: `tipo_saldo ≠ 130`, fecha fuera
    del set, concepto desconocido y `saldo_se` no numérico.
  - Fila `institucion=27` cae en `tablero.sector` y **no** en `tablero.valores`.
  - Las claves de `tablero.valores` quedan normalizadas a 6 dígitos.
  - `tablero.dato("027001", "202601", "activos")` devuelve el valor del
    fixture.
- **Verificación:**
  ```bash
  .venv/bin/python -m pytest tests/test_tablero.py
  .venv/bin/pyflakes tests/test_tablero.py
  ```
- **Revisión (post-review):** el test de descartes acoplaba el texto del log
  (`"22 registros útiles"`); separado en `test_contenido_del_tablero`
  (claves y nº de registros) + `test_descarta_las_cuatro_filas_invalidas`
  (cuenta filas del CSV vs. entradas del objeto, sin leer el log) → 5 tests
  en verde.

## T004 — `test_filas.py` (reusa el fixture)

- [x] **Archivos:** crear `tests/test_filas.py`.
- **Aceptación:**
  - Universo/nombres construidos en el test como `list[dict]`/`dict`
    (claves `clave`, `consecutivo`, `nombre` — ver plan.md §3); sin BE.
  - Nº de filas = `len(ANIOS[2026]) × len(UNIVERSO)` = 16.
  - `a_miles` convierte pesos → miles con 2 decimales.
  - Los 3 fallbacks de cartera: E12 directo, E1+E2 y total−E3.
  - `tiene_datos` es `False` en los meses de `027009` sin registros.
  - `nombres` manda sobre `reg["nombre"]`; la clave ausente usa
    `reg["nombre"]`.
  - En la fila de enero 2026: `captacion_prev` y `cartera_total_prev` vienen
    de `202512`; `mes == "Enero"` y `trimestre == "T1"`.
- **Verificación:**
  ```bash
  .venv/bin/python -m pytest tests/test_tablero.py tests/test_filas.py
  .venv/bin/pyflakes tests/test_filas.py
  ```

## T005 — `test_semaforos.py` (hoja ↔ CF ↔ FICHAS ↔ notas)

- [x] **Archivos:** crear `tests/test_semaforos.py`.
- **Aceptación:**
  - `Workbook` en memoria: `escribir_semaforos(wb)` y
    `aplicar_formato_condicional` sobre una hoja «Indicadores» vacía; no se
    genera el libro completo.
  - Por indicador: porcentajes de las etiquetas ÷ 100 == números de las
    condiciones de la nota == números de las fórmulas CF desenvuelidas.
  - Las condiciones de la nota son verbatim a las fórmulas CF desenvueltas
    (leídas con `ws.conditional_formatting`).
  - Número de reglas CF == `TOTAL_REGLAS` (31).
  - Tramo `COND_NA`: sin números en la condición; la etiqueta menciona
    `n.a.`.
  - Notas en prosa (NICAP, Crédito Neto, % Gastos, límites superiores
    IMOR/ICOR/ROA/ROE/Margen/Gastos): porcentajes coincidentes con `FICHAS`;
    «Resultado Neto NO tiene semáforo».
- **Verificación:**
  ```bash
  .venv/bin/python -m pytest tests/test_semaforos.py
  .venv/bin/pyflakes tests/test_semaforos.py
  ```

## T006 — README + borrar `requirements.txt`

- [x] **Archivos:** `README.md` (modificar), `requirements.txt` (borrar).
- **Aceptación:**
  - `grep -n "requirements.txt" README.md` → 0 coincidencias (las 4 refs:
    árbol, Requisitos, instalación Windows y macOS) → `pip install -e ".[dev]"`.
  - Línea ~67: «Python 3.11 o superior» (antes 3.9).
  - Árbol de estructura: aparecen `pyproject.toml` y `tests/`; desaparece
    `requirements.txt`.
  - Sección «Pruebas» (3–4 líneas) con `.venv/bin/python -m pytest`, tras
    «Instalación».
  - `requirements.txt` ya no existe en el repo.
- **Verificación:**
  ```bash
  grep -n "requirements.txt" README.md        # sin coincidencias
  grep -n "Python 3.11" README.md             # 1 coincidencia
  grep -n "pyproject.toml\|tests/" README.md  # árbol actualizado
  test ! -e requirements.txt                  # exit 0
  ```

## T007 — Gate final

- [x] **Archivos:** ninguno (solo verificación).
- **Aceptación:** todos los checks del bloque de `plan.md` § «Verificación
  final» pasan:
  - `pip install -e ".[dev]"` sin error.
  - `python -m pytest` → exit 0 con los 5 archivos de test.
  - `pyflakes sofipos tests generar_sofipos_2026.py scripts` → sin salida.
  - `sofipos --help` → exit 0.
  - `scripts/comparar.py` → `IDÉNTICO` (pasos 2 y 3 de AGENTS.md §
    «Verificar un cambio»; el paso 1 se corrió antes de empezar).
  - `grep requirements.txt README.md` sin coincidencias y
    `test ! -e requirements.txt`.
- **Verificación:** el bloque completo de `plan.md` § «Verificación final».

## Checkpoint

Tras T004 conviene pausar: `pytest` con 4 archivos en verde y `pyflakes`
limpio antes de entrar a T005 (la tarea más delicada).

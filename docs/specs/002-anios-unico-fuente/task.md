# Tareas: Años desde una única fuente — 002

Orden: T001 → T007 (T001 es prerequisito de todas; T007 cierra el gate).
Marca `[x]` solo cuando la verificación de esa tarea pasa.

Tras cada tarea que toca código, repite los pasos 2-3 de AGENTS.md:

```bash
.venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
.venv/bin/python scripts/comparar.py \
    --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
    --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt
```

## Tareas

- [x] **T001 — Baseline antes de tocar nada**
  - Archivos: ninguno (solo artefactos en `/tmp`).
  - Aceptación: existen `/tmp/referencia.xlsx` y `/tmp/salida1.txt`
    generados con el código actual, sin errores.
  - Verificar:
    ```bash
    .venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida1.txt
    cp /tmp/sofi.xlsx /tmp/referencia.xlsx
    ```

- [x] **T002 — `sofipos/hojas/__init__.py`: bucles, título y docstring**
  - Archivos: `sofipos/hojas/__init__.py`.
  - Aceptación:
    - `anios = sorted(filas, reverse=True)`; dos bucles (datos fuente para
      todos los años, luego indicadores) → orden de hojas actual intacto.
    - Título derivado con `rango = f"{min(anios)}-{max(anios)}" if
      len(anios) > 1 else str(anios[0])` → sin «2026-2026» en año único.
    - Docstring sin «2025 y 2026».
    - `grep -n '2025\|2026' sofipos/hojas/__init__.py` → sin salida.
  - Verificar:
    ```bash
    .venv/bin/pyflakes sofipos/hojas/__init__.py
    # comparar.py → IDÉNTICO (bloque de arriba)
    ```

- [x] **T003 — `validacion_libro.py`: `hojas_esperadas(filas)` + test nuevo**
  - Archivos: `sofipos/validacion_libro.py`, `tests/test_validacion_libro.py`
    (nuevo).
  - Aceptación:
    - Función de nivel módulo `hojas_esperadas(filas)` con la forma del
      «Code Style» de la spec; `validar_libro` la usa en lugar de la lista a
      mano.
    - Test cubre `{2019, 2021}` (orden descendente exacto), `{2026}` (5
      hojas), `{2025, 2026}` (las 7 hojas actuales en orden actual).
    - `grep -n '2025\|2026' sofipos/validacion_libro.py` → sin salida.
  - Verificar:
    ```bash
    .venv/bin/python -m pytest tests/test_validacion_libro.py
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos/validacion_libro.py tests/test_validacion_libro.py
    # comparar.py → IDÉNTICO (bloque de arriba)
    ```

- [x] **T004 — `datos_fuente.py`: `notas_por_anio.get(anio, [])`**
  - Archivos: `sofipos/hojas/datos_fuente.py`.
  - Aceptación:
    - `notas_por_anio = {2026: [...], 2025: [...]}` con **solo** las notas
      específicas; `notas = notas_comunes + notas_por_anio.get(anio, [])`.
    - Sin `if anio == 2026`; los textos de 2026 y 2025 byte a byte idénticos
      a los actuales (la nota de la serie reexpresada se repite en ambas
      entradas, sin deduplicar).
  - Verificar:
    ```bash
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos/hojas/datos_fuente.py
    # comparar.py → IDÉNTICO (criterio 6: textos intactos)
    ```

- [x] **T005 — `utilidades.enumerar()` + `semaforos.py`/`diccionario.py`**
  - Archivos: `sofipos/utilidades.py`, `sofipos/hojas/semaforos.py`,
    `sofipos/hojas/diccionario.py`, `tests/test_utilidades.py`.
  - Aceptación:
    - `enumerar()`: 0 → `""`, 1 → el valor, 2+ →
      `", ".join(v[:-1]) + " e " + v[-1]`.
    - Tests: 1 elemento → «Indicadores 2026»; 2 → «Indicadores 2026 e
      Indicadores 2025» (verbatim el texto actual); 3+ → «A, B e C»; lista
      vacía → `""`.
    - `semaforos.py` y `diccionario.py` importan `ANIOS` de `..config` y
      usan `enumerar([f"Indicadores {a}" for a in sorted(ANIOS,
      reverse=True)])`; la prosa producida es idéntica a la actual.
    - `grep -n '2025\|2026' sofipos/hojas/semaforos.py sofipos/hojas/diccionario.py`
      → sin salida.
  - Verificar:
    ```bash
    .venv/bin/python -m pytest tests/test_utilidades.py tests/test_semaforos.py
    .venv/bin/pyflakes sofipos/utilidades.py sofipos/hojas/semaforos.py \
        sofipos/hojas/diccionario.py tests/test_utilidades.py
    # comparar.py → IDÉNTICO (bloque de arriba)
    ```

- [x] **T006 — `main.py` descripción derivada (OQ2) + README (OQ1)**
  - Archivos: `sofipos/main.py`, `README.md`.
  - Aceptación:
    - Descripción del parser derivada de `ANIOS` (propuesta del plan:
      «Genera el Excel mensual de las 35 SOFIPOS (2025, 2026).»); sin años
      literales en la descripción (el `help` de `--out` cita
      `SOFIPOS_2026_Datos.xlsx` = `SALIDA_DEFECTO`: es contenido, fuera de
      alcance).
    - `sofipos --help` → exit 0; el texto nuevo se revisa (lo fija el gate).
    - README «Un año nuevo» (397-403): queda «agrega el año a `ANIOS` y
      listo», más las notas específicas del año en `notas_por_anio` si las
      hay; sin instrucciones de tocar `escribir_xlsx`/`validacion_libro`.
  - Verificar:
    ```bash
    .venv/bin/sofipos --help   # exit 0 y revisar el texto
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos/main.py
    # comparar.py → IDÉNTICO (el texto de --help no entra al libro)
    ```

- [x] **T007 — Gate final**
  - Archivos: ninguno.
  - Aceptación: los 6 criterios de «Success Criteria» de la spec en verde.
  - Verificar (todo exit 0 / sin salida):
    ```bash
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos tests generar_sofipos_2026.py scripts
    .venv/bin/sofipos --help
    .venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
    .venv/bin/python scripts/comparar.py \
        --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
        --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt   # → IDÉNTICO
    grep -n '2025\|2026' sofipos/hojas/__init__.py sofipos/validacion_libro.py \
        sofipos/hojas/semaforos.py sofipos/hojas/diccionario.py   # → sin salida
    ```

# Spec: Años desde una única fuente — 002

Estado: terminada

## Objective

Hoy los años del libro están hardcodeados en ~5 lugares además de `ANIOS`
(`sofipos/config.py`): agregar o quitar un año obliga a editar varias carpetas a
ciegas y el README lo documenta como trabajo manual. Esta spec refactoriza para
que **`ANIOS` sea la única fuente de verdad de los años del libro**: después del
cambio, agregar o quitar un año = **1 línea en `ANIOS`**.

Refactor puro: **no agrega 2024** (eso es la spec 003) ni cambia ningún texto
visible. Su criterio central: el libro debe quedar **`IDÉNTICO`**
(`scripts/comparar.py`).

Los 5 lugares a derivar (verificados en el código):

1. **`sofipos/hojas/__init__.py`** (`escribir_xlsx`) — líneas 41-44 llaman
   `escribir_datos_fuente(wb, filas[2026], 2026, …)` y `escribir_indicadores(…)`
   a mano ( KeyError si falta un año ); línea 49-52 el título
   «… 2025-2026»; línea 27 el docstring «para 2025 y 2026».
2. **`sofipos/validacion_libro.py`** — líneas 41-49: la lista `esperadas` de
   hojas está escrita a mano. Las validaciones por año (líneas 79+ y 210-221)
   ya son genéricas sobre `filas`.
3. **`sofipos/hojas/datos_fuente.py`** — líneas 140-153:
   `notas = notas_2026 if anio == 2026 else notas_2025` (bug latente: un año
   desconocido recibiría las notas de 2025).
4. **`sofipos/hojas/semaforos.py`** (líneas 205-206) — prosa «hojas Indicadores
   2026 e Indicadores 2025».
5. **`sofipos/hojas/diccionario.py`** (línea 38) — la misma prosa.

Verificado: `filas` (parámetro de `escribir_xlsx` y `validar_libro`) se construye
en `main.py` con `for anio in ANIOS`, así que sus llaves **son** las llaves de
`ANIOS`; derivar de `filas` es equivalente a derivar de `ANIOS`.

**Requisito de diseño:** soportar **subconjuntos arbitrarios y no contiguos**.
`ANIOS` puede ser `{2024, 2026}` sin 2025, `{2019, 2021}` o solo `{2026}`. Nada
asume años consecutivos, ni que 2026 exista, ni un orden fijo de llaves.

## Tech Stack

- Python ≥ 3.11; sin dependencias nuevas (openpyxl, pytest, pyflakes ya en `.[dev]`).
- Cambios solo en el paquete `sofipos/` + tests; sin cambios de CLI ni de datos.

## Commands

```bash
# 1. Regresión obligatoria del libro (baseline ANTES del cambio; ver AGENTS.md)
.venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida1.txt
cp /tmp/sofi.xlsx /tmp/referencia.xlsx
# 2. Tras el cambio, misma ruta:
.venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
# 3. Comparar → debe decir IDÉNTICO (exit 0)
.venv/bin/python scripts/comparar.py \
    --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
    --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt

# Pruebas (28 actuales + los nuevos)
.venv/bin/python -m pytest

# Lint
.venv/bin/pyflakes sofipos tests generar_sofipos_2026.py scripts

# CLI
.venv/bin/sofipos --help   # exit 0; revisar el texto si cambia

# Sin años literales en los archivos refactorizados (ver Success Criteria 5)
grep -n '2025\|2026' sofipos/hojas/__init__.py sofipos/validacion_libro.py \
    sofipos/hojas/semaforos.py sofipos/hojas/diccionario.py
```

## Project Structure

```
sofipos/
├── config.py                  # SIN CAMBIOS — fuente única ANIOS (ya existe)
├── utilidades.py              # NUEVO helper: enumerar() (uno–pocos renglones)
├── main.py                    # SOLO si el gate aprueba derivar la descripción (OQ2)
├── validacion_libro.py        # esperadas derivadas de filas (extraer función)
├── hojas/
│   ├── __init__.py            # bucle sobre años + título derivado + docstring
│   ├── datos_fuente.py        # notas_por_anio.get(anio, []) — textos intactos
│   ├── semaforos.py           # prosa vía enumerar(ANIOS)
│   └── diccionario.py         # prosa vía enumerar(ANIOS)
tests/
├── test_utilidades.py         # EXTEND: enumerar()
└── test_validacion_libro.py   # NUEVO: hojas_esperadas() con subconjuntos
README.md                      # sección «Un año nuevo» — pendiente del gate (OQ1)
docs/specs/README.md           # fila 002 en estado borrador
```

Detalle por archivo:

1. **`hojas/__init__.py`** — `anios = sorted(filas, reverse=True)`; primero un
   bucle con `escribir_datos_fuente(wb, filas[a], a, len(universo))` para todos
   los años, luego un bucle con `escribir_indicadores(wb, filas[a], a)` (así se
   conserva el orden actual: DF 2026, DF 2025, Ind 2026, Ind 2025). Título:
   `f"… y semáforos {min(anios)}-{max(anios)}"`; con un solo año el sufijo es
   solo ese año (sin «2026-2026»). Docstring sin años escritos a mano.
2. **`validacion_libro.py`** — extraer función
   `hojas_esperadas(filas) -> list[str]`:
   `["Catálogo SOFIPOS"] + [f"Datos Fuente {a}"] + [f"Indicadores {a}"] + ["Diccionario", "Semáforos"]`
   con `anios = sorted(filas, reverse=True)`. El orden coincide con lo que
   escribe `escribir_xlsx` (mismo criterio, mismas llaves).
3. **`datos_fuente.py`** — `notas_por_anio = {2026: [...], 2025: [...]}` y
   `notas = notas_comunes + notas_por_anio.get(anio, [])`. **Los textos de 2026
   y 2025 quedan idénticos, byte a byte** (incluida la nota de la serie
   reexpresada que hoy se repite en ambos años: no se «deduplica», se copia tal
   cual en cada entrada). Un año sin entrada recibe solo `notas_comunes`
   (arregla el bug latente).
4. **`semaforos.py` y `diccionario.py`** — helper compartido `enumerar(valores)`
   en `sofipos/utilidades.py` (lugar propuesto; el plan puede moverlo):
   `", ".join(v[:-1]) + " e " + v[-1]` para 2+, el valor solo para 1, `""` para
   0. Cada call site pasa
   `enumerar([f"Indicadores {a}" for a in sorted(ANIOS, reverse=True)])`.
   Con la config actual produce **exactamente** «Indicadores 2026 e Indicadores
   2025»; con 3+ años, «A, B e C».
5. **`main.py`** (líneas 36-40) — ver Open Question 2.

## Code Style

Español, docstrings solo si aportan, sin relleno. Estilo esperado:

```python
def hojas_esperadas(filas: dict[int, list[dict]]) -> list[str]:
    """Hojas esperadas en orden, derivadas de los años de ``filas``."""
    anios = sorted(filas, reverse=True)
    return (
        ["Catálogo SOFIPOS"]
        + [f"Datos Fuente {a}" for a in anios]
        + [f"Indicadores {a}" for a in anios]
        + ["Diccionario", "Semáforos"]
    )
```

- Derivar siempre de `filas`/`ANIOS`, nunca escribir el año en el código.
- Los textos de contenido (notas de `datos_fuente.py`, prosa del catálogo) no se
  reescriben: solo cambia cómo se **selecciona** el que ya existe.

## Testing Strategy

Framework: **pytest** (`testpaths = ["tests"]`). La regresión del libro completo
la cubre `scripts/comparar.py`; los tests unitarios cubren los subconjuntos
arbitrarios que el libro real no puede ejercitar.

Tests nuevos:

| Archivo | Qué cubre |
|---|---|
| `tests/test_utilidades.py` (extender) | `enumerar`: 1 elemento → «Indicadores 2026»; 2 → «Indicadores 2026 e Indicadores 2025» (verbatim el texto actual); 3+ → «A, B e C»; lista vacía → `""` |
| `tests/test_validacion_libro.py` (nuevo) | `hojas_esperadas`: `{2021, 2019}` no contiguos → orden descendente exacto; `{2026}` → año único; `{2025, 2026}` → las 7 hojas actuales en el orden actual |

Seguridad de imports: `validacion_libro.py` a nivel de módulo solo importa
constantes/funciones (mismo caso ya analizado en la spec 001); `utilidades.py`
también. No se genera el libro completo en los tests.

## Boundaries

- **Always:** `comparar.py` → `IDÉNTICO` con el baseline tomado antes del
  cambio; pytest verde y `pyflakes` limpio antes de cerrar cada tarea; los
  textos de las notas salen byte a byte idénticos; estilo español.
- **Ask first:** derivar la descripción del CLI en `main.py` (OQ2); actualizar
  la sección «Un año nuevo» del README (OQ1); mover `enumerar()` a otro módulo;
  cualquier cambio en el orden de hojas.
- **Never:** agregar 2024 o notas de 2024; reescribir contenido de notas;
  `git commit` / `git push`; bajar aserciones para que pasen.

## Success Criteria

1. **`scripts/comparar.py` → `IDÉNTICO`** (exit 0), con baseline tomado antes
   del cambio (procedimiento en `AGENTS.md` § Verificar un cambio).
2. `.venv/bin/python -m pytest` → exit 0 (28 actuales + los nuevos).
3. `.venv/bin/pyflakes sofipos tests generar_sofipos_2026.py scripts` → vacío.
4. `.venv/bin/sofipos --help` → exit 0 (revisar el texto si cambia).
5. Grep sin `2025`/`2026` literales en `hojas/__init__.py`,
   `validacion_libro.py`, `semaforos.py`, `diccionario.py`. **Excepción
   documentada:** `datos_fuente.py` (`notas_por_anio` y el contenido de las
   notas — son contenido, no configuración) y, en general, textos de contenido
   (`catalogo.py`, `config.py` como fuente de verdad).
6. Los textos de 2026 y 2025 de `datos_fuente.py` quedan idénticos a los
   actuales (verificable con `comparar.py`, criterio 1).

## Out of Scope (Deferred)

- Agregar 2024 a `ANIOS` y sus notas específicas → spec 003.
- Cambios en `validacion_datos.py` (fechas `202603`/`202606` hardcodeadas) y en
  la validación de datos en general.
- `SALIDA_DEFECTO`, nombres de archivos BE (`BE_SOFIPOS_202606.xlsx`), textos
  del catálogo y docstrings con fechas de corte: son contenido, no años del libro.
- CI / mypy.
- Manejo del caso `filas`/`ANIOS` vacío (la config siempre trae ≥ 1 año).

## Open Questions

1. **README — sección «Un año nuevo» (líneas 397-403):** hoy dice que habrá que
   «agregar la llamada en `escribir_xlsx` y el nombre esperado en
   `validacion_libro.py`», que dejará de ser cierto. ¿Se actualiza en 002
   (queda «agrega el año a `ANIOS` y listo») o se deja para 003, cuando de
   verdad se agregue un año?
2. **`main.py` — descripción del CLI («2025 completo y 2026 enero-agosto»):**
   ¿se deriva de `ANIOS` ahora o se queda hasta 003? No afecta a
   `comparar.py` (solo se ve en `--help`), así que es indiferente para el
   criterio `IDÉNTICO`.

## Decisiones del gate

1. **OQ1 — sí:** la sección «Un año nuevo» del README (líneas 397-403) se
   actualiza en 002: queda «agrega el año a `ANIOS` y listo», más las notas
   específicas del año si las hay.
2. **OQ2 — sí:** la descripción del CLI en `main.py` (líneas 36-40) se deriva
   de `ANIOS` en 002. No afecta a `comparar.py` (solo se ve en `--help`); la
   aceptación incluye `sofipos --help` → exit 0 y revisión del texto.

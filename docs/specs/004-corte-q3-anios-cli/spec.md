# Spec: Corte Q3 2026, CLI `--anios` y renombres — 004

Estado: borrador

## Objective

El libro pasa al **corte del 30 de septiembre de 2026 (Q3)** y de paso se
arregla lo que hoy impide cerrar la corrida en `exit 0`. Datos disponibles:
`027_datos.csv` llega a `202609` (905 conceptos) y `ANIOS` ya tiene
`2026: range(1, 10)` (commiteado); el BE vigente es
`BE_SOFIPOS_202609.xlsx` (33 sociedades en operación).

**Hoy la corrida sale con `exit 2`** por una DIF de NICAP de 125.75 pts
(`027032` Stori, mayo 2024): el requerimiento de Stori en 202309–202405 está
anómalo (miles de pesos en vez de millones) y la dif es solo relativa
(7.5e-6). Además el README y los textos del libro siguen hablando de agosto
2026, 903 conceptos y 7 hojas.

### Requisitos aprobados (decisiones de la usuaria)

1. **Corte Q3 manteniendo 35 filas por mes en las hojas de datos.**
   `BE_UNIVERSO = BE_SOFIPOS_202609.xlsx` (33 activas). Nu México (`027014`,
   transformada 06-ago-2026) y Acción y Evolución (`027028`, revocada
   07-sep-2026) **siguen en Datos Fuente e Indicadores** con su histórico y
   `n.d.` en los meses sin reporte, hasta cerrar 2026. En la hoja
   **Catálogo SOFIPOS**: tabla principal con las 33 en operación (título y
   conteos derivados, corte «30 de septiembre de 2026») y **apartado al
   final: «Sociedades transformadas o revocadas»** (clave, motivo, fecha),
   similar a la hoja Universo del BE, para no perder la referencia. El
   apartado y las 2 bajas se alimentan de las secciones «Sociedades
   revocadas» y «Sociedades transformadas» de la hoja Universo del
   BE 202609.
2. **`BE_VALIDACION = [BE_SOFIPOS_202603.xlsx, BE_SOFIPOS_202606.xlsx,
   BE_SOFIPOS_202609.xlsx]`** (los tres), con guard `if anio in filas` para
   que un `--anios` sin 2026 no intente cruzarlos.
3. **Renombres.** `git mv generar_sofipos_2026.py
   generar_sofipos_indicadores.py`; `SALIDA_DEFECTO =
   salidas/SOFIPOS_Indicadores_Datos.xlsx`; actualizar help `--out`
   (`main.py:58`), docstrings (wrapper y `sofipos/__init__.py`), README,
   AGENTS.md (11, 34, 38) y `entradas/README.md:42`.
4. **CLI `--anios`.** Acepta `2026` (año) o `2024-2026` (rango); default =
   todos los de `ANIOS`; valida pertenencia a `ANIOS`. Función pura
   `parse_anios`, testeable. El dict seleccionado se pasa de `main` →
   `escribir_xlsx` → `escribir_diccionario`/`escribir_semaforos`,
   `validacion_datos` (consistencia fila 27) y guard BE. (`filas.py:76`,
   `datos_fuente.py:48` e `indicadores.py:85` usan `ANIOS[anio]` con `anio`
   ya contenido en `ANIOS`: no cambian.)
5. **Fix DIF NICAP: tolerancia relativa** en NICAP/IMOR/ICOR: OK si
   `dif <= 0.02 pts` **o** `dif <= 1e-5 * abs(publicado)`. Cubre Stori
   (125.75), `27016` 202411 (0.15) y `27040` 202608 (0.17), y sigue
   detectando errores reales.
6. **Notas/fuente.** Serie `201601–202609`; `notas_por_anio[2026]` reescrita
   (corte de septiembre, bajas del universo, Capital Neto/Requerimiento
   hasta agosto → NICAP `n.d.` en septiembre, serie reexpresada); notas
   comunes que mencionan `202608` y «Universo de junio de 2026»
   (`datos_fuente.py:118-122`) actualizadas.
7. **Fix `tests/test_filas.py:39`:** sin `== 16` hardcodeado; el conteo se
   deriva de `ANIOS`.
8. **README/docs con números reales:** 905 conceptos, estructura dinámica
   de hojas (3 fijas + 2 por año), 35 entidades (33 activas + 2 bajas),
   cobertura enero–septiembre, validaciones con los números de la corrida
   final, comandos renombrados y sección «Cómo actualizar» (`ANIOS`
   `range(1, 10)` y `--anios`).
9. **Verificación:** `pytest` y `pyflakes` limpios; corridas default y
   `--anios 2026`; `scripts/comparar.py` contra baseline **documentando las
   diferencias esperadas**.

### Diferencias esperadas vs. baseline (no debe quedar IDÉNTICO)

El libro **cambia a propósito**; cada diff se documenta con su causa:

- Fila de septiembre de 2026 (9 meses en vez de 8) y datos nuevos `202609`.
- Catálogo: 25 fechas de autorización cambiaron vs 202606 (ej. Klar
  2008-06-17 → 2006-09-29), título con corte 30-sep-2026 y apartado de bajas.
- Notas de Datos Fuente 2026 (corte, bajas, NICAP `n.d.`, serie).
- Renombres: ruta de salida por defecto y textos derivados.

## Tech Stack

- Python ≥ 3.11; openpyxl, pytest, pyflakes (ya en `.[dev]`).
- Sin dependencias nuevas. `git mv` para renombrar el wrapper (preserva
  historia).

## Commands

```bash
# 1. Baseline ANTES del cambio (código actual; ver AGENTS.md)
.venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida1.txt
cp /tmp/sofi.xlsx /tmp/referencia.xlsx

# 2. Tras el cambio: corridas (ambas deben salir exit 0, "Validación general: OK")
.venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
.venv/bin/python generar_sofipos_indicadores.py --anios 2026

# 3. Comparar → marcará DIFERENCIAS (exit != 0): se documenta cada diff y su causa
.venv/bin/python scripts/comparar.py \
    --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
    --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt

# Pruebas y lint
.venv/bin/python -m pytest
.venv/bin/pyflakes sofipos tests generar_sofipos_indicadores.py scripts

# CLI
.venv/bin/sofipos --help   # exit 0; revisar help de --out y --anios
```

## Project Structure

```
generar_sofipos_indicadores.py      # RENOMBRADO (git mv); docstring: corte septiembre
sofipos/
├── config.py                       # BE_UNIVERSO 202609 · BE_VALIDACION (3) · SALIDA_DEFECTO
│                                   # ANIOS SIN CAMBIOS (2026: range(1, 10))
├── main.py                         # flag --anios + parse_anios; filas/fechas desde el dict
│                                   # seleccionado; textos (descripción, help --out)
├── be.py                           # leer secciones «revocadas» y «transformadas» del BE
├── validacion_datos.py             # años seleccionados (fila 27); guard BE; tolerancia
│                                   # relativa NICAP/IMOR/ICOR; texto «marzo, junio y septiembre»
├── validacion_libro.py             # conteos derivados: catálogo 33 + apartado; por mes desde filas
└── hojas/
    ├── catalogo.py                 # título derivado (33 · 30-sep-2026) + apartado de bajas
    ├── datos_fuente.py             # notas comunes + notas_por_anio[2026]
    ├── diccionario.py              # recibe los años seleccionados (texto, no ANIOS)
    └── semaforos.py                # ídem
tests/
├── test_filas.py                   # FIX: sin == 16 hardcodeado
├── test_cli.py                     # NUEVO: parse_anios
└── test_validacion_datos.py        # NUEVO: tolerancia relativa
README.md · AGENTS.md · entradas/README.md   # comandos, conteos, fechas
```

Sin cambios: `sofipos/filas.py`, `sofipos/hojas/indicadores.py`,
`pyproject.toml` (`sofipos` y `python -m sofipos` siguen igual),
`scripts/comparar.py`, `sofipos/__main__.py`.

Detalle no obvio:

- **`validacion_libro.py`** — hoy compara el catálogo contra
  `len(universo)` (líneas 62-69) y las filas por mes contra `len(universo)`
  (131-139, con el string «35 entidades» hardcodeado). Con el corte nuevo el
  catálogo tiene 33 filas principales + apartado, y por mes siguen siendo 35:
  ambos conteos se derivan (`filas` para lo por mes), sin números escritos a
  mano.
- **`main.py`** — el universo del BE son 33; para las hojas de datos se
  agregan las 2 bajas (leídas del BE) y así `preparar_filas` produce 35
  filas/mes. El chequeo `sin_catalogo` no se afecta: `CATALOGO` ya tiene
  `027014` y `027028`.
- **`validacion_datos.py:114`** — el título «Tablero vs boletín BE (marzo y
  junio 2026…)» pasa a mencionar también septiembre.

## Code Style

Español, docstrings solo si aportan, sin relleno. Estilo esperado:

```python
def parse_anios(texto: str | None) -> dict[int, tuple[int, ...]]:
    """Subconjunto de ``ANIOS`` indicado por ``texto`` (``AAAA`` o ``AAAA-AAAA``)."""
    if texto is None:
        return ANIOS
    ...
```

```python
# Tolerancia: absoluta por redondeo o relativa al publicado (valores grandes).
ok_dif = dif <= 0.02 or dif <= 1e-5 * abs(publicado)
```

- Derivar de `ANIOS`/`filas`, nunca escribir años o conteos a mano (salvo
  textos de contenido: notas, README).
- Errores de CLI/datos: `SystemExit` con mensaje claro (patrón de `main.py`).

## Testing Strategy

Framework: **pytest** (`testpaths = ["tests"]`). La regresión del libro
completo la cubre `scripts/comparar.py` (con diferencias documentadas); los
unitarios cubren lo que la corrida real no ejercita bien.

| Archivo | Qué cubre |
|---|---|
| `tests/test_filas.py` (fix) | filas derivadas de `ANIOS[2026]` sin `== 16` hardcodeado |
| `tests/test_cli.py` (nuevo) | `parse_anios`: `None` → `ANIOS`; `2026`; `2024-2026`; año fuera de `ANIOS` → error; formato inválido → error |
| `tests/test_validacion_datos.py` (nuevo) | tolerancia relativa: dif 125.75 con publicado grande → OK; dif 0.03 con publicado ~1 → falla; ROA/ROE sigue estricto en 0.02 |

No se genera el libro completo en los tests unitarios (patrón de la spec
002).

## Boundaries

- **Always:** `pytest` verde y `pyflakes` limpio antes de cerrar cada tarea;
  estilo español; documentar cada diferencia esperada de `comparar.py` con
  su causa; números reales en el README.
- **Ask first:** cambiar cualquier umbral de validación distinto de la
  tolerancia relativa de NICAP/IMOR/ICOR; tocar `CATALOGO` o el contenido
  del catálogo más allá de lo especificado; borrar
  `salidas/SOFIPOS_2026_Datos.xlsx`.
- **Never:** excluir `027014`/`027028` de las hojas de datos en 2026;
  investigar o reexpresar los datos de Stori; reescribir las specs históricas
  (`docs/specs/001…003`); bajar aserciones para que pasen; `git commit` /
  `git push`.

## Success Criteria

1. `.venv/bin/python -m pytest` → exit 0 (fix de `test_filas` + tests nuevos).
2. `.venv/bin/pyflakes sofipos tests generar_sofipos_indicadores.py scripts`
   → vacío.
3. Corrida por defecto → exit 0 y «Validación general: OK»; escribe
   `salidas/SOFIPOS_Indicadores_Datos.xlsx`.
4. `--anios 2026` → exit 0; libro solo con Catálogo + Datos Fuente 2026 +
   Indicadores 2026 + Diccionario + Semáforos; Diccionario y Semáforos sin
   referencias a años no generados.
5. `comparar.py` vs baseline: diferencias limitadas a las esperadas y cada
   una documentada (fila septiembre, catálogo, notas, renombres).
6. Datos Fuente/Indicadores 2026: 35 entidades × 9 meses = 315 filas;
   catálogo: 33 filas principales + apartado con `027014` y `027028`;
   títulos con corte «30 de septiembre de 2026».
7. NICAP/IMOR/ICOR en «OK» en la corrida default (Stori `027032` mayo 2024
   ya no marca DIF); ROA/ROE mantiene 0.02 estricto.
8. Los tres BE se validan en la corrida default; con `--anios` sin 2026 no
   se intenta cruzar BE de 2026 (guard).
9. Grep sin `generar_sofipos_2026` ni `SOFIPOS_2026_Datos` en `README.md`,
   `AGENTS.md`, `entradas/README.md` y código fuente (excluye
   `docs/specs/001-003`).
10. README: 905 conceptos y serie `201601–202609`; hojas «3 fijas + 2 por
    año»; 35 entidades (33 activas + 2 bajas); cobertura enero–septiembre;
    validaciones con números de la corrida final; comandos nuevos; sección
    «Cómo actualizar» con `range(1, 10)` y `--anios`.

## Out of Scope (Deferred)

- Excluir Nu México y A&E de las hojas de datos (enero 2027).
- Investigar o reexpresar el requerimiento anómalo de Stori 202309–202405
  (solo se tolera la dif).
- Tolerancias para DIF de otros indicadores que no sean NICAP/IMOR/ICOR.
- Cambiar umbrales de otras validaciones (consistencia fila 27, cartera,
  ROA/ROE): se mantienen como están.
- CI / mypy.

## Open Questions

1. **BE 202609 contra el tablero:** la comparación de 7 variables con el BE
   de septiembre aún no se ha corrido (solo 202603/202606 están hoy en
   `BE_VALIDACION`). ¿Coincide (como junio) o trae diferencias nuevas? Si
   trae, ¿se documenta excepción como la de SAE o se investiga? No se cambia
   ningún umbral sin gate.
2. **Salida vieja:** ¿se borra `salidas/SOFIPOS_2026_Datos.xlsx` tras el
   renombre o se deja en su lugar?
3. **Apartado del catálogo:** además de clave, motivo y fecha, ¿se incluye
   el nombre de la sociedad (como en la hoja Universo del BE)?
4. **Stdout «Universo: N sociedades»:** ¿33 (las en operación del BE) o 35
   (con las bajas)? Afecta solo al baseline de stdout de `comparar.py`.

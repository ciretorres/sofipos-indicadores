# Plan: Años desde una única fuente — 002

Spec: `docs/specs/002-anios-unico-fuente/spec.md` (aprobada; OQ1 y OQ2
resueltos en el gate — ver «Decisiones del gate» al final de la spec).

## Enfoque

Refactor puro: derivar los años de `filas`/`ANIOS` en los 5 puntos de la spec
+ la descripción del CLI (OQ2) + el README (OQ1). Cero cambios de contenido:
el libro debe quedar `IDÉNTICO` contra el baseline tomado **antes** de tocar
nada.

Orden: baseline → código por archivo → docs → gate final. Cada tarea deja el
repo verde (pytest + pyflakes + `comparar.py`) antes de pasar a la siguiente.

Orden de hojas conservado idéntico (lo verifica `comparar.py` y el test de
`hojas_esperadas`): `Catálogo SOFIPOS`, `Datos Fuente 2026`, `Datos Fuente
2025`, `Indicadores 2026`, `Indicadores 2025`, `Diccionario`, `Semáforos`.

## Detalle por archivo

### 1. `sofipos/hojas/__init__.py` (T002)

- `anios = sorted(filas, reverse=True)` al inicio de `escribir_xlsx`.
- Sustituir las 4 llamadas de la líneas 41-44 por **dos bucles** (así se
  conserva el orden actual DF 2026, DF 2025, Ind 2026, Ind 2025; un bucle
  interleaved daría otro orden):

  ```python
  escribir_catalogo(wb, universo, nombres)
  for a in anios:
      escribir_datos_fuente(wb, filas[a], a, len(universo))
  for a in anios:
      escribir_indicadores(wb, filas[a], a)
  escribir_diccionario(wb)
  escribir_semaforos(wb)
  ```

- Título (líneas 49-52) derivado, con caso de año único sin «2026-2026»:

  ```python
  rango = f"{min(anios)}-{max(anios)}" if len(anios) > 1 else str(anios[0])
  wb.properties.title = (
      "SOFIPOS - Catálogo, datos fuente, indicadores, diccionario "
      f"y semáforos {rango}"
  )
  ```

  Con la config actual sale byte a byte el título actual
  («… y semáforos 2025-2026»).
- Docstring (línea 27): quitar «para 2025 y 2026» (los docstrings no entran
  al libro; solo hay que dejar de escribir años a mano).

### 2. `sofipos/validacion_libro.py` (T003)

- Extraer a nivel de módulo `hojas_esperadas(filas) -> list[str]` (forma
  exacta en «Code Style» de la spec) y reemplazar la lista `esperadas` de las
  líneas 41-49 por `esperadas = hojas_esperadas(filas)`.
- Las validaciones por año (líneas 79+ y 210-221) ya son genéricas: sin
  cambios.
- Test nuevo `tests/test_validacion_libro.py`: subconjuntos `{2019, 2021}`
  (no contiguos → orden descendente exacto), `{2026}` (año único → 5 hojas),
  `{2025, 2026}` (las 7 hojas actuales en el orden actual).

### 3. `sofipos/hojas/datos_fuente.py` (T004)

- Líneas 140-153: pasar a

  ```python
  notas_por_anio = {2026: [...], 2025: [...]}   # solo notas específicas
  notas = notas_comunes + notas_por_anio.get(anio, [])
  ```

- **Los textos de 2026 y 2025 quedan byte a byte idénticos**: solo cambia
  cómo se selecciona el que ya existe (incluida la nota de la serie
  reexpresada que se repite en ambos años: se copia tal cual, no se
  deduplica). Un año sin entrada recibe solo `notas_comunes` (arregla el bug
  latente).

### 4. `sofipos/utilidades.py` + `semaforos.py` + `diccionario.py` (T005)

- Helper nuevo en `utilidades.py`:

  ```python
  def enumerar(valores) -> str:
      """Une valores con comas y «e» antes del último (2 o más)."""
      v = list(valores)
      if not v:
          return ""
      if len(v) == 1:
          return v[0]
      return ", ".join(v[:-1]) + " e " + v[-1]
  ```

- `semaforos.py` (líneas 205-206) y `diccionario.py` (línea 38): importar
  `ANIOS` de `..config` y `enumerar` de `..utilidades`; la prosa pasa a

  ```python
  enumerar([f"Indicadores {a}" for a in sorted(ANIOS, reverse=True)])
  ```

  Con la config actual produce exactamente «Indicadores 2026 e Indicadores
  2025». Estos dos módulos no reciben `filas`, por eso se importa `ANIOS`
  (mismo patrón que `indicadores.py` y `datos_fuente.py`, que ya lo hacen).

### 5. `sofipos/main.py` (T006, OQ2 aprobado)

- Líneas 36-40: descripción derivada de `ANIOS`. Propuesta concreta:

  ```python
  parser = argparse.ArgumentParser(
      description=(
          "Genera el Excel mensual de las 35 SOFIPOS "
          f"({', '.join(str(a) for a in sorted(ANIOS))})."
      )
  )
  ```

  Con la config actual: «Genera el Excel mensual de las 35 SOFIPOS (2025,
  2026).» Se pierde el detalle «2025 completo y 2026 enero-agosto»: derivarlo
  exigiría formatear rangos de meses y sería inexacto con años no contiguos.
  El texto se revisa en la aceptación de T006 (así lo fija el gate).

### 6. `README.md` (T006, OQ1 aprobado)

- Sección **«Un año nuevo» (líneas 397-403)**: se reescribe — «agrega el año
  a `ANIOS` y listo»: `escribir_xlsx`, la validación y los textos ya se
  derivan; si el año tiene notas específicas, agrégalas a `notas_por_anio` en
  `sofipos/hojas/datos_fuente.py` (si no, el año corre solo con las notas
  comunes).
- Revisado: «Un mes nuevo» (385-395) no describe el proceso manual de años
  (solo agregar el mes a `ANIOS`) → sin cambios. Líneas 40, 46 y 66 son
  descripción de estructura del repo → sin cambios.

## Riesgos y mitigación

| Riesgo | Impacto | Mitigación |
|---|---|---|
| `enumerar()` no reproduce el texto actual byte a byte | Alto: rompe `IDÉNTICO` | Test con el texto verbatim (T005) + `comparar.py` tras la tarea |
| Título mal derivado en el caso de año único («2026-2026») | Medio: `comparar.py` no lo cubre (config actual tiene 2 años) | Caso `len(anios) == 1` explícito; revisión en T002 |
| Ciclo de imports al importar `ANIOS` en `semaforos.py`/`diccionario.py` | Medio | `config.py` no importa nada del paquete (solo `pathlib`); patrón ya en uso; pytest importa ambos módulos |
| Al reestructurar `notas` cambia un byte del texto | Alto | `comparar.py` (criterio 6 de la spec); no reescribir textos |
| El texto de `--help` cambia | Bajo | Aprobado en OQ2; `--help` exit 0 + revisión visual; no entra a `comparar.py` |
| Baseline desactualizado si el repo cambia a mitad de 002 | Alto | T001 toma el baseline antes de todo; repetirlo si algo cambia |

## Verificación final (T007)

```bash
# Pruebas y lint
.venv/bin/python -m pytest
.venv/bin/pyflakes sofipos tests generar_sofipos_2026.py scripts

# CLI
.venv/bin/sofipos --help   # exit 0; revisar el texto

# Regresión del libro (baseline de T001) → IDÉNTICO, exit 0
.venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
.venv/bin/python scripts/comparar.py \
    --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
    --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt

# Sin años literales en los 4 archivos refactorizados (sin salida = OK)
grep -n '2025\|2026' sofipos/hojas/__init__.py sofipos/validacion_libro.py \
    sofipos/hojas/semaforos.py sofipos/hojas/diccionario.py
```

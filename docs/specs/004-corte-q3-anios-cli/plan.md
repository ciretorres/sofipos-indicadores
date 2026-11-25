# Plan: Corte Q3 2026, CLI `--anios` y renombres — 004

Spec: [spec.md](spec.md) (contrato aprobado). Tareas: [task.md](task.md).

## Resumen

El libro pasa al corte del 30 de septiembre de 2026 (BE `202609`, 33 activas +
2 bajas = 35 filas por mes), se renombra el entry point y la salida por
defecto, se arregla la DIF de NICAP que hoy impide el `exit 0`, se agrega la
CLI `--anios` y la documentación se actualiza con números reales. El libro
**cambia a propósito**: la verificación final documenta cada diferencia vs el
baseline en lugar de pedir `IDÉNTICO`.

## Decisiones de la usuaria (open questions de la spec, resueltas)

1. **BE 202609 vs tablero (primera corrida):** si algo no cuadra, **sí se
   atiende** — investigar la causa y, si es un caso real, documentarlo como
   diferencia esperada (patrón de la excepción SAE). No ignorarlo; ningún
   umbral se cambia sin gate.
2. **Salida vieja:** se **deja** `salidas/SOFIPOS_2026_Datos.xlsx` (no se
   borra). Borrarla queda como *Ask first*.
3. **Apartado del catálogo:** incluye el **nombre** → clave, nombre, motivo,
   fecha.
4. **Stdout:** «Universo: 33 activas · +2 bajas» (texto equivalente claro
   vale). Afecta al baseline de stdout de `comparar.py`.

## Decisiones técnicas

- **Dos listas en `main`:** `universo` (33, del BE) alimenta el catálogo y la
  validación del catálogo; `universo_datos = universo + bajas` (35) alimenta
  `preparar_filas` y `validar_datos` (cobertura). `escribir_xlsx` recibe
  `universo` y `bajas` por separado y deriva `n_entidades = len(universo) +
  len(bajas)` (35) para las hojas de datos.
  `CATALOGO` ya tiene `027014` y `027028`; el chequeo `sin_catalogo` no cambia.
- **Bajas derivadas, sin lista hardcodeada:** `be.leer_bajas(ruta)` lee las
  secciones «Sociedades revocadas» y «Sociedades transformadas» de la hoja
  Universo del BE y filtra por pertenencia a `CATALOGO` → exactamente `027014`
  (transformada) y `027028` (revocada), con clave, nombre, motivo y fecha.
- **Consecutivo de las bajas:** 34 y 35 (`len(universo) + i`) para la columna
  «No.» de las hojas de datos (los consecutivos de las secciones del BE se
  repiten con los activos). El apartado no muestra consecutivo.
  Nombre del BE **sin el marcador de nota al pie** («Nu México 7/» →
  «Nu México»); motivo = columna «Observaciones» (con fallback «Transformada»
  / «Revocada») y fecha = la columna F del BE (fecha de transformación o de
  autorización, igual que la hoja Universo).
- **Conteos derivados, nunca a mano:** catálogo vs `len(universo)` (33);
  filas por mes desde `filas` (claves distintas del año); desaparece el string
  «35 entidades» de `validacion_libro.py:138`.
- **Tolerancia relativa solo en NICAP/IMOR/ICOR** (`validacion_datos.py`
  bloque de la línea ~301): `dif <= 0.02` **o** `dif <= 1e-5 * abs(publicado)`,
  evaluado **por comparación** (no sobre el `peor` agregado: una dif grande
  tolerada no puede ocultar una dif chica con publicado pequeño). ROA/ROE
  (línea ~231) y cartera siguen estrictos en 0.02.
- **`--anios`:** `parse_anios` puro en `main.py` (`None` → `ANIOS`; `2026`;
  `2024-2026`; fuera de `ANIOS` o formato inválido → `SystemExit` con mensaje).
  El dict seleccionado se pasa de `main` → `escribir_xlsx` →
  `escribir_diccionario`/`escribir_semaforos` y a `validar_datos` (consistencia
  fila 27 + guard `if anio not in filas` para no cruzar BE de 2026). También
  `main` usa el dict para fechas y `filas`. `filas.py:76`, `datos_fuente.py:48`
  e `indicadores.py:85` usan `ANIOS[anio]` con `anio` ya contenido en `ANIOS`:
  no cambian.
- **Baseline y la «fila de septiembre»:** el baseline se genera con el código
  actual, que **ya** trae `ANIOS = {2026: range(1, 10)}` → ya incluye 9 meses.
  Por eso la diferencia «fila de septiembre (9 meses en vez de 8)» de la spec
  se observa frente al **libro publicado anterior**
  (`salidas/SOFIPOS_2026_Datos.xlsx`, 8 meses), no frente al baseline fresco.
  Si `comparar.py` reporta diffs en filas de 2026 (p. ej. por renumeración de
  la columna «No.»), se documentan con su causa.

## Orden de implementación (verticales verificables)

Detalle y comandos en `task.md` (T001…T011).

| Fase | Tareas | Qué entrega | Checkpoint |
|---|---|---|---|
| 1 · Base | T001–T003 | Baseline con el nombre viejo; `pytest` verde (fix `test_filas`); tolerancia NICAP/IMOR/ICOR → corrida default `exit 0` | **CP1** |
| 2 · Renombres | T004 | `git mv` + `SALIDA_DEFECTO` + help/docstrings; salida vieja intacta | grep de nombres viejos en código limpio |
| 3 · Corte Q3 | T005–T007 | BE `202609` + validación de los 3 boletines; bajas → 35 filas/mes y apartado del catálogo; notas reescritas | **CP2** |
| 4 · CLI | T008–T009 | `parse_anios` + guard/bucle por `filas` en `validacion_datos` + thread del dict por main y hojas | **CP3** |
| 5 · Docs | T010 | README/AGENTS/entradas + docstrings con números reales | **CP4** |
| 6 · Cierre | T011 | Verificación global y diferencias documentadas | **gate final** |

Dependencias: T001 primero (el baseline exige el nombre viejo); T004 antes de
T005 (los comandos ya usan el nombre nuevo); T005 → T006 → T007 (el corte se
construye de afuera hacia adentro: config → main/hojas → notas); T008 → T009
(función antes que thread); T010 después de todo el código; T011 al final.
T002 y T003 son independientes entre sí.

Desde T003 la corrida queda en `exit 0` (T001/T002 conservan el `exit 2`
del baseline). Los estados intermedios se acotan en `task.md`: T005 deja el
catálogo con 33 filas y título viejo hasta T006.

## Riesgos y mitigaciones

| Riesgo | Impacto | Mitigación |
|---|---|---|
| `--anios` a medias: queda un `ANIOS` importado directo (`diccionario`, `semaforos`, `validacion_datos`) | Alto — el flag miente en silencio | Al cerrar T009, grep `from .*config import ANIOS` y revisar cada uso; corridas `--anios 2026` y `--anios 2025` verificando hojas y textos |
| Conteos 33+2 vs 35 (catálogo vs datos) | Alto — validación roja o libro incompleto | Derivar siempre (`len(universo)`, `filas`); revisar mensajes de `validacion_libro` en CP2 |
| Reexpresión del catálogo (25 fechas, título, apartado, fuente 202609) | Medio — se toma por bug | Son diffs esperadas de la spec; documentarlas, no «arreglar» fechas |
| Primera corrida con BE 202609 trae DIF nuevas | Medio — bloquea el `exit 0` o se ignora | Respuesta 1: investigar; caso real → documentar (y si aplica, excepción tipo SAE con gate). No cambiar umbrales |
| Tolerancia agregada vs por comparación | Medio — deja pasar errores reales | Evaluar por comparación (ver decisión técnica); test con dif 0.03 y publicado ~1 que debe fallar |
| Stdout cambia («Universo: 33 activas · +2 bajas», bloque BE septiembre, NICAP OK) | Bajo — `comparar.py` sale ≠ 0 | Diffs esperadas; documentarlas en T011 |
| Borrar la salida vieja por error | Bajo | Respuesta 2: no borrar; *Ask first* |

## Checkpoints de verificación (resumen)

```bash
# CP1 (tras T003; el renombre aún no existe → nombre viejo):
.venv/bin/python -m pytest                       # exit 0
.venv/bin/pyflakes sofipos tests generar_sofipos_2026.py scripts   # vacío
.venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx      # exit 0, "Validación general: OK"

# CP2 (tras T007) y CP3 (tras T009): ídem con el nombre nuevo
.venv/bin/python -m pytest
.venv/bin/pyflakes sofipos tests generar_sofipos_indicadores.py scripts
.venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx   # exit 0, "Validación general: OK"
.venv/bin/python generar_sofipos_indicadores.py --anios 2026 --out /tmp/sofi_2026.xlsx  # solo hojas 2026

# CP4 (tras T010): sin nombres viejos en docs
grep -rn "generar_sofipos_2026\|SOFIPOS_2026_Datos" README.md AGENTS.md entradas/README.md   # vacío
```

El detalle por tarea (con `--anios 2025`, apartado, notas, etc.) está en
`task.md`. La verificación final (T001 baseline → `comparar.py` con diffs
documentadas, grep completo, README) también está ahí.

## Open Questions

Resueltas (de la spec): las 4 decisiones de la usuaria, arriba.

Pendiente (menor, para el gate):

1. **Fecha del apartado de bajas:** se asume la columna F del BE (fecha de
   transformación para Nu México, fecha de autorización para Acción y
   Evolución), igual que la hoja Universo. Si se prefiere la fecha de
   revocación/transformación real (en el texto de «Observaciones»), es un
   cambio de una línea — pero requiere parsear texto del BE.

El único resultado desconocido es la primera corrida BE 202609 vs tablero
(respuesta 1): se investiga en T005 y se documenta en T011.

Sin `git commit` / `git push` (AGENTS.md): se deja el cambio listo y se
sugieren los comandos.

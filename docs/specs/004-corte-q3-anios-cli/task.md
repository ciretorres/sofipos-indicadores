# Tareas: Corte Q3 2026, CLI `--anios` y renombres — 004

Orden: T001 → T011 (T001 es prerequisito de todas; T011 cierra el gate).
Marca `[x]` solo cuando la verificación de esa tarea pasa.

Notas:

- El libro **cambia a propósito**: desde T003 `comparar.py` deja de dar
  `IDÉNTICO`. Cada tarea indica qué diffs esperar; la lista final, con causa
  de cada uno, se documenta en T011.
- **Hasta T004** se corre con `generar_sofipos_2026.py`; **desde T004**, con
  `generar_sofipos_indicadores.py`.
- Tras cada tarea que toca código, repite los pasos 2-3 de AGENTS.md (con el
  nombre de entry point que corresponda según la nota anterior):

```bash
.venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
.venv/bin/python scripts/comparar.py \
    --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
    --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt
```

- Sin `git commit` / `git push`: se deja el cambio listo y se sugieren los
  comandos (AGENTS.md).

## Tareas

- [x] **T001 — Baseline antes de tocar nada**
  - Archivos: ninguno (solo artefactos en `/tmp`).
  - Aceptación: existen `/tmp/referencia.xlsx` y `/tmp/salida1.txt`
    generados con el código actual; la corrida **sale con exit 2** y
    `salida1.txt` termina en «Validación general: CON OBSERVACIONES» (es el
    estado actual: DIF de NICAP Stori, no un error de la tarea).
  - Verificar:
    ```bash
    .venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida1.txt
    echo "exit: $?"          # 2 (esperado)
    cp /tmp/sofi.xlsx /tmp/referencia.xlsx
    test -s /tmp/referencia.xlsx && grep -q "CON OBSERVACIONES" /tmp/salida1.txt && echo OK
    ```

- [x] **T002 — `tests/test_filas.py`: sin `== 16` hardcodeado**

  - Hallazgo: el fix ya estaba aplicado en el working tree (sesión previa);
    verificado con `pytest --ignore=tests/test_validacion_datos.py` (35
    passed) y `pyflakes tests/test_filas.py` limpio. El único rojo de la
    suite completa era el import de `_tolera_dif` en `tests/test_validacion_datos.py`,
    que corresponde a T003 (estado rojo del TDD, verificado ahí).
  - Archivos: `tests/test_filas.py`.
  - Aceptación:
    - Línea 39: `assert len(filas) == len(ANIOS[2026]) * len(UNIVERSO)`
      (sin `== 16`); docstring «8 meses × 2 sociedades = 16» actualizada al
      conteo derivado.
    - `pytest` completo en verde (el rojo actual es
      `test_una_fila_por_sociedad_y_mes`; si aparece otro rojo, investigarlo
      aquí).
  - Verificar:
    ```bash
    .venv/bin/python -m pytest
    .venv/bin/pyflakes tests/test_filas.py
    ```
  - comparar: `IDÉNTICO` (solo cambió un test).

- [x] **T003 — Tolerancia relativa NICAP/IMOR/ICOR**
  - Archivos: `sofipos/validacion_datos.py`, `tests/test_validacion_datos.py`
    (nuevo).
  - Aceptación:
    - Helper puro `_tolera_dif(dif, publicado)` →
      `dif <= 0.02 or dif <= 1e-5 * abs(publicado)`, evaluado **por
      comparación** en el bloque NICAP/IMOR/ICOR (~línea 301): solo las difs
      que no pasan la tolerancia cuentan como falla y se reporta la peor de
      esas (una dif grande tolerada no puede ocultar una dif chica con
      publicado pequeño).
    - ROA/ROE (~línea 231), cartera y bloque BE (~línea 157) siguen estrictos
      en 0.02: sin cambios.
    - Tests nuevos: dif 125.75 con publicado grande (Stori `027032` mayo
      2024, `nicap_pub` ≈ 16,772,979) → pasa; dif 0.03 con publicado ~1 →
      falla; dif 0.015 → pasa.
    - Corrida por defecto: **exit 0** y «Validación general: OK».
  - Verificar:
    ```bash
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos/validacion_datos.py tests/test_validacion_datos.py
    .venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
    echo "exit: $?"          # 0
    grep -q "Validación general: OK" /tmp/salida2.txt && echo OK
    ```
  - comparar: xlsx `IDÉNTICO`; stdout con diffs esperados (NICAP/IMOR/ICOR
    `DIF` → `OK`; «CON OBSERVACIONES» → «OK»).

- [x] **T004 — Renombres (entry point y salida por defecto)**

  - Hallazgo: `grep` en `sofipos tests scripts *.py` solo matchea el `.pyc`
    de `sofipos/__pycache__` (no es fuente; se ignora). La corrida con `--out`
    explícito da IDÉNTICO al estado post-T003.
  - Archivos: `generar_sofipos_2026.py` → `generar_sofipos_indicadores.py`
    (`git mv`), `sofipos/config.py`, `sofipos/main.py`,
    `sofipos/__init__.py`.
  - Aceptación:
    - `git mv` hecho (preserva historia).
    - `SALIDA_DEFECTO = salidas/SOFIPOS_Indicadores_Datos.xlsx`; help de
      `--out` (`main.py:58`) cita el nombre nuevo.
    - Docstrings del wrapper (incluye el corte de septiembre de 2026) y de
      `sofipos/__init__.py` sin `generar_sofipos_2026` ni `SOFIPOS_2026_Datos`.
    - `salidas/SOFIPOS_2026_Datos.xlsx` **intacta** (no se borra).
    - `grep -rn "generar_sofipos_2026\|SOFIPOS_2026_Datos" sofipos tests
      scripts *.py` → sin salida.
  - Verificar:
    ```bash
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos tests generar_sofipos_indicadores.py scripts
    .venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
    echo "exit: $?"          # 0
    .venv/bin/sofipos --help # exit 0; revisar help de --out
    ```
  - comparar: `IDÉNTICO` (siempre salimos con `--out` explícito).

- [x] **T005 — BE 202609 + validación de los tres boletines**

  - Hallazgos:
    - **OQ1 (BE 202609 vs tablero): resuelta sin diferencias nuevas.** El
      bloque «Septiembre 2026» sale todo OK (peor dif ≈ 0.005 en las 7
      variables); no hizo falta ninguna excepción ni cambio de umbral.
    - Corrida exit 0; stdout con «Universo: 33 sociedades ·
      BE_SOFIPOS_202609.xlsx» y bloque «Septiembre 2026 (corte
      definitivo)».
    - comparar vs baseline: diffs esperadas — Catálogo con las 33 filas del
      BE 202609 (fechas y orden distintos: 25 fechas cambiaron vs 202606,
      p. ej. Klar 2008-06-17 → 2006-09-29; Nu México y A&E fuera de la
      tabla principal), hojas de datos 2024/2025/2026 bajan de 35 a 33
      filas/mes, título/fuente del catálogo aún con corte junio (queda
      transitorio hasta T006/T007).
  - Archivos: `sofipos/config.py`, `sofipos/validacion_datos.py`.
  - Aceptación:
    - `BE_UNIVERSO = "BE_SOFIPOS_202609.xlsx"` (33 en operación);
      `BE_VALIDACION = [202603, 202606, 202609]` (los tres).
    - Texto de la línea 114: «Tablero vs boletín BE (marzo, junio y
      septiembre de 2026, miles de pesos)».
    - Corrida por defecto exit 0; el stdout incluye el bloque
      «Septiembre 2026».
    - **Primera corrida BE 202609 vs tablero (OQ1):** revisar lo que imprime
      el bloque de septiembre; si algo no cuadra, investigar la causa y
      documentarlo. No se cambia ningún umbral ni se agrega excepción sin
      gate.
  - Verificar:
    ```bash
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos/config.py sofipos/validacion_datos.py
    .venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
    echo "exit: $?"                                            # 0
    grep -q "Septiembre 2026" /tmp/salida2.txt && echo OK
    ```
  - comparar: diffs esperados (transitorios hasta T006): catálogo con las
    fechas del BE 202609 (25 cambiaron vs 202606); Datos Fuente/Indicadores
    2026 bajan a 33 filas/mes (el universo de junio tenía 35 en operación);
    stdout con «Universo: 33 sociedades · BE_SOFIPOS_202609.xlsx» y bloque de
    septiembre.

- [x] **T006 — Bajas: 35 filas/mes y apartado del catálogo**

  - Hallazgos:
    - `leer_bajas` devuelve exactamente `027028` (revocada, motivo =
      Observaciones del BE) y `027014` (transformada, motivo =
      «Transformada» por no traer Observaciones), con nombre sin el marcador
      de nota al pie («Nu México 7/» → «Nu México») y fecha de la columna F
      (A&E: fecha de autorización 17-jun-2008; Nu México: fecha de
      transformación 06-ago-2026 — coincide con el open question pendiente
      del plan).
    - Corrida exit 0; stdout «Universo: 33 activas · +2 bajas ·
      BE_SOFIPOS_202609.xlsx»; catálogo 33 filas principales (filas 5-37) +
      apartado en filas 39-42; Datos Fuente/Indicadores 2026: 315 filas
      (35 × 9); 2024/2025: 420 (35 × 12); «todas con 35» derivado de
      `filas` (sin el string hardcodeado).
    - comparar vs baseline (14,962 diffs de xlsx + 139 de stdout):
      - `Catálogo SOFIPOS` (228): 33 filas vs 35 (bajas fuera de la tabla
        principal), reorden por el BE 202609 y sus 25 fechas distintas,
        título/fuente con corte 30-sep-2026, apartado nuevo con 027014 y
        027028, notas actualizadas.
      - `Datos Fuente 2024/2025/2026` + `Indicadores …` (≈14.7k): se
        reordenan las filas y cambian la columna «No.» (consecutivos de las
        bajas pasan a 34 y 35 al final) y el nombre de las bajas en las
        hojas de datos; el número de filas por mes sigue siendo 35.
      - stdout: «33 activas · +2 bajas», bloque «Septiembre 2026», NICAP
        DIF→OK, «todas con 35» y «Validación general: OK».
  - Archivos: `sofipos/be.py`, `sofipos/main.py`,
    `sofipos/hojas/__init__.py`, `sofipos/hojas/catalogo.py`,
    `sofipos/validacion_libro.py` (5 archivos).
  - Aceptación:
    - `be.leer_bajas(ruta)`: lee las secciones «Sociedades revocadas…» y
      «Sociedades transformadas» de la hoja Universo del BE, filtra por
      pertenencia a `CATALOGO` → exactamente `027014` (transformada) y
      `027028` (revocada), con clave, nombre (sin marcador de nota al pie),
      motivo («Observaciones», con fallback «Transformada»/«Revocada») y
      fecha (columna F del BE).
    - `main`: `bajas = leer_bajas(...)`; `universo_datos = universo + bajas`
      con consecutivos **34 y 35**; stdout «Universo: 33 activas · +2
      bajas»; `preparar_filas` y `validar_datos` reciben `universo_datos`;
      `escribir_xlsx(universo, bajas, ...)`; `validar_libro` recibe
      `universo` (33).
    - `hojas/__init__.py`: `n_entidades = len(universo) + len(bajas)` (35) y
      `bajas` se pasa a `escribir_catalogo`.
    - `catalogo.py`: título derivado «33 sociedades autorizadas en operación ·
      Información al 30 de septiembre de 2026» + fuente
      `BE_SOFIPOS_202609.xlsx`; apartado «Sociedades transformadas o
      revocadas» (clave, nombre, motivo, fecha) al final; notas de fuente
      actualizadas.
    - `validacion_libro.py`: filas por mes derivadas de `filas` (35); sin el
      string «35 entidades» hardcodeado; catálogo sigue comparándose contra
      `len(universo)` (33).
    - Corrida exit 0; catálogo: 33 filas principales + apartado con
      `027014` y `027028`; Datos Fuente/Indicadores 2026: 35 × 9 = 315
      filas.
  - Verificar:
    ```bash
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos
    .venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
    echo "exit: $?"          # 0
    grep -q "33 activas · +2 bajas" /tmp/salida2.txt && echo OK
    .venv/bin/python -c "
    from openpyxl import load_workbook
    ws = load_workbook('/tmp/sofi.xlsx')['Catálogo SOFIPOS']
    for r in ws.iter_rows(min_row=38, max_row=46, values_only=True):
        if any(v is not None for v in r): print(r)"   # ver apartado con 027014 y 027028
    ```
  - comparar: diffs esperados: 2 filas devueltas por mes (35), apartado
    nuevo, título/fuente del catálogo, stdout con activas/bajas.

- [x] **T007 — Notas y fuente de Datos Fuente**

  - Hallazgos:
    - Nota FUENTE con serie `201601–202609`; nota de universo reescrita (33
      en operación + 2 bajas = 35 filas/mes, CAME fuera); `notas_por_anio[2026]`
      reescrita (corte de septiembre, bajas del universo, Capital
      Neto/Requerimiento no publicados en septiembre → NICAP `n.d.`, serie
      reexpresada).
    - Verificado en el tablero: `capital_neto`/`requerimiento` tienen 0
      registros para `202609` (33 para `202608`), así que la nota es
      correcta.
    - `grep -n "202608\|junio de 2026" sofipos/hojas/datos_fuente.py` → sin
      hits (los hits restantes en `sofipos/` son `BE_SOFIPOS_202606` en
      `BE_VALIDACION` y la excepción SAE, correctos).
    - comparar: diffs incrementales vs T006 solo en los bloques de notas de
      las 3 hojas de datos (+3/+2/+2 celdas).
  - Archivos: `sofipos/hojas/datos_fuente.py`.
  - Aceptación:
    - Nota FUENTE: serie `201601–202609` (no `202608`).
    - Nota de universo reescrita: 33 «en operación» en el Universo de
      septiembre de 2026 + 2 bajas (Nu México, Acción y Evolución) = 35
      filas/mes; CAME sigue fuera.
    - `notas_por_anio[2026]` reescrita: corte de septiembre, bajas del
      universo, Capital Neto/Requerimiento hasta agosto → NICAP `n.d.` en
      septiembre, serie reexpresada.
    - Sin textos de «junio de 2026» ni `202608` que ya no apliquen
      (revisar cada hit de `grep -n "202608\|junio de 2026"
      sofipos/hojas/datos_fuente.py`).
  - Verificar:
    ```bash
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos/hojas/datos_fuente.py
    .venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
    echo "exit: $?"          # 0
    .venv/bin/python -c "
    from openpyxl import load_workbook
    ws = load_workbook('/tmp/sofi.xlsx')['Datos Fuente 2026']
    for r in ws.iter_rows(min_row=321, values_only=True):
        for v in r:
            if isinstance(v, str) and v: print(v)"   # notas: revisar 201601–202609 y septiembre
    ```
  - comparar: diffs esperados: solo los bloques de notas (las 3 hojas de
    datos, por la nota común).

- [x] **T008 — `parse_anios` + validación lista para el subconjunto + tests**

  - Cierre: implementado en el working tree de una sesión previa; verificado
    aquí con `pytest` (45 passed, incluido `tests/test_cli.py` con los 5 casos)
    y `pyflakes` limpio. La función aún no se hilvanaba a la CLI (eso es T009),
    así que la corrida no cambia.
  - Archivos: `sofipos/main.py`, `sofipos/validacion_datos.py`,
    `tests/test_cli.py` (nuevo).
  - Aceptación:
    - `parse_anios(texto: str | None) -> dict[int, tuple[int, ...]]` en
      `main.py`: `None` → `ANIOS`; `"2026"` → solo 2026; `"2024-2026"` →
      los tres años; año fuera de `ANIOS` o formato inválido → `SystemExit`
      con mensaje claro (patrón de `main.py`).
    - Cada año del rango debe pertenecer a `ANIOS` (si no, `SystemExit`).
    - `validacion_datos.py`: consistencia fila 27 itera `filas` (líneas
      89-90, ya no `ANIOS`); guard `if anio not in filas: continue` antes de
      cruzar cada BE; sin import de `ANIOS`. **Sin efecto en la corrida
      default** (filas contiene todos los años).
    - `tests/test_cli.py` cubre los 5 casos de la spec.
    - La función **aún no se hilvana** a la CLI: la corrida no cambia.
  - Verificar:
    ```bash
    .venv/bin/python -m pytest tests/test_cli.py
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos/main.py sofipos/validacion_datos.py tests/test_cli.py
    ```
  - comparar: `IDÉNTICO`.

- [x] **T009 — Hilvanado de `--anios` por main y las hojas**

  - Hallazgos:
    - Comprobado con una copia del paquete con los cambios de T009 revertidos
      (`/tmp/pre009`): el `.xlsx` es **IDÉNTICO** al actual (solo cambian las
      líneas de ruta del stdout), así que el hilvanado no altera el libro por
      defecto.
    - `--anios 2026` → Catálogo + Datos/Indicadores 2026 + Diccionario +
      Semáforos (5 hojas); `--anios 2025` → ídem con 2025 (guard BE activo, no
      cruza boletines de 2026); `--anios 2027` → exit 1 con «año no
      contemplado 2027 (disponibles: 2024, 2025, 2026)».
    - Diccionario y Semáforos citan solo los años generados (revisado en los
      dos libros).
  - Archivos: `sofipos/main.py`, `sofipos/hojas/__init__.py`,
    `sofipos/hojas/diccionario.py`, `sofipos/hojas/semaforos.py`,
    `tests/test_semaforos.py` (5 archivos).
  - Aceptación:
    - `main`: flag `--anios` (help «AAAA o AAAA-AAAA»); `anios_sel =
      parse_anios(args.anios)`; descripción del parser, bucle de `fechas` y
      dict `filas` usan `anios_sel`.
    - `hojas/__init__.py`: pasa `anios = sorted(filas, reverse=True)` (ya
      derivado) a `escribir_diccionario` y `escribir_semaforos`.
    - `diccionario.py` y `semaforos.py`: reciben `anios` como parámetro;
      sin `ANIOS`.
    - `tests/test_semaforos.py`: la llamada pasa los años explícitos
      (`sorted(ANIOS, reverse=True)`) para conservar la prosa actual.
    - `filas.py:76`, `datos_fuente.py:48` e `indicadores.py:85` **no
      cambian** (usan `ANIOS[anio]` con `anio` ya contenido en `ANIOS`).
    - Corrida por defecto sin cambios; `--anios 2026` y `--anios 2025` exit
      0; `--anios 2027` → `SystemExit` con mensaje.
  - Verificar:
    ```bash
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos tests generar_sofipos_indicadores.py scripts
    grep -rn "ANIOS" sofipos/   # solo config (def), main (parse_anios), filas, datos_fuente, indicadores
    .venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
    echo "exit: $?"                   # 0
    .venv/bin/python generar_sofipos_indicadores.py --anios 2026 --out /tmp/sofi_2026.xlsx > /tmp/salida_a26.txt
    echo "exit: $?"                   # 0; hojas: Catálogo + Datos/Indicadores 2026 + Diccionario + Semáforos
    .venv/bin/python generar_sofipos_indicadores.py --anios 2025 --out /tmp/sofi_2025.xlsx > /tmp/salida_a25.txt
    echo "exit: $?"                   # 0 (guard BE activo: no cruza boletines de 2026)
    .venv/bin/python generar_sofipos_indicadores.py --anios 2027 ; echo "exit: $?"   # ≠ 0, con mensaje
    ```
  - comparar (corrida default): `IDÉNTICO`.

- [x] **T010 — Docs con números reales**

  - Hallazgos:
    - README: 905 conceptos y serie `201601–202609`; estructura dinámica («3
      fijas + 2 por año», 9 hojas con 2024–2026); 35 entidades (33 activas + 2
      bajas); cobertura enero–septiembre en 2026; validaciones con los números
      de la corrida final (33 meses × 7 conceptos, 1,065 ROA/ROE, 1,031/933/933
      NICAP/IMOR/ICOR, 1,046 entidades-mes de cartera); comandos nuevos;
      «Cómo actualizar» con `range(1, 10)` y `--anios`; excepción de Stori.
    - `entradas/README.md`: referencias a `BE_VALIDACION` (202603/202606) y
      `BE_UNIVERSO` (202609) en `sofipos/config.py`.
    - Docstring del wrapper: año 2024 incluido y las 9 hojas reales (traía «7
      hojas» y omitía 2024).
  - Archivos: `README.md`, `AGENTS.md`, `entradas/README.md`,
    `generar_sofipos_indicadores.py` (docstring).
  - Aceptación (criterios 9-10 de la spec):
    - `grep -rn "generar_sofipos_2026\|SOFIPOS_2026_Datos" README.md
      AGENTS.md entradas/README.md` → sin salida.
    - README: 905 conceptos y serie `201601–202609`; hojas «3 fijas + 2 por
      año» (estructura dinámica); 35 entidades (33 activas + 2 bajas);
      cobertura enero–septiembre; validaciones con los números de la
      corrida final; comandos renombrados; árbol con
      `generar_sofipos_indicadores.py` y `SOFIPOS_Indicadores_Datos.xlsx`;
      sección «Cómo actualizar» con `range(1, 10)` y `--anios`.
    - `AGENTS.md` líneas 11, 34 y 38 con el nombre nuevo.
    - `entradas/README.md:42` con el nombre nuevo; árbol/fecha con el BE
      202609 si aparece.
  - Verificar:
    ```bash
    grep -rn "generar_sofipos_2026\|SOFIPOS_2026_Datos" README.md AGENTS.md entradas/README.md   # vacío
    grep -n "905\|201601–202609\|33 activas\|Cómo actualizar\|--anios" README.md                  # revisar
    .venv/bin/python -m pytest && .venv/bin/pyflakes sofipos tests generar_sofipos_indicadores.py scripts
    ```
  - comparar: `IDÉNTICO` (los docs no cambian el libro).

- [x] **T011 — Verificación global (gate final)**

  - Estado: `pytest` 45 passed · `pyflakes` limpio · corrida default exit 0
    («Validación general: OK») · `--anios 2026` exit 0 · `--anios 2025` exit 0 ·
    `--anios 2027` exit 1 con mensaje · `--help` correcto · grep de nombres
    viejos vacío en docs, código y scripts.
  - Libro por defecto: **9 hojas** (el `# 7 hojas por defecto` del comando de
    esta tarea quedó obsoleto: son 3 fijas + 2 por año con 2024–2026).
    `Datos Fuente 2026`: última fila de datos A319/D319 = `35`/`Septiembre`
    (315 = 35 × 9). Catálogo: tabla principal hasta `B37`
    (`027051`, 33 filas) y apartado en las filas 39–42 con `027028` (revocada)
    y `027014` (transformada).
  - Observe (no se corrige por alcance): con `--anios 2025` se imprime el
    encabezado «Tablero vs boletín BE (marzo, junio y septiembre de 2026…)»
    sin bloques debajo; el guard funciona (no se cruza ningún BE de 2026).

  #### Diferencias documentadas vs baseline de T001

  `comparar.py` sale con **15,108 diferencias** (14,969 de xlsx + 139 de
  stdout), todas esperadas:

  | Hoja / salida | Diffs | Causa |
  |---|---|---|
  | `Catálogo SOFIPOS` | 228 | Universo BE 202609 (33 filas vs 35): reorden y 25 fechas de autorización distintas (p. ej. Klar 2008-06-17 → 2006-09-29); título con corte 30-sep-2026 y fuente `BE_SOFIPOS_202609.xlsx`; apartado nuevo «Sociedades transformadas o revocadas» con `027014` y `027028`; notas al pie actualizadas |
  | `Datos Fuente 2026` | 3,775 | Fila de septiembre nueva (`202609`, 9 meses), consecutivos 34/35 para las bajas, orden del universo 202609 y notas 2026 reescritas |
  | `Datos Fuente 2025` | 5,081 | Reorden por el universo 202609, consecutivos de las bajas y nota común del universo |
  | `Datos Fuente 2024` | 5,060 | Ídem (reorden, consecutivos y nota común) |
  | `Indicadores 2026` | 225 | Misma fila de septiembre y alineación con Datos Fuente 2026 |
  | `Indicadores 2025` / `2024` | 300 / 300 | Reorden y alineación con sus hojas de datos |
  | stdout (139 líneas) | — | «Universo: 33 activas · +2 bajas · BE_SOFIPOS_202609.xlsx»; bloque BE de septiembre («corte definitivo», todo OK, sin excepción); `NICAP … DIF 125.7501` → `OK` (tolerancia relativa); «35 filas (esperado 35)» → 33 en el catálogo; «CON OBSERVACIONES» → «OK» |

  Nota de la fila de septiembre: el baseline de T001 ya se generó con
  `ANIOS = {2026: range(1, 10)}`, así que esa fila aparece frente al libro
  publicado anterior (8 meses), no frente al baseline; frente al baseline lo
  que se ve son los datos `202609` y el reorden, como en la tabla.

  - Archivos: ninguno.
  - Aceptación: los 10 criterios de «Success Criteria» de la spec en verde;
    cada diff de `comparar.py` documentado con su causa (fila/datos de
    septiembre y bajas, catálogo 25 fechas + título + apartado, notas 2026,
    renombres y stdout).
  - Verificar (todo exit 0 salvo `comparar.py`):
    ```bash
    .venv/bin/python -m pytest
    .venv/bin/pyflakes sofipos tests generar_sofipos_indicadores.py scripts
    .venv/bin/sofipos --help                                          # revisar --out y --anios

    .venv/bin/python generar_sofipos_indicadores.py --out /tmp/sofi.xlsx > /tmp/salida2.txt
    echo "exit: $?"                                                   # 0
    grep -q "Validación general: OK" /tmp/salida2.txt && echo OK
    .venv/bin/python generar_sofipos_indicadores.py --anios 2026 --out /tmp/sofi_anios.xlsx > /tmp/salida_a26.txt
    echo "exit: $?"                                                   # 0

    .venv/bin/python scripts/comparar.py \
        --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
        --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt       # ≠ 0: documentar diffs

    grep -rn "generar_sofipos_2026\|SOFIPOS_2026_Datos" \
        README.md AGENTS.md entradas/README.md sofipos tests scripts \
        generar_sofipos_indicadores.py                                 # vacío

    .venv/bin/python -c "
    from openpyxl import load_workbook
    wb = load_workbook('salidas/SOFIPOS_Indicadores_Datos.xlsx')
    print(wb.sheetnames)                                              # 9 hojas por defecto (3 fijas + 2 por año; --anios 2026 da 5)
    ws = wb['Datos Fuente 2026']
    print(ws['A319'].value, ws['D319'].value)                         # última fila de datos: 35×9 = 315
    cat = wb['Catálogo SOFIPOS']
    print(cat['B37'].value)                                           # última fila principal (33); abajo el apartado"
    ```
  - Al cerrar: resumir las diferencias documentadas y **sugerir** los
    comandos de commit (el agente no hace `git commit`/`git push`).

    ```bash
    git add -A
    git commit -m "Corte Q3 2026, CLI --anios y renombres (spec 004)"
    ```

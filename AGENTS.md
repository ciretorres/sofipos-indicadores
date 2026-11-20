# AGENTS.md

Prácticas para los agentes que trabajan en este repo. Léela antes de tocar nada.

## El proyecto

Generador del libro Excel de indicadores de las SOFIPOS (sociedades financieras
populares). Toma datos crudos de la CNBV (`entradas/`) y arma un único `.xlsx`
en `salidas/`.

- Entry point: `.venv/bin/python generar_sofipos_2026.py` (o `python -m sofipos`).
- Código: paquete `sofipos/` — `sofipos/hojas/` tiene una carpeta por pestaña del libro.
- Documentación completa: `README.md` (estructura, método, semáforos, validaciones).

## Desarrollo con specs

- Todo cambio no trivial (que toca varios archivos o tomaría más de 30 min)
  empieza con una spec en `docs/specs/NNN-slug/{spec.md,plan.md,task.md}`.
  Numeración e índice: `docs/specs/README.md`.
- Flujo: SPECIFY → PLAN → TASKS → IMPLEMENT → REVIEW, con **gate humano**
  entre fases (el usuario aprueba `spec.md` y `plan.md` antes de codificar).
  Skills de apoyo: `spec-driven-development`, `planning-and-task-breakdown`,
  `incremental-implementation`, `test-driven-development`, `context-engineering`.
- Los agentes **no hacen `git commit` ni `git push`**: dejan el cambio listo y
  sugieren los comandos; el usuario commitea.

## Verificar un cambio

Obligatorio antes de dar una tarea por terminada: el libro debe quedar idéntico
(si la spec pide cambiarlo, documenta la diferencia esperada).

```bash
# 1. referencia con el código actual
.venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida1.txt
cp /tmp/sofi.xlsx /tmp/referencia.xlsx

# 2. hacer el cambio y volver a generar (misma ruta)
.venv/bin/python generar_sofipos_2026.py --out /tmp/sofi.xlsx > /tmp/salida2.txt

# 3. comparar
.venv/bin/python scripts/comparar.py \
    --xlsx-a /tmp/referencia.xlsx --xlsx-b /tmp/sofi.xlsx \
    --stdout-a /tmp/salida1.txt --stdout-b /tmp/salida2.txt
```

`IDÉNTICO` con exit 0 = OK. Si además cambió la CLI, revisa `--help`; si tocaste
archivos `.py`, `pyflakes` debe salir limpio. Detalle en README §
"Verificar que un cambio no altera el libro".

## MCP de Excel

### Qué es / dónde se configura

`opencode.jsonc` conecta un servidor MCP de Excel (`excel-mcp-server`, sobre
openpyxl; no requiere Excel instalado) con acceso a `entradas/` y `salidas/`.
Sirve de referencia también para Claude Code.

### Requisito: uv

El servidor se lanza con `uvx`. Instalar uv:


#### Windows (PowerShell)

```bash
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
(Get-Command uvx).Source   # → %USERPROFILE%\.local\bin\uvx.exe
```

#### macOS / Linux

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh      
# Windows: instalador de https://docs.astral.sh/uv/getting-started/installation/
which uvx   # → ~/.local/bin/uvx
```

La config usa la **ruta explícita** `%USERPROFILE%\.local\bin\uvx.exe` o `{env:HOME}/.local/bin/uvx`.

### Reviewer

El subagente `reviewer` usa las tools de solo lectura del MCP (`read_range`,
`describe_sheet`); las de escritura están bloqueadas en
`.opencode/agents/reviewer.md`.

### Nota CNBV (customXml)

Los `.xlsx` de la CNBV traen relaciones `../customXml/…` fuera del paquete:
`describe_workbook` falla con ellos (usa `read_range` o openpyxl directo); los
libros de `salidas/` se leen completos.

## Estilo

- Documentación clara, simple y concisa. Docstrings solo cuando aporten:
  no rellenar por rellenar.
- Español.
- Sigue el patrón del código existente en `sofipos/` antes de inventar uno nuevo.

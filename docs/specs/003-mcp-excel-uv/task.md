# Tareas: MCP de Excel con uvx — 003

Orden: T001 → T006. **T001 va primero** (petición del usuario). Marca `[x]`
solo cuando la verificación de esa tarea pasa.

**Nota de verificación de esta spec:** no se toca código Python, así que
`scripts/comparar.py` y `pyflakes` **no aplican**; la verificación de "no
alterar el libro" es `git diff --stat -- sofipos/ scripts/ tests/` → vacío.

Tras editar `opencode.jsonc`, OpenCode recarga la config sola (watcher); si
`opencode mcp list` no refleja el cambio, `opencode service restart`.

## Tareas

- [x] **T001 — `opencode.jsonc`: ruta explícita a uvx + comentarios**
  - Archivos: `opencode.jsonc`.
  - Aceptación: `command` =
    `["{env:HOME}/.local/bin/uvx", "excel-mcp-server", "stdio", "--allow-dir", "entradas", "--allow-dir", "salidas"]`;
    comentarios documentan instalación de uv, el problema de PATH del servicio,
    el workaround histórico de realpath (Monterey → Ventura) y la nota de
    Windows; `$schema` y `default_agent` intactos.
  - Fallback aprobado: si `{env:HOME}` no resuelve, ruta literal
    `/Users/eric/.local/bin/uvx` (documentarlo en AGENTS.md, T003).
  - Verificar: `opencode mcp list` → `excel` conectado (T002).

- [x] **T002 — Verificar conexión del MCP**
  - Archivos: ninguno (solo observación).
  - Aceptación:
    - `opencode mcp list` → `excel` conectado.
    - Sin entradas nuevas de `mcp connect failed` en
      `~/.local/share/opencode/log/opencode.log` posteriores al cambio.
    - Prueba funcional: `read_range` responde sobre un `.xlsx` (p. ej. un
      rango pequeño de `entradas/BE_SOFIPOS_202606.xlsx`).

- [x] **T003 — `AGENTS.md` § MCP de Excel**
  - Archivos: `AGENTS.md`.
  - Aceptación: la sección documenta (ver plan.md § T003 para el detalle):
    instalación de `uv` (`curl -LsSf https://astral.sh/uv/install.sh | sh` +
    `which uvx`), troubleshooting del `NotFound: ChildProcess.spawn`, comando
    de Claude Code, uso por el reviewer, la nota de `customXml` **conservada
    textual** y la nota de realpath (solo Monterey). Estilo: claro, simple,
    conciso, español.

- [x] **T004 — `README.md`: nota opcional del MCP**
  - Archivos: `README.md`.
  - Aceptación: en § Instalación, después de «Instalar dependencias», hay una
    nota corta «MCP de Excel (opcional)» que apunta a `AGENTS.md` § MCP de
    Excel e indica que el generador no necesita `uv`.

- [x] **T005 — `reviewer.md`: tools excel en modo lectura**
  - Archivos: `.opencode/agents/reviewer.md`.
  - Aceptación:
    - Permisos agregados al final del frontmatter, en orden: deny `excel_*`
      y luego allow de `excel_describe_workbook`, `excel_describe_sheet`,
      `excel_read_range`, `excel_list_workbooks` (plan.md § T005).
    - Prompt: viñeta de uso de `read_range`/`describe_sheet` y de por qué no
      `describe_workbook` con archivos CNBV.
    - Test: lanzar el reviewer pidiendo un `read_range` → responde (permisos
      y conexión OK).

- [x] **T006 — Índice de specs + gate final**
  - Archivos: `docs/specs/README.md`.
  - Aceptación:
    - Fila `| [003-mcp-excel-uv](003-mcp-excel-uv/spec.md) | … |` agregada al
      índice con el estado final.
    - Success Criteria 1–7 de la spec en verde.
  - Verificar:
    ```bash
    opencode mcp list                      # excel conectado
    git diff --stat -- sofipos/ scripts/ tests/   # vacío
    ```

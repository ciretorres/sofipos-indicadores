# Plan: MCP de Excel con uvx — 003

Aprobado junto con la spec (gate en conversación, 2026-11-19).

Orden: T001 → T006 (T001 primero, a petición del usuario; T002 verifica
T001; T006 cierra el gate). Sin cambios de código Python.

## T001 — `opencode.jsonc`

Reemplazar el bloque `mcp` y sus comentarios por:

```jsonc
  // MCP de Excel (openpyxl, no requiere Excel instalado; funciona en macOS, Windows y Linux)
  //
  // Requiere uv: https://docs.astral.sh/uv/getting-started/installation/
  //   curl -LsSf https://astral.sh/uv/install.sh | sh      (macOS / Linux)
  //   Verificar: which uvx   →  ~/.local/bin/uvx
  //
  // Se usa la RUTA EXPLÍCITA a uvx porque el servicio de fondo de OpenCode hereda un
  // PATH mínimo (launchd) sin ~/.local/bin: con `uvx` o `uv` sueltos el spawn falla con
  // "NotFound: ChildProcess.spawn". Detalle y troubleshooting en AGENTS.md § MCP de Excel.
  //
  // Nota histórica: en macOS 12 (Monterey) el script del paquete usaba `realpath`, que no
  // existía allí, por eso antes se lanzaba como "python -m excel_mcp". Desde macOS 13
  // (Ventura) realpath está en /bin/realpath y `uvx excel-mcp-server` funciona directo.
  //
  // Windows: si {env:HOME} no aplica, sustituir por la ruta literal de uvx
  // (%USERPROFILE%\.local\bin\uvx.exe) o dejar `uvx` si está en el PATH del servicio.
  "mcp": {
    "servers": {
      "excel": {
        "type": "local",
        "command": [
          "{env:HOME}/.local/bin/uvx",
          "excel-mcp-server",
          "stdio",
          "--allow-dir",
          "entradas",
          "--allow-dir",
          "salidas"
        ]
      }
    }
  }
```

Conservar `$schema` y `default_agent` tal cual están.

**Fallback (aprobado):** si `opencode mcp list` no conecta porque la
sustitución `{env:HOME}` no resuelve, usar la ruta literal
`/Users/eric/.local/bin/uvx` y documentarlo en AGENTS.md.

## T002 — Verificar que el MCP conecta

- `opencode mcp list` → `excel` conectado.
- Sin entradas nuevas de `mcp connect failed` en
  `~/.local/share/opencode/log/opencode.log`.
- Prueba funcional: pedirle al reviewer (o una tool MCP) un `read_range` sobre
  un `.xlsx` (p. ej. `entradas/BE_SOFIPOS_202606.xlsx`, rango pequeño).

## T003 — `AGENTS.md` § MCP de Excel

Reescribir la sección con estas subsecciones (conservar la nota de
`customXml`):

1. **Qué es / dónde se configura** — `opencode.jsonc` conecta el MCP sobre
   `entradas/` y `salidas/`; también sirve de referencia para Claude Code.
2. **Requisito: uv** —
   `curl -LsSf https://astral.sh/uv/install.sh | sh` (macOS/Linux; en Windows
   el instalador de la guía oficial), verificar `which uvx` →
   `~/.local/bin/uvx`. La config usa la ruta explícita
   `{env:HOME}/.local/bin/uvx`.
3. **Troubleshooting** — `NotFound: ChildProcess.spawn (uv …)` en el log =
   el servicio de fondo de OpenCode no ve `~/.local/bin` en su PATH (hereda un
   PATH mínimo de launchd aunque tu shell sí lo tenga); por eso se usa la ruta
   explícita. Verificar con `opencode mcp list`. Reiniciar la TUI no cambia el
   PATH del servicio.
4. **Claude Code** —
   `claude mcp add excel --scope project -- uvx excel-mcp-server stdio --allow-dir entradas --allow-dir salidas`.
5. **Reviewer** — usa las tools de solo lectura (`read_range`); las de
   escritura están bloqueadas en `.opencode/agents/reviewer.md`.
6. **Nota CNBV (conservar textual)** — los `.xlsx` de la CNBV traen
   relaciones `../customXml/…` fuera del paquete: `describe_workbook` falla
   con ellos (usa `read_range` o openpyxl directo); los libros de `salidas/`
   se leen completos.
7. **realpath** — solo era problema en macOS 12 (Monterey); desde Ventura
   existe `/bin/realpath` y no hace falta el workaround `python -m excel_mcp`.

## T004 — `README.md`

En § **Instalación**, después de «Instalar dependencias», agregar una nota
corta «MCP de Excel (opcional)»: para que OpenCode/Claude Code lean los
`.xlsx` con el MCP se instala `uv` (comando en `AGENTS.md` § MCP de Excel);
el generador no lo necesita.

## T005 — `.opencode/agents/reviewer.md`

- **Permisos** (agregar al final del frontmatter, en este orden — gana el
  último que matchea):

  ```yaml
    # Tools del MCP excel: solo lectura
    - action: "excel_*"
      resource: "*"
      effect: deny
    - action: excel_describe_workbook
      resource: "*"
      effect: allow
    - action: excel_describe_sheet
      resource: "*"
      effect: allow
    - action: excel_read_range
      resource: "*"
      effect: allow
    - action: excel_list_workbooks
      resource: "*"
      effect: allow
  ```

- **Prompt** (agregar una viñeta): si el MCP de Excel está disponible, usa
  sus tools de solo lectura (`read_range`, `describe_sheet`) para cotejar el
  libro y los `.xlsx` de `entradas/`; con `describe_workbook` fallan los
  archivos de la CNBV (customXml).

## T006 — Índice de specs

Agregar a `docs/specs/README.md`:
`| [003-mcp-excel-uv](003-mcp-excel-uv/spec.md) | … |` con el estado que
corresponda al cierre.

## Verificación final

- `opencode mcp list` → conectado; log sin `mcp connect failed` nuevo.
- Test del reviewer con `read_range` responde.
- `git diff --stat -- sofipos/ scripts/ tests/` → vacío.
- `pyflakes` no aplica (no se toca `.py`); `comparar.py` no aplica por la
  misma razón — queda documentado en `task.md`.

## Riesgos

1. `{env:HOME}` no resuelve en `command` → fallback a ruta literal (aprobado).
2. El formato de permisos `<server>_<tool>` no matchea → deny inofensivo o
   tools sin bloquear; el test del reviewer lo detecta (spec OQ2).

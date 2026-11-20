# Spec: MCP de Excel con uvx y reviewer en modo lectura — 003

Estado: terminada (review 2026-11-19: 7/7 SC en verde; 1 hallazgo menor
corregido —estado de este encabezado— y 2 notas documentadas abajo).

## Objective

Que el MCP de Excel funcione en OpenCode **sin depender del PATH del servicio
de fondo**, quede documentado para OpenCode y Claude Code, y el subagente
`reviewer` lo use en **modo lectura** real.

### Contexto / diagnóstico (por qué se llega aquí)

1. `opencode.jsonc` lanzaba el servidor con `uv run --no-project --with
   excel-mcp-server …`; **`uv` no estaba instalado** → el log de OpenCode
   registraba `mcp connect failed … NotFound: ChildProcess.spawn (uv …)`.
2. El usuario instaló `uv` (`~/.local/bin/uv`, v0.13.9) y reinició OpenCode,
   pero **el error persistió**: el servicio de fondo (`opencode serve
   --service`, padre launchd) hereda un **PATH mínimo de launchd** que **no
   incluye `~/.local/bin`**, aunque la shell y la TUI sí lo tengan (verificado
   con `ps eww`). El spawn del MCP lo hace el servicio (`role=server` en el
   log), por eso reiniciar la TUI no arregla nada.
3. `uvx excel-mcp-server --help` **sí funciona** en la máquina (paquete se
   descarga y arranca; subcomando `stdio` confirmado): el bloqueo es solo la
   resolución del ejecutable.
4. El comentario de `opencode.jsonc` sobre `realpath` quedó **obsoleto**:
   `realpath` solo faltaba en macOS 12 (Monterey); en Ventura (13.7.8, ya
   actualizado por el usuario) está en `/bin/realpath`.

## Cambios (resumen)

| Archivo | Cambio |
|---|---|
| `opencode.jsonc` | `command` con la ruta explícita `{env:HOME}/.local/bin/uvx` + comentarios actualizados (instalación de uv, problema del PATH, realpath, Windows) |
| `AGENTS.md` § MCP de Excel | Instalación de uv (comando), troubleshooting del `NotFound`, comando para Claude Code, uso por el reviewer, conservar la nota de `customXml` |
| `README.md` | Nota corta «MCP de Excel (opcional)» en § Instalación |
| `.opencode/agents/reviewer.md` | Prompt con las tools excel de lectura + permisos: deny `excel_*` y allow de las tools de solo lectura |
| `docs/specs/README.md` | Índice: spec 003 |

## Tech Stack

- Sin cambios de código Python: **no se toca `sofipos/`** ni nada que altere
  el libro. Sin dependencias nuevas (el usuario ya instaló `uv` manualmente).
- JSONC (`opencode.jsonc`, schema oficial `https://opencode.ai/config.json`)
  y Markdown.

## Commands (verificación)

```bash
# El MCP reconecta al guardar la config (OpenCode la observa); verificar con:
opencode mcp list                                  # excel → conectado
grep "mcp connect failed" ~/.local/share/opencode/log/opencode.log | tail
# → solo deben quedar entradas HISTÓRICAS con el comando viejo (`uv run …`);
#   ninguna posterior al cambio (ni ninguna con `uvx` en el comando).

# La sustitución {env:HOME} debe resolver; si no, fallback (aprobado en plan.md):
#   ruta literal /Users/eric/.local/bin/uvx en el command.

# Cero cambios en el generador:
git diff --stat -- sofipos/ scripts/ tests/        # → vacío
```

## Success Criteria

1. `opencode.jsonc` lanza `{env:HOME}/.local/bin/uvx excel-mcp-server stdio
   --allow-dir entradas --allow-dir salidas` (o la ruta literal aprobada como
   fallback), y los comentarios documentan uv, el PATH del servicio y realpath.
2. `opencode mcp list` muestra `excel` **conectado**; no hay entradas nuevas
   de `mcp connect failed` en el log.
3. `AGENTS.md` § MCP de Excel documenta: instalación de `uv`
   (`curl -LsSf https://astral.sh/uv/install.sh | sh` + `which uvx`),
   troubleshooting del `NotFound`, el comando `claude mcp add …`, el uso del
   reviewer y la nota de `customXml` (conservada).
4. `README.md` tiene la nota opcional del MCP apuntando a AGENTS.md.
5. `reviewer.md` instruye el uso de `read_range` (no `describe_workbook` con
   archivos CNBV) y sus permisos bloquean las tools excel de escritura
   (deny `excel_*` + allow de lectura), con un `test` del reviewer que
   confirme que `read_range` responde.
6. Índice `docs/specs/README.md` incluye la 003.
7. `git diff` sin tocar `sofipos/`, `scripts/`, `tests/` ni el libro.

## Boundaries

- **Always:** documentación clara, simple y concisa; español; conservar la
  nota de `customXml` en AGENTS.md; sin `git commit`/`git push`.
- **Ask first:** ruta literal absoluta en la config si la sustitución
  `{env:HOME}` no resuelve (fallback previamente comentado, pero se documenta
  en AGENTS.md al usarlo).
- **Never:** instalar `uv` desde un agente (lo hizo el usuario); tocar
  `sofipos/` o el contenido del libro; borrar la nota de `customXml`; romper
  el bloqueo de escritura del reviewer.

## Out of Scope (Deferred)

- Agregar 2024 a `ANIOS` con sus notas → **spec 004** (la deferral «spec 003»
  de 002 se refiere a este trabajo, que nunca existió; se renumera para no
  colisionar).
- Instalar/configurar Claude Code (solo se documenta el comando).
- Probar en Windows (solo queda la nota en `opencode.jsonc`/AGENTS.md).

## Open Questions

1. **`{env:HOME}` en `command`:** la guía de variables de OpenCode dice que la
   sustitución aplica a los strings de los archivos de config, pero no está
   probada en `mcp.servers.*.command`. Si no resuelve, fallback a la ruta
   literal (aprobado).
2. **Formato de acción MCP en permisos:** `<server>_<tool>`
   (`excel_read_range`, …) según la guía de permisos V2. Si el deny
   `excel_*` no matchea, las tools quedan permitidas por la política base
   (fail-open) — el test del reviewer lo detecta.
   **Resolución (2026-11-19):** matchea. El catálogo del reviewer expone
   exactamente las 4 tools de su `allow` (`remaining: 0`), mientras una
   sesión sin reglas excel (coordinador) ve las 46 del servidor, incluidas
   `create_workbook` y demás de escritura ⇒ OpenCode filtra el catálogo MCP
   por los permisos del agente; el bloqueo es previo a la llamada.

## Notas de la revisión (2026-11-19)

1. **Deny de escritura no ejercitado en runtime** → cerrada (ver OQ2): la
   frontera es hermética — el reviewer ni siquiera ve las tools de escritura.
2. **`grep` del log devuelve entradas viejas** → aclarado en § Commands: las 2
   entradas de `mcp connect failed` son históricas (comando `uv run` anterior);
   solo se cuenta lo posterior al cambio. El log ya registra
   `mcp connected server=excel tools=46`.

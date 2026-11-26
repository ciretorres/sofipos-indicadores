# Specs

Cada desarrollo no trivial vive en su propia carpeta, con la convención:

```
docs/specs/NNN-slug-de-la-spec/
├── spec.md   # qué y por qué — la aprueba el usuario
├── plan.md   # cómo se implementa — lo aprueba el usuario
└── task.md   # tareas T001… con casillas [ ] que se marcan [x]
```

## Numeración

- `NNN`: siguiente número disponible con 3 dígitos (001, 002, 003…).
- `slug`: minúsculas, sin acentos, guiones en lugar de espacios
  (ej. `002-agregar-hoja-de-resumen`).

## Índice

| Spec | Estado |
| --- | --- |
| [001-tests-y-pyproject](001-tests-y-pyproject/spec.md) | terminada |
| [002-anios-unico-fuente](002-anios-unico-fuente/spec.md) | terminada |
| [003-mcp-excel-uv](003-mcp-excel-uv/spec.md) | terminada |
| [004-corte-q3-anios-cli](004-corte-q3-anios-cli/spec.md) | terminada |

Estados: `borrador` → `aprobada` → `en implementación` → `revisada` → `terminada`.

## Flujo

1. **SPECIFY** — el planner escribe `spec.md`; gate: aprueba el usuario.
2. **PLAN** — el planner escribe `plan.md` y `task.md`; gate: aprueba el usuario.
3. **IMPLEMENT** — el implementer ejecuta tareas de `task.md` y marca `[x]`.
4. **REVIEW** — el reviewer revisa contra la spec; hallazgos → vuelta al implementer.

Detalle en la skill `spec-driven-development` y en `AGENTS.md`.

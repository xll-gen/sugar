# Claude Code Instructions

**This file is intentionally minimal.** All durable agent guidance — the **xlwings-parity mission**, the Excel object-model roadmap, the `.Options()` framework spec, and the improvement backlog — lives in [`AGENTS.md`](./AGENTS.md).

Before doing anything in this repository:

1. Read **[AGENTS.md](./AGENTS.md)** in full. It is the single source of truth.
2. Pay particular attention to **Section 0 (Mission Statement)** and **Section 2 (xlwings Feature-Parity Roadmap)** — they shape every API decision in `sugar/excel/`.
3. When in doubt about an API name or behavior, default to the xlwings equivalent: <https://docs.xlwings.org/en/stable/api.html>.
4. Do **not** add project-specific guidance to this file. Add it to `AGENTS.md` so every agent tool (Claude Code, Codex, Cursor, Aider, etc.) sees it.

Updating `CLAUDE.md` to anything other than this redirect is a policy violation; update `AGENTS.md` instead.

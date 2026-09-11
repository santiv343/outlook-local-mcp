# Project instructions

- Outlook access is read-only. Never send, save, mark read, move, delete, open attachments,
  add stores, change account settings, or call Outlook.Quit.
- Attach only to an already running classic Outlook with a loaded profile.
- Create, use and release every COM object on the worker's main STA thread.
  Only JSON data crosses the private worker pipe. The parent owns worker termination.
- stdout is MCP only. Log operation, duration, counts and stable error codes only.
  Never log mail data, queries, identifiers, raw COM exceptions or validation inputs.
- Fixtures, examples, screenshots and committed artifacts must be entirely synthetic.
- Keep all code, comments, tool names, descriptions, errors and documentation in English.
- Reuse existing mechanisms. Do not add frameworks for state, dependency injection,
  telemetry, persistence, HTTP, providers or background synchronization.
- Current decisions and progress: docs/status.md and docs/architecture.md.
- Validate with `uv run --locked ruff check .`, `uv run --locked ruff format --check .`,
  `uv run --locked mypy`, `uv run --locked pytest`, and `uv build --no-sources`.
- Run `uv run --locked python scripts/smoke_test.py` for the real MCP/Outlook journey.
  Report counts and pass/fail only; never persist mailbox contents.
- Review the exact candidate independently before public delivery. Tests alone do
  not establish Outlook or Claude compatibility. Record unverified steps honestly.

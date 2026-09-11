# Verification status

Target: a public, read-only Windows Outlook MCP package for local stdio clients,
with automatic Python provisioning through uvx.

## Verified locally

- Real MCP client against classic Outlook on Windows 11: all six tools discovered;
  status, mailbox/folder listing, five recent messages, known-subject search,
  body continuation, unchanged read/unread state, invalid arguments and missing
  item handling passed. Repeated after process ownership and typing changes.
- Codex CLI: server registered in user configuration while preserving other
  entries. A fresh Codex session discovered six tools and successfully called
  `outlook_status` against Outlook. This initially used the local development package.
- Real subprocess tests cover hard timeout, recovery, bounded admission,
  queued timeout isolation, cancellation and client EOF cleanup.
- Windows venv launchers spawn a child interpreter. The supervisor invokes the
  CPython base executable with the venv launcher environment, so its owned PID is
  the process executing COM. Tests compare the owned and executing process IDs.
- Synthetic tests cover dates, regional filter formatting, cursors, inaccessible
  properties, read-only access and configuration preservation.

## v0.1.0 published

- Release: https://github.com/santiv343/outlook-local-mcp/releases/tag/v0.1.0
- Reviewed source: `4bcca4936fc25775cf9a58d83eab2aa7ae7627c8`.
- All 49 tests passed on Python 3.11 and 3.12. Lint, formatting, strict types and
  wheel/source builds passed. Independent correction review passed 26 targeted tests.
- GitHub Windows CI passed: https://github.com/santiv343/outlook-local-mcp/actions/runs/34635044295
- The public wheel and constraints were fetched outside the checkout into a fresh
  uv cache with a fresh managed Python installation. Doctor and the complete real
  Outlook MCP smoke test passed using that installation.
- Codex now uses the published uvx command. A fresh Codex session called
  `outlook_status` successfully against the release.
- Claude Desktop configuration was preserved and backed up. After launch, its logs
  confirm protocol initialization and a `tools/list` response from the release.
  An AI-triggered tool call inside Claude Desktop was not exercised.
- Other documented clients have not been exercised locally. Windows 10 has not
  been tested on a physical machine; the real Outlook validation used Windows 11.

## Next authorized work

Keep v0.1.0 available while implementing a separate feature branch for opening
messages in Outlook, creating drafts/replies, reading conversations, additional
recipient/category/importance/attachment-name filters, and optional reviewed sending.
Mutation tools remain disabled by default. The implementation contract and its
challenge must settle write recovery and send confirmation before implementation.

Research and specification challenge were required and completed for MCP version,
COM, filters, process ownership and distribution. Independent implementation review
approved v0.1.0 before publication. Verification records contain no mailbox contents,
account names, personal paths or credentials.

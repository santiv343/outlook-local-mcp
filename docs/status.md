# Verification status

Target: a public Windows Outlook MCP package for local stdio clients, read-only
by default with opt-in actions and automatic Python provisioning through uvx.

## Initial read-only validation

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
- Codex switched to the published uvx command. A fresh Codex session called
  `outlook_status` successfully against the release.
- Claude Desktop configuration was preserved and backed up. After launch, its logs
  confirm protocol initialization and a `tools/list` response from the release.
  An AI-triggered tool call inside Claude Desktop was not exercised.
- Other documented clients have not been exercised locally. Windows 10 has not
  been tested on a physical machine; the real Outlook validation used Windows 11.

## v0.2.0 published

- Release: https://github.com/santiv343/outlook-local-mcp/releases/tag/v0.2.0
- Independently approved source: `04fdc1d457ccb5e372cb8ede397ae449fb0c717c`.
- Published source: `9082742f07f3b0ed7b471415d49fff82eeea458e`, merged through
  https://github.com/santiv343/outlook-local-mcp/pull/1 with an identical source tree.
- All 136 tests passed on Python 3.11 and 3.12, with lint, formatting, strict types
  and wheel/source builds. Windows release CI passed:
  https://github.com/santiv343/outlook-local-mcp/actions/runs/34643716209
- All four uploaded asset hashes match the inspected local packages, constraints
  and checksum file. Package members and tracked files contain no personal paths,
  mailbox data, client configuration or the supplied mailbox transcript.
- Public wheel installation outside the checkout, using an empty uv cache and a
  new managed Python installation, passed version/doctor and the complete real
  MCP smoke test: seven-tool discovery, mailbox/folder listing, recent messages,
  subject/date searches, short-reference reuse, body continuation, unchanged read
  state, invalid/missing inputs and clean shutdown.
- Real MCP checks on the reviewed implementation passed native conversation
  traversal/continuation, combined available metadata filters, draft creation,
  complete send preview, opening, native replies and original read-state preservation.
  The final From correction also passed a preview of an existing synthetic draft.
- Four test drafts remain in Outlook. No live send was attempted. Submission and
  failure recovery are covered by synthetic COM and real subprocess tests.

### Client and context results

Codex and Claude Desktop configurations now reference the public v0.2.0 command,
with the optional write/send capabilities enabled locally. Other settings were
preserved and backups were created. The package remains read-only by default.

A fresh Codex session answered a monthly Inbox question with one `search_emails`
call, using date filters, the default Inbox and a twenty-item limit. It produced a
final response mentioning the Inbox scope; no discovery or body calls were made.
The serialized MCP result was 1,411 bytes in that run. This is one observed client
journey, not a guarantee of every agent's future planning or answer correctness.
Only operation names, argument names, sizes and completion metadata were retained.

Synthetic twenty-message JSON fell from 25,430 to 7,076 bytes (72.2%) after replacing
long native IDs with short references. Tokens and end-to-end billing were not measured.
The date-search regression also passes the real Windows formatter and public MCP test.

Existing Codex sessions must reload their MCP server or start a new session. Claude
Desktop must be restarted to load its updated command; v0.2.0 initialization and an
AI-triggered tool invocation inside Claude Desktop have not been exercised. Other
documented clients and a physical Windows 10 installation remain untested.

### Review history

The initial implementation review found four issues: final account verification,
represented From access, cleanup masking unknown mutation outcomes and required
metadata denial classification. The first correction closed three findings. The
repeated From failure triggered a bounded contract revisit, independently approved
at `7bae397`, to distinguish exact property absence from general locator errors.
The second correction closed the remaining finding; all four were closed before
integration. The original review and both correction cycles remain recorded in
docs/v0.2-plan.md and the PR history. Research, contract challenge, exact independent
review and public installation verification are complete for this release.

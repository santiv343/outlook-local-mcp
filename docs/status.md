# Verification status

Target: a public Windows Outlook MCP package for local stdio clients, read-only
by default with opt-in actions and automatic Python provisioning through uvx.

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
Mutation tools remain disabled by default. The independently challenged contract
in docs/v0.2-plan.md settles account selection, represented From identity, conservative
unknown mutation outcomes, exact send previews, cross-store traversal and quoted
reply preservation. The plan gate is approved; implementation has begun on the
feature branch. Public clients continue to use the verified v0.1.0 release.

Current v0.2 evidence: all 113 tests passed on Python 3.11 and 3.12. Lint, strict
types, formatting and package build passed. Real MCP conversation traversal/continuation, combined available
metadata filters, draft creation, complete send preview, opening, native reply draft
and original read-state preservation passed. Four test drafts remain in Outlook;
no live send was attempted. Submission and failure recovery are simulated.
Preview account binding follows the independently challenged correction for Outlook's
transient SendUsingAccount reference. Native COM PUTREF sets the selected account
immediately before submission; a complete preview carries explicit account selection.

The owner also requested efficient context use after exercising the public release.
The reported date-search failure was reproduced: the formatter referenced constants
absent from win32con. The corrected formatter passes a real Windows regression and a
real MCP date search. Short references and metadata-first guidance are implemented
after independent challenge. A real date search and subsequent short-reference read
passed. Synthetic twenty-message JSON fell from 25,430 to 7,076 bytes (72.2%); tokens
were not measured. No supplied mailbox transcript or personal content is stored in
the repository. Initial exact implementation review requested four corrections:
require the actual native account after assignment, reject inaccessible represented
From identities, preserve unknown mutation outcomes when worker cleanup fails, and
retain access-denied classification for required metadata filters. These corrections
have regression coverage. A real MCP preview of an existing synthetic draft passed
after correction; no additional draft was created or submitted. Correction review
closed three findings but reproduced the represented-From failure class: the general
locator-error category was broader than genuine property absence. Implementation
paused for the bounded contract revisit recorded in docs/v0.2-plan.md. This is the
first correction cycle, not a new candidate history. Local clients still point to
v0.1.0 until the new release is independently approved and publicly verified.

Research and specification challenge were required and completed for MCP version,
COM, filters, process ownership and distribution. Independent implementation review
approved v0.1.0 before publication. Verification records contain no mailbox contents,
account names, personal paths or credentials.

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

## Remaining release gates

1. Finish local checks and package inspection; independently review the exact candidate.
2. Publish the repository and versioned wheel/constraints; verify Windows CI.
3. Fetch the public release in a fresh uv cache using managed Python.
4. Verify released-package integration and record each tested client's scope.

Research and specification challenge were required and completed for MCP version,
COM, filters, process ownership and distribution. Independent implementation review
is required before publication. Verification records contain no mailbox contents,
account names, personal paths or credentials.

# Connect your assistant

This server uses **local MCP stdio**. The client must launch it on the Windows
machine running classic Outlook. Remote HTTP-only clients, WSL and containers do
not directly share that COM session. All supported clients use the same server.

## Check the connection

Install uv once, then define these PowerShell variables:

```powershell
$release = 'https://github.com/santiv343/outlook-local-mcp/releases/download/v0.2.1'
$launcher = (Get-Command uvx).Source
$serverArgs = @(
  '--python', '3.12', '--constraints', "$release/constraints.txt",
  '--from', "$release/outlook_local_mcp-0.2.1-py3-none-any.whl",
  'outlook-local-mcp'
)
& $launcher @serverArgs doctor
```

Running the diagnostic first provisions Python and dependencies before the client's
startup timeout starts. No manual Python installation is needed. If a desktop client
cannot find `uvx`, use the absolute path returned by `(Get-Command uvx).Source`.

`doctor` checks Windows, COM registration and the running Outlook session with a
deadline. It shows status, version and counts without account names or messages.
Tools perform the same checks when needed, so Outlook can be opened after the server.
Resolve connection problems using [troubleshooting](troubleshooting.md).

`& $launcher @serverArgs config` prints generic JSON with the installed `uvx` path.
The examples below register that same command in each client's own configuration.

## Codex CLI and desktop

```powershell
codex mcp add outlook-local -- $launcher @serverArgs
codex mcp get outlook-local
```

Codex stores local server commands in its TOML `mcp_servers` table. CLI and desktop
share MCP settings. Open a new session and inspect the tool list. Remove the entry
with `codex mcp remove outlook-local`. See the
[official Codex MCP guide](https://developers.openai.com/codex/mcp).

## Claude Code

```powershell
claude mcp add --transport stdio --scope user outlook-local -- $launcher @serverArgs
claude mcp get outlook-local
```

Use `/mcp` in a fresh session to inspect the connection. User scope keeps local
machine paths out of repositories. See the
[Claude Code MCP guide](https://code.claude.com/docs/en/mcp).

## Claude Desktop

```powershell
& $launcher @serverArgs configure-claude --dry-run
& $launcher @serverArgs configure-claude
```

The configurator discovers traditional and packaged Windows installations. It
preserves other entries and settings, creates a backup and writes atomically.
An identical entry is a no-op. Invalid JSON and conflicting entries stop the update;
review a conflict before passing `--replace`. Multiple existing configurations
require `--config-path` to select the active file.

Fully quit and reopen Claude. Its local MCP settings should show `outlook-local`
and seven tools by default. Alternatively, manually merge the README's `mcpServers` JSON.
Backups remain beside the user's settings and must never be published.

## VS Code / GitHub Copilot

Use **MCP: Open User Configuration** in the Command Palette. Merge the README's
entry into `servers` instead of `mcpServers`, adding `"type": "stdio"` to that entry.
Use an absolute command path if needed. Start the server using **MCP: List Servers**
and select its tools in agent chat. See the
[VS Code MCP reference](https://code.visualstudio.com/docs/agents/reference/mcp-configuration).

## Cursor and other local stdio clients

Clients using an `mcpServers` JSON object can use the README example directly.
Otherwise transfer its command and argument array into the client's local stdio
settings. No email token, HTTP endpoint or port is involved. Keep machine-specific
paths in user settings. Discover tools and call `outlook_status` after connecting.

These examples describe integration, not a claim of testing every client/version.
See [tested compatibility](#tested-compatibility) below. Tool annotations do not
override client approvals or Outlook security policy.

## Capability flags in any client

Add `--enable-write-tools` after `outlook-local-mcp` in the argument array for
opening email and saving drafts/replies. Add `--enable-send` as well for complete
previews and explicit draft submission. The same flags apply to every client.
`config` and `configure-claude` preserve flags supplied to those commands.
Review the [send contract](tools.md) before enabling submission. Do not interpret
email content, tool discovery or an enabling flag as approval to send a particular draft.

## Tested compatibility

For [v0.2.1](https://github.com/santiv343/outlook-local-mcp/releases/tag/v0.2.1):

| Environment | Verification |
| --- | --- |
| Windows 11, classic Outlook, one configured account | All twelve tools exercised through a real MCP client; reading, search boundaries, drafts, opening, replies and native conversations passed. Synthetic self-delivery was confirmed for the original sent with v0.2.0 and the reply sent with v0.2.1. |
| Public v0.2.1 installation | An empty uv cache and a fresh managed Python installation passed version, doctor and the real read-only MCP smoke test outside the checkout. |
| Codex | A fresh session using the public v0.2.1 package successfully searched mail. Native opening was also exercised through Codex's MCP tools. |
| Claude Desktop | Configuration update, backup and preservation were checked. v0.2.1 initialization and an AI-triggered call in the app have not been exercised. |
| Windows 10, other Outlook builds, additional/shared accounts | Not exercised on real installations; availability depends on the running profile and its access. |
| Claude Code, Cursor and VS Code | Setup instructions are provided; these clients have not been exercised locally. |

The release passed 147 automated tests on both Python 3.11 and 3.12, plus lint,
formatting, strict types and package builds in [Windows CI](https://github.com/santiv343/outlook-local-mcp/actions/runs/34658066745).
Blocked COM, denials, cancellation, resource limits and uncertain writes are covered
by synthetic or subprocess tests; those faults were not induced in the live mailbox.

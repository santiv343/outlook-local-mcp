# Troubleshooting

Append `doctor` to the versioned uvx command in the README. Tools perform the same
Outlook checks lazily when a connection is needed.

| Symptom | Action |
| --- | --- |
| `uvx` not found | Install uv, reopen the client, or use the absolute path from `(Get-Command uvx).Source` |
| First startup exceeds timeout | Run `doctor` in a terminal to provision Python and dependencies first |
| `UNSUPPORTED_PLATFORM` | Run on native Windows with Outlook, outside WSL/remote containers |
| `OUTLOOK_NOT_INSTALLED` | Install or repair classic Outlook; new Outlook lacks this Object Model |
| `OUTLOOK_UNAVAILABLE` | Open classic Outlook, load the profile and resolve visible dialogs |
| `OUTLOOK_TIMEOUT` | Resolve blocked Outlook dialogs before retrying; old cursors may be invalid |
| `ACCESS_DENIED` | Check profile permissions and corporate Object Model policy |
| `BODY_UNAVAILABLE` | Let Outlook load/synchronize the body normally under its policy, then retry |
| Store/folder not found | Rediscover IDs; archive stores may require selecting a folder explicitly |
| `ITEM_NOT_FOUND` | A message may have moved or disappeared; search for its current ID |
| `CURSOR_EXPIRED` | Restart the search; tokens are one-use and expire after ten minutes or a restart |
| `SEARCH_SESSION_LIMIT` | Narrow the folder/date range or let unused sessions expire |
| `SERVER_BUSY` | Let active and queued operations finish before submitting more work |
| Empty search page | Inspect coverage and continue its cursor; a partial empty page is inconclusive |
| Configurator rejects JSON | Repair the existing file first; invalid files are never overwritten |
| Multiple Claude configurations | Select the active file with `--config-path` |

The server never starts or closes Outlook. Different Windows users, sessions or
elevation levels may prevent access to a running instance. This project does not
require administrator rights or changes to PowerShell execution policy.

If corporate policy blocks scripts/downloads, use an approved uv installation and
manually merge client configuration. Developers with an approved Python runtime
can use the checkout instructions. Do not weaken corporate controls.

Report only package/Python/Windows versions, operation name, stable error code and
synthetic reproduction steps. Never include mailbox names, IDs, addresses, contents,
searches, PST/OST files or full client configuration. stderr metrics are content-free;
stdout is reserved for MCP.

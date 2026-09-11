# Architecture

```mermaid
flowchart TD
    Client[Local Windows MCP client] -->|MCP stdio| Server[Python MCP server]
    Server -->|Private JSON pipes| Worker[Owned worker process / one STA thread]
    Worker -->|Read-only COM| Outlook[Running classic Outlook / existing profile]
```

The official MCP Python SDK 2.x owns protocol handling. The server uses explicit
schemas and read-only annotations. All clients share the same six tools.

## Responsibilities

- `server.py`, `contracts.py`, `models.py`: validate inputs, publish schemas,
  sanitize errors and log content-free metrics. stdout is reserved for MCP.
- `supervisor.py`: one active operation and one queued operation. The 30-second
  deadline starts at admission, including queue time. A third request gets `SERVER_BUSY`.
- `worker.py`, `outlook.py`, `com_types.py`: COM objects live exclusively on the
  worker's main STA thread. Connection is lazy and attaches only to running Outlook.
- `mail.py`, `filters.py`, `pagination.py`: map properties and traverse bounded
  collections. Only serializable data crosses the private worker pipes.
- `client_config.py`: optional client setup, outside the MCP transport.

## Process lifecycle

A task timeout cannot interrupt blocked COM. Active timeouts and cancellations
terminate and reap the owned worker. A subsequent call creates a new worker,
invalidating old cursors. Queued timeouts never kill the operation ahead of them.
The server never starts Outlook, opens login dialogs or calls Outlook.Quit.

Client EOF cancels active work through the SDK connection lifecycle; final cleanup
closes the supervisor. Cleanup has a two-second deadline beyond the operation
deadline. An operating-system kill of the MCP parent is outside the graceful
connection-shutdown guarantee.

On Windows a CPython venv executable launches another process. The supervisor
invokes the base interpreter and sets `__PYVENV_LAUNCHER__` to preserve the current
environment. This CPython-specific boundary has real PID/exit tests. See
[CPython path initialization](https://github.com/python/cpython/blob/3.12/Modules/getpath.py).
No process-name searches or blanket process-tree termination are used.

## Search and consistency

Calls inspect up to 1,000 candidates or ten cooperative seconds, with an external
30-second deadline for blocking COM. Date filters generate a conservatively widened
UTC DASL window using Windows regional formatting. Python enforces exact boundaries
and literal text predicates. User text never enters DASL expressions.

One-use cursors bind the tool, location and filters; only page size can change.
Validate before consuming. Invalid arguments preserve the cursor, while a failure
after advancing discards that continuation. Cursor state contains no message bodies.

Cursors have an absolute ten-minute lifetime and twenty active sessions maximum.
Each session remembers up to 20,000 IDs or approximately 4 MiB of ID bookkeeping.
Duplicates are suppressed by EntryID; arrivals, moves and deletions can still change
results. There is no snapshot or exact total. Exhaustion means traversal ended;
`evaluation_complete` additionally requires no omitted candidates across the session.

References: [STA requirement](https://learn.microsoft.com/en-us/office/client-developer/outlook/selecting-an-api-or-technology-for-developing-solutions-for-outlook),
[Items.Restrict](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.restrict),
[date comparisons](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/filtering-items-using-a-date-time-comparison).

## Distribution

`uv.lock` records resolved dependencies. Releases include a wheel and runtime
constraints exported from that lock. The uvx command passes both: a wheel alone
would not enforce the source lock. Versioned URLs make updates deliberate. A release
is installation-verified only after fetching its public assets outside the checkout.

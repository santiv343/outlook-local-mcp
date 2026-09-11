# Efficient agent use

For a question such as “Which emails arrived this month?”, call `search_emails`
with the relevant `after`/`before` dates. Omit location for the default Inbox; do not
discover mailboxes and folders first. Start with the normal twenty-item page, answer
from subjects/senders/dates and follow a cursor only if the requested coverage needs it.

Make calls sequentially. Outlook handles one active operation; parallel requests can
fill the bounded queue and cause `SERVER_BUSY`. Wait for active work before trying again.

For “What do those messages say?”, select the relevant results and use `read_email`
with a short body page. The default is 2,000 Unicode characters, maximum 30,000.
Continue only if needed for the answer. Bodies are plain text as supplied by Outlook;
the server does not remove quotations, signatures or links and pretend the content
is complete. Complete send previews always include the whole supported body.

Reuse returned short identifiers. Native IDs can be hundreds of characters, especially
store IDs repeated across a page. In a synthetic twenty-message example, replacing
IDs reduced compact JSON from 25,430 to 7,076 bytes (72.2%). This measures bytes,
not tokens or end-to-end billing. Actual savings depend on the mailbox and client.
The references are temporary: ten minutes since last use, and invalid after a reset.

Search covers one local folder per call. Do not silently broaden an Inbox question
to archives, spam or Deleted Items. Conversely, do not conclude “none in the account”
from an Inbox-only search. Report the scope and inspect `coverage.evaluation_complete`;
an exhausted traversal with omitted candidates is still inconclusive.

The MCP initialization instructions and tool descriptions teach these choices to
any client. They guide an agent; they cannot force a client's planning policy.
Client history compaction can reduce accumulated context, but does not prevent
unnecessary calls or oversized responses at their source.

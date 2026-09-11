"""Stable error descriptions; operational code refers to named error values."""

from .enums import EErrorCode

ERRORS: dict[EErrorCode, tuple[str, bool]] = {
    EErrorCode.REFERENCE_EXPIRED: (
        "This short reference expired or belongs to an earlier worker. "
        "Search/list again to obtain current references.",
        False,
    ),
    EErrorCode.REFERENCE_LIMIT: (
        "The reference budget was reached. All references, cursors and send previews were reset. "
        "Search/list again with a narrower range.",
        False,
    ),
    EErrorCode.UNSUPPORTED_PLATFORM: ("This server requires Windows and classic Outlook.", False),
    EErrorCode.OUTLOOK_NOT_INSTALLED: (
        "Classic Outlook is not registered. Install classic Outlook.",
        False,
    ),
    EErrorCode.OUTLOOK_UNAVAILABLE: (
        "Open classic Outlook, load your profile, and close any pending dialogs; then retry.",
        True,
    ),
    EErrorCode.OUTLOOK_TIMEOUT: (
        "The request timed out. Check Outlook for pending dialogs and retry. "
        "Existing references, cursors and send previews may have expired.",
        True,
    ),
    EErrorCode.ACCESS_DENIED: (
        "Outlook denied access. Check your profile permissions and security policy.",
        False,
    ),
    EErrorCode.STORE_NOT_FOUND: ("Mailbox not found. List the available mailboxes again.", False),
    EErrorCode.FOLDER_NOT_FOUND: (
        "Folder not found in that mailbox. List its folders again.",
        False,
    ),
    EErrorCode.ITEM_NOT_FOUND: (
        "Email not found; it may have moved or been deleted. Search again.",
        False,
    ),
    EErrorCode.BODY_UNAVAILABLE: (
        "Outlook did not provide the email body. Check its availability in Outlook.",
        False,
    ),
    EErrorCode.METADATA_UNAVAILABLE: (
        "Outlook did not expose metadata required to evaluate this filter.",
        False,
    ),
    EErrorCode.INVALID_ARGUMENT: (
        "Invalid arguments. Check the tool schema and parameter relationships.",
        False,
    ),
    EErrorCode.CURSOR_EXPIRED: (
        "Cursor expired, was already consumed, or was invalidated. Start again.",
        False,
    ),
    EErrorCode.SEARCH_SESSION_LIMIT: (
        "Search session limit reached. Narrow the date range or wait for expiry.",
        False,
    ),
    EErrorCode.SERVER_BUSY: (
        "The server is busy. Retry after the current operation finishes.",
        True,
    ),
    EErrorCode.INTERNAL_ERROR: ("The operation failed unexpectedly. Run doctor and retry.", True),
    EErrorCode.CAPABILITY_DISABLED: (
        "This capability is disabled in the server configuration.",
        False,
    ),
    EErrorCode.WRITE_OUTCOME_UNKNOWN: (
        "The Outlook change may have succeeded. Do not repeat it automatically. "
        "Inspect Outlook, Drafts, Outbox or Sent Items before taking another action.",
        False,
    ),
    EErrorCode.UNSUPPORTED_COMPOSITION: (
        "This draft or sending identity is not supported for programmatic submission. "
        "Review and send it in Outlook.",
        False,
    ),
    EErrorCode.CONFIRMATION_INVALID: (
        "Send confirmation is missing, expired, replaced or consumed. Prepare a new preview.",
        False,
    ),
    EErrorCode.DRAFT_CHANGED: (
        "The draft changed after its preview. Prepare and review a new preview before sending.",
        False,
    ),
    EErrorCode.CONVERSATION_UNAVAILABLE: (
        "Outlook does not expose a conversation for this item or store.",
        False,
    ),
}

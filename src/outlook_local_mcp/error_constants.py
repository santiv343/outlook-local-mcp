"""Stable error descriptions; operational code refers to named error values."""

from .enums import EErrorCode

ERRORS: dict[EErrorCode, tuple[str, bool]] = {
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
        "Existing cursors may have expired.",
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
}

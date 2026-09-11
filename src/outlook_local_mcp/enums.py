"""Named values for public contracts and operational state."""

from enum import StrEnum


class EToolName(StrEnum):
    OUTLOOK_STATUS = "outlook_status"
    LIST_MAILBOXES = "list_mailboxes"
    LIST_FOLDERS = "list_folders"
    RECENT_EMAILS = "recent_emails"
    SEARCH_EMAILS = "search_emails"
    READ_EMAIL = "read_email"


class EErrorCode(StrEnum):
    UNSUPPORTED_PLATFORM = "UNSUPPORTED_PLATFORM"
    OUTLOOK_NOT_INSTALLED = "OUTLOOK_NOT_INSTALLED"
    OUTLOOK_UNAVAILABLE = "OUTLOOK_UNAVAILABLE"
    OUTLOOK_TIMEOUT = "OUTLOOK_TIMEOUT"
    ACCESS_DENIED = "ACCESS_DENIED"
    STORE_NOT_FOUND = "STORE_NOT_FOUND"
    FOLDER_NOT_FOUND = "FOLDER_NOT_FOUND"
    ITEM_NOT_FOUND = "ITEM_NOT_FOUND"
    BODY_UNAVAILABLE = "BODY_UNAVAILABLE"
    METADATA_UNAVAILABLE = "METADATA_UNAVAILABLE"
    INVALID_ARGUMENT = "INVALID_ARGUMENT"
    CURSOR_EXPIRED = "CURSOR_EXPIRED"
    SEARCH_SESSION_LIMIT = "SEARCH_SESSION_LIMIT"
    SERVER_BUSY = "SERVER_BUSY"
    INTERNAL_ERROR = "INTERNAL_ERROR"


class EStopReason(StrEnum):
    EXHAUSTED = "exhausted"
    PAGE_LIMIT = "page_limit"
    SCAN_LIMIT = "scan_limit"
    TIME_BUDGET = "time_budget"


class ERecipientKind(StrEnum):
    TO = "to"
    CC = "cc"
    BCC = "bcc"
    UNKNOWN = "unknown"

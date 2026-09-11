"""Documented Outlook Object Model, MAPI and Windows COM identifiers."""

from .enums import ERecipientKind
from .models import TImportance

OUTLOOK_PROGRAM_ID = "Outlook.Application"
OUTLOOK_REGISTRATION = r"Outlook.Application\CLSID"
MAPI_NAMESPACE = "MAPI"
MAIL_ITEM_CLASS = 43
INBOX_FOLDER = 6
DRAFTS_FOLDER = 16
PLAIN_TEXT_FORMAT = 1
SENDING_ACCOUNT_PROPERTY = "SendUsingAccount"
COM_NEUTRAL_LOCALE = 0
WINDOWS_DATE_SHORTDATE = 0x00000001
WINDOWS_TIME_NOSECONDS = 0x00000002
MAIL_MESSAGE_CLASS = "IPM.Note"
STORE_ENTRY_ID_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x0FFB0102"
REPRESENTING_SMTP_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F"
WINDOWS_REGIONAL_SETTINGS = r"Control Panel\International"
WINDOWS_LIST_SEPARATOR = "sList"
HEADER_ONLY = 0
SMTP_ADDRESS_TYPE = "SMTP"
SMTP_ADDRESS_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
RECEIVED_TIME_PROPERTY = "[ReceivedTime]"
RECEIVED_TIME_DASL_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
RECIPIENT_KINDS = {1: ERecipientKind.TO, 2: ERecipientKind.CC, 3: ERecipientKind.BCC}
IMPORTANCE_VALUES: dict[int, TImportance] = {0: "low", 1: "normal", 2: "high"}
ACCESS_DENIED_HRESULTS = frozenset({0x80070005, 0x80030005})
MAPI_NOT_FOUND_HRESULT = 0x8004010F
DISPATCH_EXCEPTION_HRESULT = 0x80020009
NOT_FOUND_HRESULTS = frozenset({MAPI_NOT_FOUND_HRESULT, 0x80070002, 0x80040107})
HRESULT_MASK = 0xFFFFFFFF
EXCEPINFO_STATUS_INDEX = 5

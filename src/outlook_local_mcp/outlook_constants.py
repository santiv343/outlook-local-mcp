"""Documented Outlook Object Model, MAPI and Windows COM identifiers."""

from .enums import ERecipientKind

OUTLOOK_PROGRAM_ID = "Outlook.Application"
OUTLOOK_REGISTRATION = r"Outlook.Application\CLSID"
MAPI_NAMESPACE = "MAPI"
MAIL_ITEM_CLASS = 43
INBOX_FOLDER = 6
HEADER_ONLY = 0
SMTP_ADDRESS_TYPE = "SMTP"
SMTP_ADDRESS_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
RECEIVED_TIME_PROPERTY = "[ReceivedTime]"
RECEIVED_TIME_DASL_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
RECIPIENT_KINDS = {1: ERecipientKind.TO, 2: ERecipientKind.CC, 3: ERecipientKind.BCC}
ACCESS_DENIED_HRESULTS = frozenset({0x80070005, 0x80030005, 0x80040102})
NOT_FOUND_HRESULTS = frozenset({0x8004010F, 0x80070002, 0x80040107})
HRESULT_MASK = 0xFFFFFFFF
EXCEPINFO_STATUS_INDEX = 5

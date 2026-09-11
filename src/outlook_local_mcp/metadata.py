"""Optional message metadata and bounded literal recipient/attachment predicates."""

from .addresses import address_entry_email
from .com_types import IOutlookItem
from .config import MAX_ATTACHMENTS, MAX_RECIPIENTS
from .enums import EErrorCode
from .error_constants import CATEGORIES_ACCESS_DENIED_WARNING, IMPORTANCE_ACCESS_DENIED_WARNING
from .errors import OutlookError, com_error
from .models import MessageMetadata, SearchArguments, WarningInfo
from .outlook_constants import IMPORTANCE_VALUES, WINDOWS_LIST_SEPARATOR, WINDOWS_REGIONAL_SETTINGS


def category_names(item: IOutlookItem) -> list[str]:
    raw = item.Categories
    if not raw:
        return []
    import winreg

    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, WINDOWS_REGIONAL_SETTINGS) as key:
        separator: object = winreg.QueryValueEx(key, WINDOWS_LIST_SEPARATOR)[0]
    if not isinstance(separator, str) or not separator:
        raise OutlookError(EErrorCode.METADATA_UNAVAILABLE)
    return [name.strip() for name in raw.split(separator) if name.strip()]


def message_metadata(item: IOutlookItem) -> tuple[MessageMetadata, list[WarningInfo]]:
    result = MessageMetadata()
    warnings = []
    try:
        result.conversation_id = item.ConversationID or None
    except Exception:
        warnings.append(
            WarningInfo(code="CONVERSATION_ID_UNAVAILABLE", message="Conversation ID unavailable.")
        )
    try:
        result.categories = category_names(item)
    except Exception as error:
        warnings.append(
            WarningInfo(
                code=CATEGORIES_ACCESS_DENIED_WARNING
                if com_error(error).code == EErrorCode.ACCESS_DENIED
                else "CATEGORIES_UNAVAILABLE",
                message="Category metadata unavailable.",
            )
        )
    try:
        result.importance = IMPORTANCE_VALUES[item.Importance]
    except Exception as error:
        warnings.append(
            WarningInfo(
                code=IMPORTANCE_ACCESS_DENIED_WARNING
                if com_error(error).code == EErrorCode.ACCESS_DENIED
                else "IMPORTANCE_UNAVAILABLE",
                message="Importance unavailable.",
            )
        )
    return result, warnings


def recipient_matches(item: IOutlookItem, needle: str) -> bool:
    incomplete = False
    denied = False
    try:
        recipients = item.Recipients
        count = recipients.Count
        incomplete = count > MAX_RECIPIENTS
        for index in range(1, min(count, MAX_RECIPIENTS) + 1):
            try:
                recipient = recipients.Item(index)
            except Exception as error:
                incomplete = True
                denied = denied or com_error(error).code == EErrorCode.ACCESS_DENIED
                continue
            try:
                if needle in recipient.Name.casefold():
                    return True
            except Exception as error:
                incomplete = True
                denied = denied or com_error(error).code == EErrorCode.ACCESS_DENIED
            try:
                email = address_entry_email(recipient.AddressEntry, preserve_denials=True)
                if email is not None and needle in email.casefold():
                    return True
                incomplete = incomplete or email is None
            except Exception as error:
                incomplete = True
                denied = denied or com_error(error).code == EErrorCode.ACCESS_DENIED
    except Exception as error:
        incomplete = True
        denied = denied or com_error(error).code == EErrorCode.ACCESS_DENIED
    if incomplete:
        raise OutlookError(EErrorCode.ACCESS_DENIED if denied else EErrorCode.METADATA_UNAVAILABLE)
    return False


def attachment_matches(item: IOutlookItem, needle: str) -> bool:
    incomplete = False
    denied = False
    try:
        attachments = item.Attachments
        count = attachments.Count
        incomplete = count > MAX_ATTACHMENTS
        for index in range(1, min(count, MAX_ATTACHMENTS) + 1):
            try:
                if needle in attachments.Item(index).FileName.casefold():
                    return True
            except Exception as error:
                incomplete = True
                denied = denied or com_error(error).code == EErrorCode.ACCESS_DENIED
    except Exception as error:
        incomplete = True
        denied = denied or com_error(error).code == EErrorCode.ACCESS_DENIED
    if incomplete:
        raise OutlookError(EErrorCode.ACCESS_DENIED if denied else EErrorCode.METADATA_UNAVAILABLE)
    return False


def collection_filters_match(item: IOutlookItem, arguments: SearchArguments) -> bool:
    if arguments.recipient and not recipient_matches(item, arguments.recipient.casefold()):
        return False
    return not arguments.attachment_name or attachment_matches(
        item, arguments.attachment_name.casefold()
    )

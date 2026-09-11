"""Read-only mapping from documented Outlook properties into public models."""

from datetime import datetime

from .com_types import IAddressEntry, IOutlookItem
from .config import MAX_ATTACHMENTS, MAX_RECIPIENTS
from .enums import EErrorCode, ERecipientKind
from .errors import OutlookError, com_error
from .filters import local_timezone
from .models import (
    Attachment,
    EmailDetail,
    EmailSummary,
    Person,
    ReadArguments,
    Recipient,
    WarningInfo,
)
from .outlook_constants import (
    HEADER_ONLY,
    MAIL_ITEM_CLASS,
    RECIPIENT_KINDS,
    SMTP_ADDRESS_PROPERTY,
    SMTP_ADDRESS_TYPE,
)


def timestamp(value: datetime) -> str:
    if not isinstance(value, datetime):
        raise OutlookError(
            EErrorCode.OUTLOOK_UNAVAILABLE, "Outlook did not provide a valid message date."
        )
    if value.tzinfo is None:
        value = value.replace(tzinfo=local_timezone())
    return value.isoformat()


def smtp_address(value: object) -> str | None:
    if isinstance(value, str) and "@" in value and not value.startswith("/"):
        return value
    return None


def address_entry_email(entry: IAddressEntry | None) -> str | None:
    if entry is None:
        return None
    try:
        if entry.Type == SMTP_ADDRESS_TYPE:
            return smtp_address(entry.Address)
        for resolve_exchange in (entry.GetExchangeUser, entry.GetExchangeDistributionList):
            try:
                exchange = resolve_exchange()
                if exchange is not None:
                    address = smtp_address(exchange.PrimarySmtpAddress)
                    if address:
                        return address
            except Exception:
                continue
        return smtp_address(entry.PropertyAccessor.GetProperty(SMTP_ADDRESS_PROPERTY))
    except Exception:
        return None


def sender(item: IOutlookItem) -> tuple[Person, list[WarningInfo]]:
    warnings = []
    try:
        name = item.SenderName
    except Exception:
        name = None
        warnings.append(
            WarningInfo(code="SENDER_NAME_UNAVAILABLE", message="Sender name unavailable.")
        )
    try:
        email = (
            smtp_address(item.SenderEmailAddress)
            if item.SenderEmailType == SMTP_ADDRESS_TYPE
            else address_entry_email(item.Sender)
        )
    except Exception:
        email = None
    if email is None:
        warnings.append(
            WarningInfo(code="SMTP_UNAVAILABLE", message="Sender SMTP address unavailable.")
        )
    return Person(name=name, email=email), warnings


def summary(item: IOutlookItem, store_id: str, folder_id: str) -> EmailSummary:
    if item.Class != MAIL_ITEM_CLASS:
        raise OutlookError(EErrorCode.INVALID_ARGUMENT, "The requested item is not an email.")
    person, warnings = sender(item)
    return EmailSummary(
        entry_id=item.EntryID,
        store_id=store_id,
        folder_id=folder_id,
        subject=item.Subject,
        sender=person,
        received_at=timestamp(item.ReceivedTime),
        unread=item.UnRead,
        has_attachments=item.Attachments.Count > 0,
        warnings=warnings,
    )


def plain_body(item: IOutlookItem) -> str:
    download_state: int | None
    try:
        download_state = item.DownloadState
    except Exception:
        download_state = None
    if download_state == HEADER_ONLY:
        raise OutlookError(EErrorCode.BODY_UNAVAILABLE)
    try:
        body = item.Body
        if not isinstance(body, str) or (not body and download_state is None):
            raise OutlookError(EErrorCode.BODY_UNAVAILABLE)
        return body
    except OutlookError:
        raise
    except Exception as error:
        mapped = com_error(error)
        if mapped.code == EErrorCode.ACCESS_DENIED:
            raise mapped from None
        raise OutlookError(EErrorCode.BODY_UNAVAILABLE) from None


def detail(item: IOutlookItem, arguments: ReadArguments) -> EmailDetail:
    email = summary(item, arguments.store_id, item.Parent.EntryID)
    body = plain_body(item)
    if arguments.body_offset > len(body):
        raise OutlookError(EErrorCode.INVALID_ARGUMENT, "body_offset exceeds the body length.")
    end = min(arguments.body_offset + arguments.body_limit, len(body))
    warnings = list(email.warnings)
    recipients: list[Recipient] = []
    attachments: list[Attachment] = []
    omitted_recipients = omitted_attachments = 0
    try:
        recipient_collection = item.Recipients
        omitted_recipients = max(0, recipient_collection.Count - MAX_RECIPIENTS)
        for index in range(1, min(recipient_collection.Count, MAX_RECIPIENTS) + 1):
            try:
                recipient = recipient_collection.Item(index)
                kind = RECIPIENT_KINDS.get(recipient.Type, ERecipientKind.UNKNOWN)
                address = address_entry_email(recipient.AddressEntry)
                if address is None:
                    warnings.append(
                        WarningInfo(
                            code="SMTP_UNAVAILABLE",
                            message="A recipient SMTP address is unavailable.",
                        )
                    )
                recipients.append(Recipient(name=recipient.Name, email=address, kind=kind))
            except Exception:
                omitted_recipients += 1
    except Exception:
        warnings.append(
            WarningInfo(
                code="RECIPIENTS_UNAVAILABLE",
                message="Outlook did not expose the recipient collection; its size is unknown.",
            )
        )
    try:
        attachment_collection = item.Attachments
        omitted_attachments = max(0, attachment_collection.Count - MAX_ATTACHMENTS)
        for index in range(1, min(attachment_collection.Count, MAX_ATTACHMENTS) + 1):
            try:
                attachment = attachment_collection.Item(index)
                attachments.append(Attachment(name=attachment.FileName, size=attachment.Size))
            except Exception:
                omitted_attachments += 1
    except Exception:
        warnings.append(
            WarningInfo(
                code="ATTACHMENTS_UNAVAILABLE",
                message="Outlook did not expose attachment metadata; its count is unknown.",
            )
        )
    try:
        sent_at = timestamp(item.SentOn)
    except Exception:
        sent_at = None
        warnings.append(WarningInfo(code="SENT_DATE_UNAVAILABLE", message="Sent date unavailable."))
    if omitted_recipients or omitted_attachments:
        warnings.append(
            WarningInfo(
                code="METADATA_OMITTED",
                message="Some recipient or attachment metadata was omitted; see the counts.",
            )
        )
    # Avoid repeating the same warning for a large distribution list.
    unique = list({warning.code + warning.message: warning for warning in warnings}.values())
    return EmailDetail(
        **email.model_dump(exclude={"warnings"}),
        warnings=unique,
        recipients=recipients,
        sent_at=sent_at,
        attachments=attachments,
        body=body[arguments.body_offset : end],
        body_offset=arguments.body_offset,
        next_body_offset=end if end < len(body) else None,
        body_truncated=end < len(body),
        body_length=len(body),
        omitted_recipients=omitted_recipients,
        omitted_attachments=omitted_attachments,
    )

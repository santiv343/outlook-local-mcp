"""One-use in-memory draft revisions and explicitly enabled Outlook submission."""

import hashlib
import secrets
import time
from collections.abc import Callable
from typing import TYPE_CHECKING

from .accounts import (
    account_person,
    draft_account,
    resolved_recipients,
    select_account,
    set_sending_account,
)
from .action_models import (
    DraftReview,
    PreviewArguments,
    SendArguments,
    SendConfirmation,
    SendPreview,
    SendResult,
)
from .addresses import address_entry_email, explicit_smtp
from .com_types import IOutlookItem
from .config import (
    CURSOR_RANDOM_BYTES,
    MAX_DRAFT_BODY_CHARACTERS,
    MAX_SEND_PREVIEWS,
    SEND_PREVIEW_TTL_SECONDS,
)
from .enums import EErrorCode
from .errors import OutlookError, com_error
from .mail import plain_body, timestamp
from .models import Person
from .outlook_constants import PLAIN_TEXT_FORMAT, REPRESENTING_SMTP_PROPERTY

if TYPE_CHECKING:
    from .outlook import Outlook


class PreviewStore:
    def __init__(self, clock: Callable[[], float] = time.monotonic) -> None:
        self.clock = clock
        self.previews: dict[str, SendConfirmation] = {}

    def clear(self) -> None:
        self.previews.clear()

    def discard(self, entry_id: str, store_id: str) -> None:
        self.previews = {
            token: value
            for token, value in self.previews.items()
            if (value.entry_id, value.store_id) != (entry_id, store_id)
        }

    def put(self, arguments: PreviewArguments, digest: str) -> str:
        self.discard(arguments.entry_id, arguments.store_id)
        self.previews = {
            token: value
            for token, value in self.previews.items()
            if self.clock() - value.created < SEND_PREVIEW_TTL_SECONDS
        }
        if len(self.previews) >= MAX_SEND_PREVIEWS:
            raise OutlookError(
                EErrorCode.SEARCH_SESSION_LIMIT, "Too many active send previews; wait for expiry."
            )
        token = secrets.token_urlsafe(CURSOR_RANDOM_BYTES)
        self.previews[token] = SendConfirmation(
            entry_id=arguments.entry_id,
            store_id=arguments.store_id,
            digest=digest,
            created=self.clock(),
            account_email=arguments.account_email,
        )
        return token

    def take(self, arguments: SendArguments) -> SendConfirmation:
        value = self.previews.get(arguments.confirmation_token)
        if value is None or (value.entry_id, value.store_id) != (
            arguments.entry_id,
            arguments.store_id,
        ):
            raise OutlookError(EErrorCode.CONFIRMATION_INVALID)
        self.discard(value.entry_id, value.store_id)
        if self.clock() - value.created >= SEND_PREVIEW_TTL_SECONDS:
            raise OutlookError(EErrorCode.CONFIRMATION_INVALID)
        return value


def represented_smtp(item: IOutlookItem) -> str | None:
    try:
        value = item.PropertyAccessor.GetProperty(REPRESENTING_SMTP_PROPERTY)
    except Exception as error:
        mapped = com_error(error, EErrorCode.ITEM_NOT_FOUND)
        if mapped.code == EErrorCode.ITEM_NOT_FOUND:
            return None
        if mapped.code == EErrorCode.ACCESS_DENIED:
            raise mapped from None
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION, "The represented From identity is inaccessible."
        ) from None
    if value is None or value == "":
        return None
    try:
        if not isinstance(value, str):
            raise ValueError
        return explicit_smtp(value)
    except ValueError:
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION, "The represented From SMTP address is invalid."
        ) from None


def review(
    backend: "Outlook",
    item: IOutlookItem,
    account_email: str,
    *,
    require_saved: bool = True,
    require_native_account: bool = False,
) -> DraftReview:
    account = account_person(
        draft_account(
            backend.namespace(),
            item,
            account_email,
            require_saved=require_saved,
            require_native_account=require_native_account,
        )
    )
    if item.BodyFormat != PLAIN_TEXT_FORMAT or item.Attachments.Count:
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION,
            "Programmatic sending supports plain-text drafts without attachments. "
            "Review and send this draft in Outlook.",
        )
    body = plain_body(item)
    if len(body) > MAX_DRAFT_BODY_CHARACTERS:
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION,
            "This body exceeds the complete preview limit. Review and send it in Outlook.",
        )
    represented_name = item.SentOnBehalfOfName
    represented_address = represented_smtp(item)
    if represented_name:
        recipient = backend.namespace().CreateRecipient(represented_name)
        if not recipient.Resolve():
            raise OutlookError(
                EErrorCode.UNSUPPORTED_COMPOSITION, "The represented From identity is unresolved."
            )
        resolved = address_entry_email(recipient.AddressEntry)
        if (
            resolved is None
            or account.email is None
            or resolved.casefold() != account.email.casefold()
        ):
            raise OutlookError(
                EErrorCode.UNSUPPORTED_COMPOSITION,
                "The represented From identity differs from the sending account.",
            )
        represented_address = represented_address or resolved
    if (represented_name or represented_address) and (
        represented_address is None
        or account.email is None
        or represented_address.casefold() != account.email.casefold()
    ):
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION,
            "Delegated or unresolved From overrides require review and sending in Outlook.",
        )
    sender = Person(
        name=represented_name or account.name, email=represented_address or account.email
    )
    return DraftReview(
        account=account,
        sender=sender,
        recipients=resolved_recipients(item),
        subject=item.Subject,
        body=body,
        draft_modified_at=timestamp(item.LastModificationTime),
    )


def review_digest(value: DraftReview) -> str:
    return hashlib.sha256(value.model_dump_json().encode("utf-8")).hexdigest()


def prepare_send(backend: "Outlook", arguments: PreviewArguments) -> SendPreview:
    value = review(
        backend, backend.item(arguments.entry_id, arguments.store_id), arguments.account_email
    )
    digest = review_digest(value)
    token = backend.previews.put(arguments, digest)
    return SendPreview(
        **value.model_dump(),
        entry_id=arguments.entry_id,
        store_id=arguments.store_id,
        confirmation_token=token,
        digest=digest,
        expires_in_seconds=SEND_PREVIEW_TTL_SECONDS,
    )


def send_draft(backend: "Outlook", arguments: SendArguments) -> SendResult:
    confirmation = backend.previews.take(arguments)
    item = backend.item(arguments.entry_id, arguments.store_id)
    approved = review(backend, item, confirmation.account_email)
    if review_digest(approved) != confirmation.digest:
        raise OutlookError(EErrorCode.DRAFT_CHANGED)
    account = select_account(backend.namespace(), arguments.store_id, confirmation.account_email)
    backend.mutation_attempted = True
    set_sending_account(item, account)
    assigned = review(
        backend,
        item,
        confirmation.account_email,
        require_saved=False,
        require_native_account=True,
    )
    if assigned.model_dump(exclude={"draft_modified_at"}) != approved.model_dump(
        exclude={"draft_modified_at"}
    ):
        raise OutlookError(EErrorCode.DRAFT_CHANGED)
    item.Send()
    return SendResult()

"""Explicit Outlook UI and draft mutations; never send from these operations."""

from typing import TYPE_CHECKING

from .accounts import (
    account_person,
    draft_account,
    drafts_folder,
    resolved_recipients,
    select_account,
    set_sending_account,
)
from .action_models import DraftArguments, DraftResult, ItemArguments, OpenResult, ReplyArguments
from .com_types import IOutlookItem
from .config import DEFAULT_BODY_CHARACTERS
from .enums import EErrorCode
from .errors import OutlookError
from .mail import body_page, plain_body
from .outlook_constants import MAIL_MESSAGE_CLASS, PLAIN_TEXT_FORMAT, RECIPIENT_KINDS

if TYPE_CHECKING:
    from .outlook import Outlook


def open_email(backend: "Outlook", arguments: ItemArguments) -> OpenResult:
    item = backend.item(arguments.entry_id, arguments.store_id)
    backend.mutation_attempted = True
    item.Display(False)
    item.GetInspector.Activate()
    return OpenResult(entry_id=item.EntryID, store_id=item.Parent.StoreID)


def draft_result(backend: "Outlook", item: IOutlookItem) -> DraftResult:
    account = draft_account(backend.namespace(), item)
    return DraftResult(
        entry_id=item.EntryID,
        store_id=item.Parent.StoreID,
        folder_id=item.Parent.EntryID,
        account=account_person(account),
        recipients=resolved_recipients(item),
        subject=item.Subject,
        **body_page(plain_body(item), 0, DEFAULT_BODY_CHARACTERS).model_dump(),
    )


def create_draft(backend: "Outlook", arguments: DraftArguments) -> DraftResult:
    account = select_account(backend.namespace(), arguments.store_id, arguments.account_email)
    folder = drafts_folder(account)
    backend.mutation_attempted = True
    draft = folder.Items.Add(MAIL_MESSAGE_CLASS)
    set_sending_account(draft, account)
    draft.BodyFormat = PLAIN_TEXT_FORMAT
    draft.Subject = arguments.subject
    draft.Body = arguments.body
    recipients = {"to": arguments.to, "cc": arguments.cc, "bcc": arguments.bcc}
    for native_kind, kind in RECIPIENT_KINDS.items():
        for address in recipients[kind.value]:
            draft.Recipients.Add(address).Type = native_kind
    if not draft.Recipients.ResolveAll():
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION, "Outlook could not resolve every recipient."
        )
    resolved_recipients(draft)
    draft.Save()
    return draft_result(backend, draft)


def reply_to_email(backend: "Outlook", arguments: ReplyArguments) -> DraftResult:
    original = backend.item(arguments.entry_id, arguments.store_id)
    explicit = (
        select_account(backend.namespace(), account_email=arguments.account_email)
        if arguments.account_email is not None
        else None
    )
    backend.mutation_attempted = True
    draft = original.ReplyAll() if arguments.reply_all else original.Reply()
    account = explicit or draft.SendUsingAccount
    if account is None or account.DeliveryStore is None:
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION,
            "Outlook did not select a reply account; specify account_email.",
        )
    account = select_account(
        backend.namespace(), account.DeliveryStore.StoreID, account.SmtpAddress
    )
    set_sending_account(draft, account)
    quote = plain_body(draft)
    draft.BodyFormat = PLAIN_TEXT_FORMAT
    draft.Body = arguments.body + "\r\n\r\n" + quote
    if not draft.Recipients.ResolveAll():
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION, "Outlook could not resolve every reply recipient."
        )
    resolved_recipients(draft)
    draft.Save()
    return draft_result(backend, draft)

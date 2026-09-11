"""Account selection and complete, resolved draft recipients."""

from .addresses import address_entry_email, explicit_smtp
from .com_types import IAccount, IFolder, INamespace, IOutlookItem
from .config import MAX_DRAFT_RECIPIENTS
from .enums import EErrorCode
from .errors import OutlookError
from .models import Person, Recipient
from .outlook_constants import (
    COM_NEUTRAL_LOCALE,
    DRAFTS_FOLDER,
    RECIPIENT_KINDS,
    SENDING_ACCOUNT_PROPERTY,
)


def set_sending_account(item: IOutlookItem, account: IAccount) -> None:
    import pythoncom

    # Outlook requires a reference assignment. Its type library advertises PUT,
    # which pywin32's ordinary property assignment can silently ignore. Resolve
    # the member by name and use COM's PUTREF; never hardcode an Outlook DISPID.
    dispatch = item._oleobj_
    member = dispatch.GetIDsOfNames(SENDING_ACCOUNT_PROPERTY)
    dispatch.Invoke(member, COM_NEUTRAL_LOCALE, pythoncom.DISPATCH_PROPERTYPUTREF, False, account)
    actual = item.SendUsingAccount
    if (
        actual is None
        or actual.DeliveryStore is None
        or account.DeliveryStore is None
        or (
            actual.SmtpAddress.casefold() != account.SmtpAddress.casefold()
            or actual.DeliveryStore.StoreID != account.DeliveryStore.StoreID
        )
    ):
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION,
            "Outlook did not retain the selected sending account.",
        )


def account_person(account: IAccount) -> Person:
    try:
        address = explicit_smtp(account.SmtpAddress)
    except ValueError:
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION, "Account SMTP address unavailable."
        ) from None
    return Person(name=account.DisplayName, email=address)


def select_account(
    namespace: INamespace, store_id: str | None = None, account_email: str | None = None
) -> IAccount:
    selected_store = store_id
    if selected_store is None and account_email is None:
        selected_store = namespace.DefaultStore.StoreID
    matches = []
    accounts = namespace.Accounts
    for index in range(1, accounts.Count + 1):
        account = accounts.Item(index)
        store = account.DeliveryStore
        if store is None or (selected_store is not None and store.StoreID != selected_store):
            continue
        if account_email is not None and account.SmtpAddress.casefold() != account_email.casefold():
            continue
        matches.append(account)
    if len(matches) != 1:
        raise OutlookError(
            EErrorCode.INVALID_ARGUMENT,
            "Select exactly one sending account with account_email and its delivery store.",
        )
    account_person(matches[0])
    return matches[0]


def drafts_folder(account: IAccount) -> IFolder:
    store = account.DeliveryStore
    folder = store.GetDefaultFolder(DRAFTS_FOLDER) if store is not None else None
    if folder is None:
        raise OutlookError(
            EErrorCode.FOLDER_NOT_FOUND, "This account has no accessible Drafts folder."
        )
    return folder


def draft_account(
    namespace: INamespace,
    item: IOutlookItem,
    account_email: str | None = None,
    *,
    require_saved: bool = True,
    require_native_account: bool = False,
) -> IAccount:
    account = item.SendUsingAccount
    if account is None and (require_native_account or account_email is None):
        raise OutlookError(EErrorCode.UNSUPPORTED_COMPOSITION, "The draft has no sending account.")
    selected = select_account(
        namespace,
        item.Parent.StoreID,
        account_email or (account.SmtpAddress if account is not None else None),
    )
    if account is not None and (
        account.DeliveryStore is None
        or account.SmtpAddress.casefold() != selected.SmtpAddress.casefold()
        or account.DeliveryStore.StoreID != item.Parent.StoreID
    ):
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION,
            "The native draft account conflicts with the selected account.",
        )
    folder = drafts_folder(selected)
    parent = item.Parent
    if (
        (require_saved and not item.Saved)
        or item.Sent
        or parent.EntryID != folder.EntryID
        or parent.StoreID != folder.StoreID
    ):
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION,
            "Use a saved, unsent draft in its account's Drafts folder.",
        )
    return selected


def resolved_recipients(item: IOutlookItem) -> list[Recipient]:
    collection = item.Recipients
    if not 1 <= collection.Count <= MAX_DRAFT_RECIPIENTS:
        raise OutlookError(
            EErrorCode.UNSUPPORTED_COMPOSITION, "Recipient count is outside the supported range."
        )
    recipients = []
    for index in range(1, collection.Count + 1):
        recipient = collection.Item(index)
        address = address_entry_email(recipient.AddressEntry)
        kind = RECIPIENT_KINDS.get(recipient.Type)
        if not recipient.Resolved or address is None or kind is None:
            raise OutlookError(
                EErrorCode.UNSUPPORTED_COMPOSITION,
                "Resolve every recipient to an SMTP address in Outlook.",
            )
        try:
            address = explicit_smtp(address)
        except ValueError:
            raise OutlookError(
                EErrorCode.UNSUPPORTED_COMPOSITION, "A recipient SMTP address is invalid."
            ) from None
        recipients.append(Recipient(name=recipient.Name, email=address, kind=kind))
    return recipients

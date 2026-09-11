"""Classic Outlook adapter. Used only on the worker's main STA thread."""

import sys

from .com_types import IFolder, INamespace, IOutlookApplication, IOutlookItem, IStore
from .enums import EErrorCode, EToolName
from .errors import OutlookError, com_error
from .filters import date_range, metadata_matches, restrict_filter
from .mail import detail, plain_body, summary
from .models import (
    EmailDetail,
    EmailSummary,
    Folder,
    FolderArguments,
    Mailbox,
    Mailboxes,
    OutlookStatus,
    Page,
    ReadArguments,
    RecentArguments,
    SearchArguments,
)
from .outlook_constants import (
    INBOX_FOLDER,
    MAIL_ITEM_CLASS,
    MAPI_NAMESPACE,
    OUTLOOK_PROGRAM_ID,
    OUTLOOK_REGISTRATION,
    RECEIVED_TIME_PROPERTY,
)
from .pagination import CursorStore, FolderSession, MailSession, scan_page, signature


class Outlook:
    def __init__(self, application: IOutlookApplication | None = None) -> None:
        self.application = application
        self.cursors = CursorStore()

    def connect(self) -> IOutlookApplication:
        if self.application is not None:
            return self.application
        if sys.platform != "win32":
            raise OutlookError(EErrorCode.UNSUPPORTED_PLATFORM)
        import winreg

        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, OUTLOOK_REGISTRATION):
                pass
        except FileNotFoundError:
            raise OutlookError(EErrorCode.OUTLOOK_NOT_INSTALLED) from None
        try:
            import win32com.client

            # Dispatching this existing object does not launch Outlook or log on.
            application: IOutlookApplication = win32com.client.GetActiveObject(OUTLOOK_PROGRAM_ID)
            self.application = application
            return application
        except Exception as error:
            raise com_error(error) from None

    def namespace(self) -> INamespace:
        return self.connect().GetNamespace(MAPI_NAMESPACE)

    def close(self) -> None:
        self.cursors.clear()
        self.application = None

    def status(self) -> OutlookStatus:
        try:
            application = self.connect()
            namespace = application.GetNamespace(MAPI_NAMESPACE)
            count = namespace.Stores.Count
            if count == 0:
                raise OutlookError(
                    EErrorCode.OUTLOOK_UNAVAILABLE, "Open Outlook and finish setting up a profile."
                )
            return OutlookStatus(available=True, version=application.Version, store_count=count)
        except Exception as error:
            mapped = error if isinstance(error, OutlookError) else com_error(error)
            self.close()
            return OutlookStatus(available=False, error=mapped.info())

    def mailboxes(self) -> Mailboxes:
        namespace = self.namespace()
        default_id = namespace.DefaultStore.StoreID
        stores = namespace.Stores
        return Mailboxes(
            items=[
                Mailbox(
                    store_id=store.StoreID,
                    name=store.DisplayName,
                    is_default=store.StoreID == default_id,
                )
                for store in (stores.Item(index) for index in range(1, stores.Count + 1))
            ]
        )

    def store(self, store_id: str) -> IStore:
        try:
            return self.namespace().GetStoreFromID(store_id)
        except Exception as error:
            raise com_error(error, EErrorCode.STORE_NOT_FOUND) from None

    def folder(self, store_id: str | None, folder_id: str | None, root: bool = False) -> IFolder:
        namespace = self.namespace()
        store = self.store(store_id) if store_id else namespace.DefaultStore
        try:
            if folder_id:
                folder = namespace.GetFolderFromID(folder_id, store_id)
                if folder.StoreID != store.StoreID:
                    raise OutlookError(EErrorCode.FOLDER_NOT_FOUND)
                return folder
            selected = store.GetRootFolder() if root else store.GetDefaultFolder(INBOX_FOLDER)
            if selected is None:
                raise OutlookError(
                    EErrorCode.FOLDER_NOT_FOUND,
                    "This mailbox has no default Inbox. Select a folder explicitly.",
                )
            return selected
        except OutlookError:
            raise
        except Exception as error:
            raise com_error(error, EErrorCode.FOLDER_NOT_FOUND) from None

    def folders(self, arguments: FolderArguments) -> Page[Folder]:
        key = signature(EToolName.LIST_FOLDERS, arguments.model_dump())
        if arguments.cursor:
            restored = self.cursors.take(arguments.cursor, key)
            if not isinstance(restored, FolderSession):
                raise OutlookError(EErrorCode.INTERNAL_ERROR)
            session = restored
        else:
            self.cursors.check_capacity()
            parent = self.folder(arguments.store_id, arguments.parent_folder_id, root=True)
            session = FolderSession(
                parent.Folders, parent.StoreID, parent.EntryID, key, self.cursors.clock()
            )

        def project(folder: IFolder) -> Folder:
            return Folder(
                folder_id=folder.EntryID,
                store_id=folder.StoreID,
                name=folder.Name,
                has_children=folder.Folders.Count > 0,
            )

        return scan_page(self.cursors, session, arguments.limit, project)

    def emails(self, arguments: RecentArguments | SearchArguments) -> Page[EmailSummary]:
        searching = isinstance(arguments, SearchArguments)
        filters = (
            arguments
            if isinstance(arguments, SearchArguments)
            else SearchArguments(**arguments.model_dump())
        )
        after, before = date_range(filters)
        operation = EToolName.SEARCH_EMAILS if searching else EToolName.RECENT_EMAILS
        key = signature(operation, arguments.model_dump())
        if arguments.cursor:
            restored = self.cursors.take(arguments.cursor, key)
            if not isinstance(restored, MailSession):
                raise OutlookError(EErrorCode.INTERNAL_ERROR)
            session = restored
        else:
            self.cursors.check_capacity()
            folder = self.folder(arguments.store_id, arguments.folder_id)
            collection = folder.Items
            restriction = restrict_filter(after, before)
            if restriction:
                collection = collection.Restrict(restriction)
            collection.Sort(RECEIVED_TIME_PROPERTY, True)
            session = MailSession(
                collection, folder.StoreID, folder.EntryID, key, self.cursors.clock()
            )

        def project(item: IOutlookItem) -> EmailSummary | None:
            if item.Class != MAIL_ITEM_CLASS:
                return None
            email = summary(item, session.store_id, session.folder_id)
            if not metadata_matches(email, filters, after, before):
                return None
            needle = filters.query.casefold()
            if needle:
                subject_match = needle in email.subject.casefold()
                if filters.query_in == "subject" and not subject_match:
                    return None
                needs_body = filters.query_in == "body" or (
                    filters.query_in == "subject_body" and not subject_match
                )
                if needs_body and needle not in plain_body(item).casefold():
                    return None
            return email

        return scan_page(self.cursors, session, arguments.limit, project)

    def read(self, arguments: ReadArguments) -> EmailDetail:
        self.store(arguments.store_id)
        try:
            item = self.namespace().GetItemFromID(arguments.entry_id, arguments.store_id)
        except Exception as error:
            raise com_error(error, EErrorCode.ITEM_NOT_FOUND) from None
        if item is None:
            raise OutlookError(EErrorCode.ITEM_NOT_FOUND)
        if item.Parent.StoreID != arguments.store_id:
            raise OutlookError(EErrorCode.ITEM_NOT_FOUND)
        return detail(item, arguments)

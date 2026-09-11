"""Native conversation tables adapted to the existing bounded cursor engine."""

import json
from typing import TYPE_CHECKING

from .action_models import ConversationArguments
from .com_types import ConversationItem, INamespace, ITable
from .enums import EErrorCode, EToolName
from .errors import OutlookError, com_error
from .mail import detail
from .models import EmailDetail, Page, ReadArguments
from .outlook_constants import MAIL_ITEM_CLASS, STORE_ENTRY_ID_PROPERTY
from .pagination import ConversationSession, scan_page, signature

if TYPE_CHECKING:
    from .outlook import Outlook


class ConversationRows:
    def __init__(self, table: ITable, namespace: INamespace) -> None:
        self.table = table
        self.namespace = namespace
        self.table.Columns.Add(STORE_ENTRY_ID_PROPERTY)

    def GetFirst(self) -> ConversationItem | None:
        self.table.MoveToStart()
        return self.GetNext()

    def GetNext(self) -> ConversationItem | None:
        if self.table.EndOfTable:
            return None
        row = self.table.GetNextRow()
        entry_id = row.Item("EntryID")
        store_id = row.BinaryToString(STORE_ENTRY_ID_PROPERTY)
        if not isinstance(entry_id, str) or not entry_id or not store_id:
            raise OutlookError(EErrorCode.METADATA_UNAVAILABLE)
        item = self.namespace.GetItemFromID(entry_id, store_id)
        if item is None or item.Parent.StoreID != store_id:
            raise OutlookError(EErrorCode.ITEM_NOT_FOUND)
        return ConversationItem(json.dumps([store_id, entry_id]), item, store_id)


def read_conversation(backend: "Outlook", arguments: ConversationArguments) -> Page[EmailDetail]:
    key = signature(EToolName.READ_CONVERSATION, arguments.model_dump())
    if arguments.cursor:
        restored = backend.cursors.take(arguments.cursor, key)
        if not isinstance(restored, ConversationSession):
            raise OutlookError(EErrorCode.INTERNAL_ERROR)
        session = restored
    else:
        backend.cursors.check_capacity()
        store = backend.store(arguments.store_id)
        item = backend.item(arguments.entry_id, arguments.store_id)
        try:
            if not store.IsConversationEnabled:
                raise OutlookError(EErrorCode.CONVERSATION_UNAVAILABLE)
            native = item.GetConversation()
            if native is None:
                raise OutlookError(EErrorCode.CONVERSATION_UNAVAILABLE)
            rows = ConversationRows(native.GetTable(), backend.namespace())
        except OutlookError:
            raise
        except Exception as error:
            mapped = com_error(error)
            if mapped.code == EErrorCode.ACCESS_DENIED:
                raise mapped from None
            raise OutlookError(EErrorCode.CONVERSATION_UNAVAILABLE) from None
        session = ConversationSession(
            rows, arguments.store_id, item.Parent.EntryID, key, backend.cursors.clock()
        )

    def project(value: ConversationItem) -> EmailDetail | None:
        if value.item.Class != MAIL_ITEM_CLASS:
            return None
        return detail(
            value.item,
            ReadArguments(
                entry_id=value.item.EntryID,
                store_id=value.store_id,
                body_limit=arguments.body_limit,
            ),
        )

    return scan_page(backend.cursors, session, arguments.limit, project)

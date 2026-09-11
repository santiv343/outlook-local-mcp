from types import SimpleNamespace

import pytest

from outlook_local_mcp.action_models import ConversationArguments
from outlook_local_mcp.conversation import read_conversation
from outlook_local_mcp.enums import EErrorCode
from outlook_local_mcp.errors import OutlookError
from outlook_local_mcp.outlook import Outlook

from .action_fakes import Application, Draft, Table


def test_native_cross_store_paging_deduplicates_composite_ids_and_skips_missing():
    backend = Outlook(Application())
    namespace = backend.namespace()
    first = Draft(namespace.DefaultStore.inbox)
    second = Draft(namespace.secondary.inbox)
    first.EntryID = second.EntryID = "same-id-in-two-stores"
    nonmail = Draft(namespace.secondary.inbox)
    nonmail.EntryID, nonmail.Class = "appointment", 26
    namespace.DefaultStore.inbox.Items.items.append(first)
    namespace.secondary.inbox.Items.items.extend([second, nonmail])
    identities = [
        ("default", first.EntryID),
        ("default", first.EntryID),
        ("secondary", "deleted"),
        ("secondary", second.EntryID),
        ("secondary", nonmail.EntryID),
    ]
    table = Table(identities)
    first.GetConversation = lambda: SimpleNamespace(GetTable=lambda: table)
    arguments = ConversationArguments(
        entry_id=first.EntryID, store_id="default", limit=1, body_limit=3
    )
    page = read_conversation(backend, arguments)
    assert page.items[0].store_id == "default" and page.items[0].body_truncated
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.INVALID_ARGUMENT):
        read_conversation(
            backend, arguments.model_copy(update={"cursor": page.next_cursor, "body_limit": 4})
        )
    next_page = read_conversation(
        backend, arguments.model_copy(update={"cursor": page.next_cursor, "limit": 2})
    )
    assert [item.store_id for item in next_page.items] == ["secondary"]
    assert next_page.items[0].folder_id == namespace.secondary.inbox.EntryID
    assert next_page.store_id == "default" and next_page.coverage.exhausted
    assert next_page.omitted == 1 and not next_page.coverage.evaluation_complete
    assert table.starts == 1 and first.UnRead and second.UnRead


@pytest.mark.parametrize("enabled", [False, True])
def test_unsupported_or_missing_conversation_is_explicit(enabled):
    backend = Outlook(Application())
    store = backend.namespace().DefaultStore
    store.IsConversationEnabled = enabled
    mail = Draft(store.inbox)
    mail.EntryID = "anchor"
    mail.GetConversation = lambda: None
    store.inbox.Items.items.append(mail)
    with pytest.raises(
        OutlookError, check=lambda error: error.code == EErrorCode.CONVERSATION_UNAVAILABLE
    ):
        read_conversation(
            backend, ConversationArguments(entry_id=mail.EntryID, store_id=store.StoreID)
        )

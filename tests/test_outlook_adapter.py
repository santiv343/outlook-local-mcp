from dataclasses import FrozenInstanceError, replace
from types import SimpleNamespace

import pytest

from outlook_local_mcp.enums import EErrorCode
from outlook_local_mcp.errors import OutlookError, com_error
from outlook_local_mcp.mail import plain_body, sender
from outlook_local_mcp.models import (
    FolderArguments,
    ReadArguments,
    RecentArguments,
    SearchArguments,
)
from outlook_local_mcp.outlook import Outlook

from .fakes import Application, Collection, ComFailure, Mail


def test_mailboxes_folders_and_default_inbox_use_ids():
    app = Application([Mail()])
    outlook = Outlook(app)
    assert outlook.status().available
    assert outlook.mailboxes().items[0].is_default
    assert outlook.folders(FolderArguments(store_id="store")).items[0].folder_id == "inbox"
    assert outlook.emails(RecentArguments()).items[0].store_id == "store"
    app.namespace.DefaultStore.inbox = None
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.FOLDER_NOT_FOUND):
        outlook.emails(RecentArguments(store_id="store"))


@pytest.mark.parametrize("query", ["O'BRIEN'S", "[draft]", "quarterly"])
def test_literal_subject_search_does_not_put_user_text_in_restrict(query):
    mail = Mail()
    app = Application([mail, replace(mail, EntryID="appointment", Class=26)])
    outlook = Outlook(app)
    page = outlook.emails(SearchArguments(query=query))
    assert [item.entry_id for item in page.items] == [mail.EntryID]
    assert app.namespace.inbox.Items.restrictions == []
    assert "body" not in page.items[0].model_dump()


def test_body_search_is_explicit_and_combines_filters():
    outlook = Outlook(Application([Mail(), Mail(EntryID="other", Body="different")]))
    page = outlook.emails(SearchArguments(query="EMOJI", query_in="body", unread=True))
    assert len(page.items) == 1
    assert not outlook.emails(SearchArguments(query="EMOJI", query_in="subject")).items
    assert not outlook.emails(SearchArguments(query="EMOJI", query_in="body", unread=False)).items


def test_exchange_resolution_and_internal_address_warning():
    exchange = SimpleNamespace(
        Type="EX",
        GetExchangeUser=lambda: SimpleNamespace(PrimarySmtpAddress="exchange@example.com"),
        GetExchangeDistributionList=lambda: None,
    )
    mail = Mail(
        SenderEmailType="EX", SenderEmailAddress="/o=internal/ou=synthetic", Sender=exchange
    )
    person, warnings = sender(mail)
    assert person.email == "exchange@example.com" and not warnings
    person, warnings = sender(replace(mail, Sender=None))
    assert person.email is None and person.name == mail.SenderName
    assert warnings[0].code == "SMTP_UNAVAILABLE"


def test_denied_sender_property_does_not_leak_exception():
    class InaccessibleSender:
        SenderName = "Synthetic name"

        @property
        def SenderEmailType(self):
            raise ComFailure()

    person, warnings = sender(InaccessibleSender())
    assert person.email is None and warnings
    assert "private-error" not in str(warnings)
    assert com_error(ComFailure()).code == EErrorCode.ACCESS_DENIED


def test_body_unavailable_is_distinct_from_empty_and_denied():
    assert plain_body(Mail(Body="")) == ""
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.BODY_UNAVAILABLE):
        plain_body(Mail(Body="", DownloadState=0))

    class DeniedBody:
        DownloadState = 1

        @property
        def Body(self):
            raise ComFailure()

    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.ACCESS_DENIED):
        plain_body(DeniedBody())


def test_read_body_continuation_and_metadata_never_mutate_message():
    mail = Mail(Attachments=Collection([SimpleNamespace(FileName="example.txt", Size=42)]))
    outlook = Outlook(Application([mail]))
    arguments = ReadArguments(store_id="store", entry_id=mail.EntryID, body_limit=5)
    first = outlook.read(arguments)
    second = outlook.read(arguments.model_copy(update={"body_offset": first.next_body_offset}))
    assert first.body + second.body == mail.Body[:10]
    assert first.body_truncated and first.attachments[0].size == 42
    assert mail.UnRead is True
    with pytest.raises(FrozenInstanceError):
        mail.UnRead = False
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.INVALID_ARGUMENT):
        outlook.read(arguments.model_copy(update={"body_offset": len(mail.Body) + 1}))


def test_missing_and_non_mail_items_have_explicit_errors():
    outlook = Outlook(Application([Mail(Class=26)]))
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.ITEM_NOT_FOUND):
        outlook.read(ReadArguments(store_id="store", entry_id="deleted"))
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.INVALID_ARGUMENT):
        outlook.read(ReadArguments(store_id="store", entry_id="synthetic-message"))

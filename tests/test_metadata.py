from contextlib import nullcontext
from dataclasses import replace
from types import SimpleNamespace

import pytest

from outlook_local_mcp.models import SearchArguments
from outlook_local_mcp.outlook import Outlook

from .fakes import Application, Collection, ComFailure, Mail


def test_combined_literal_metadata_filters_and_windows_category_separator(monkeypatch):
    import winreg

    monkeypatch.setattr(winreg, "OpenKey", lambda *args: nullcontext(None))
    monkeypatch.setattr(winreg, "QueryValueEx", lambda *args: (";", winreg.REG_SZ))
    recipient = SimpleNamespace(
        Name="O'Brien, Example",
        AddressEntry=SimpleNamespace(Type="SMTP", Address="person@example.com"),
    )
    mail = Mail(
        Categories="Budget, planning; Follow up",
        Importance=2,
        Recipients=Collection([recipient]),
        Attachments=Collection([SimpleNamespace(FileName="Quarterly [final].PDF", Size=10)]),
    )
    backend = Outlook(Application([mail]))
    arguments = SearchArguments(
        query="report",
        sender="sender@example.com",
        recipient="o'BRIEN",
        category="budget, PLANNING",
        importance="high",
        attachment_name="[FINAL].pdf",
        unread=True,
        has_attachments=True,
    )
    page = backend.emails(arguments)
    assert len(page.items) == 1 and page.items[0].categories == ["Budget, planning", "Follow up"]
    assert page.items[0].conversation_id == mail.ConversationID
    assert not backend.emails(arguments.model_copy(update={"importance": "low"})).items
    assert not backend.emails(arguments.model_copy(update={"category": "Budget"})).items
    assert backend.emails(arguments.model_copy(update={"recipient": "PERSON@"})).items


@pytest.mark.parametrize("field", ["recipient", "attachment_name"])
def test_inaccessible_filter_metadata_reports_omission_instead_of_absence(field):
    recipient = SimpleNamespace(Name="Available name", AddressEntry=None)
    backend = Outlook(
        Application(
            [Mail(Recipients=Collection([recipient]), Attachments=Collection([ComFailure()]))]
        )
    )
    page = backend.emails(SearchArguments(**{field: "unknown"}))
    assert not page.items and page.omitted == 1
    assert page.coverage.exhausted and not page.coverage.evaluation_complete
    assert backend.emails(SearchArguments(recipient="available NAME")).items


def test_unused_collection_properties_are_not_read_and_missing_categories_are_inconclusive():
    class CountOnly:
        Count = 1

        def Item(self, index):
            raise AssertionError("Attachment filenames must only be read when requested")

    mail = replace(Mail(), Attachments=CountOnly(), Recipients=None)
    backend = Outlook(Application([mail]))
    assert backend.emails(SearchArguments(query="report")).items

    class MissingCategories:
        @property
        def Categories(self):
            raise ComFailure()

        def __getattr__(self, name):
            return getattr(mail, name)

    page = Outlook(Application([MissingCategories()])).emails(SearchArguments(category="unknown"))
    assert page.omitted == 1 and not page.coverage.evaluation_complete

from contextlib import nullcontext
from dataclasses import replace
from types import SimpleNamespace

import pytest

from outlook_local_mcp.enums import EErrorCode
from outlook_local_mcp.errors import OutlookError
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


@pytest.mark.parametrize(
    "field,property_name",
    [
        ("recipient", "Recipients"),
        ("attachment_name", "Attachments"),
        ("category", "Categories"),
        ("importance", "Importance"),
    ],
)
def test_general_required_metadata_denial_stops_search(field, property_name):
    class DeniedProperty:
        def __init__(self, index):
            self.mail = Mail(EntryID=f"synthetic-denied-{index}")

        def __getattr__(self, name):
            if name == property_name:
                raise ComFailure()
            return getattr(self.mail, name)

    arguments = SearchArguments(**{field: "high" if field == "importance" else "unknown"})
    backend = Outlook(Application([DeniedProperty(index) for index in range(5)]))
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.ACCESS_DENIED):
        backend.emails(arguments)


@pytest.mark.parametrize("field", ["recipient", "attachment_name"])
def test_partial_metadata_denial_keeps_proven_matches_but_stops_unmatched_search(field):
    recipient = SimpleNamespace(
        Name="Available", AddressEntry=SimpleNamespace(Type="SMTP", Address="person@example.com")
    )
    mails = [
        Mail(
            EntryID=f"synthetic-partial-{index}",
            Recipients=Collection([ComFailure(), recipient]),
            Attachments=Collection([ComFailure(), SimpleNamespace(FileName="Available.pdf")]),
        )
        for index in range(5)
    ]
    backend = Outlook(Application(mails))
    assert len(backend.emails(SearchArguments(**{field: "available"})).items) == 5
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.ACCESS_DENIED):
        backend.emails(SearchArguments(**{field: "unknown"}))


def test_denied_recipient_smtp_resolution_stops_unmatched_search():
    class DeniedAddress:
        Type = "SMTP"

        @property
        def Address(self):
            raise ComFailure()

    recipient = SimpleNamespace(Name="Available", AddressEntry=DeniedAddress())
    backend = Outlook(
        Application(
            [
                Mail(EntryID=f"synthetic-address-{index}", Recipients=Collection([recipient]))
                for index in range(5)
            ]
        )
    )
    assert len(backend.emails(SearchArguments(recipient="available")).items) == 5
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.ACCESS_DENIED):
        backend.emails(SearchArguments(recipient="unknown"))

import json
from types import SimpleNamespace

import pytest
from pydantic import ValidationError

from outlook_local_mcp.accounts import select_account
from outlook_local_mcp.action_models import (
    DraftArguments,
    PreviewArguments,
    ReplyArguments,
    SendArguments,
)
from outlook_local_mcp.drafts import create_draft, open_email, reply_to_email
from outlook_local_mcp.enums import EErrorCode
from outlook_local_mcp.errors import OutlookError
from outlook_local_mcp.models import RuntimeOptions
from outlook_local_mcp.outlook import Outlook
from outlook_local_mcp.sending import PreviewStore, prepare_send, send_draft
from outlook_local_mcp.worker import handle_request

from .action_fakes import Application, Draft
from .fakes import ComFailure


@pytest.fixture
def backend():
    return Outlook(Application())


def make_draft(backend, **extra):
    result = create_draft(
        backend,
        DraftArguments(to=["to@example.com"], subject="Synthetic preview", body="Body", **extra),
    )
    arguments = PreviewArguments(
        entry_id=result.entry_id, store_id=result.store_id, account_email=result.account.email
    )
    return backend.item(result.entry_id, result.store_id), arguments


def confirmed(arguments, preview):
    return SendArguments(
        entry_id=arguments.entry_id,
        store_id=arguments.store_id,
        confirmation_token=preview.confirmation_token,
    )


def test_draft_selects_account_and_preserves_recipient_roles_without_sending(backend):
    item, identifiers = make_draft(
        backend, account_email="second@example.com", cc=["cc@example.com"], bcc=["bcc@example.com"]
    )
    assert identifiers.store_id == "secondary"
    assert item.SendUsingAccount.SmtpAddress == "second@example.com"
    assert [recipient.Type for recipient in item.Recipients.items] == [1, 2, 3]
    assert item.events == ["save"] and item.Saved and not item.Sent
    assert open_email(backend, identifiers).opened
    assert item.events[-2:] == [("display", False), "activate"]


def test_ambiguous_account_requires_explicit_email_and_matching_store(backend):
    namespace = backend.namespace()
    namespace.Accounts.items.append(
        SimpleNamespace(
            SmtpAddress="alias@example.com",
            DeliveryStore=namespace.DefaultStore,
            DisplayName="Alias",
        )
    )
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.INVALID_ARGUMENT):
        select_account(namespace)
    assert (
        select_account(namespace, account_email="alias@example.com").SmtpAddress
        == "alias@example.com"
    )
    with pytest.raises(OutlookError):
        select_account(namespace, "secondary", "alias@example.com")
    assert not backend.mutation_attempted


@pytest.mark.parametrize(
    "address", ["Display <person@example.com>", "no-domain", "a@example.com\r\nBcc: b@example.com"]
)
def test_invalid_smtp_is_rejected_before_com(address):
    with pytest.raises(ValidationError):
        DraftArguments(to=[address], subject="Subject", body="Body")


def test_reply_preserves_full_quotation_original_and_native_account(backend):
    namespace = backend.namespace()
    original = Draft(namespace.DefaultStore.inbox, namespace.Accounts.Item(2))
    original.EntryID = "original"
    original.Body = "quotation" * 5000
    namespace.DefaultStore.inbox.Items.items.append(original)
    before = original.Body, original.UnRead, original.Subject
    result = reply_to_email(
        backend,
        ReplyArguments(entry_id="original", store_id="default", body="Reply", reply_all=True),
    )
    draft = backend.item(result.entry_id, result.store_id)
    assert draft.Body == "Reply\r\n\r\n" + before[0]
    assert result.body_truncated and result.body_length == len(draft.Body)
    assert result.store_id == "secondary" and len(result.recipients) == 2
    assert (original.Body, original.UnRead, original.Subject) == before
    assert original.events == [] and draft.events == ["save"]
    with pytest.raises(
        OutlookError, check=lambda error: error.code == EErrorCode.UNSUPPORTED_COMPOSITION
    ):
        prepare_send(
            backend,
            PreviewArguments(
                entry_id=result.entry_id,
                store_id=result.store_id,
                account_email=result.account.email,
            ),
        )


def test_preview_consumption_and_simulated_submission(backend):
    item, arguments = make_draft(backend)
    preview = prepare_send(backend, arguments)
    assert preview.body == item.Body and preview.sender.email == "first@example.com"
    assert preview.recipients[0].kind == "to" and item.events == ["save"]
    assert send_draft(backend, confirmed(arguments, preview)).status == "submitted_to_outlook"
    with pytest.raises(
        OutlookError, check=lambda error: error.code == EErrorCode.CONFIRMATION_INVALID
    ):
        send_draft(backend, confirmed(arguments, preview))
    assert item.events.count("send") == 1


@pytest.mark.parametrize(
    "change",
    [
        "subject",
        "body",
        "role",
        "address",
        "account",
        "from",
        "attachments",
        "format",
        "saved",
        "location",
    ],
)
def test_changed_or_unsupported_draft_is_never_sent(backend, change):
    item, arguments = make_draft(backend)
    preview = prepare_send(backend, arguments)
    if change == "subject":
        item.Subject += " changed"
    elif change == "body":
        item.Body += " changed"
    elif change == "role":
        item.Recipients.Item(1).Type = 3
    elif change == "address":
        item.Recipients.Item(1).AddressEntry.Address = "changed@example.com"
    elif change == "account":
        item.SendUsingAccount = backend.namespace().Accounts.Item(2)
    elif change == "from":
        item.SentOnBehalfOfName = "other@example.com"
    elif change == "attachments":
        item.Attachments.items.append(SimpleNamespace(FileName="added.txt", Size=1))
    elif change == "format":
        item.BodyFormat = 2
    elif change == "saved":
        item.Saved = False
    elif change == "location":
        item.Parent = backend.namespace().DefaultStore.inbox
    with pytest.raises(OutlookError):
        send_draft(backend, confirmed(arguments, preview))
    assert "send" not in item.events and not backend.previews.previews


def test_preview_expiry_replacement_ids_and_own_from_alias(backend):
    clock = [0.0]
    backend.previews = PreviewStore(lambda: clock[0])
    item, arguments = make_draft(backend)
    item.SentOnBehalfOfName = "First"
    first = prepare_send(backend, arguments)
    second = prepare_send(backend, arguments)
    with pytest.raises(OutlookError):
        send_draft(backend, confirmed(arguments, first))
    wrong = confirmed(arguments, second).model_copy(update={"entry_id": "another"})
    with pytest.raises(OutlookError):
        send_draft(backend, wrong)
    assert second.confirmation_token in backend.previews.previews
    clock[0] = 300
    with pytest.raises(OutlookError):
        send_draft(backend, confirmed(arguments, second))
    assert "send" not in item.events


def test_worker_defaults_deny_writes_and_post_mutation_failure_is_unknown(backend):
    request = json.dumps(
        {
            "operation": "create_draft",
            "arguments": {"to": ["to@example.com"], "subject": "Subject", "body": "Body"},
        }
    ).encode()
    denied = json.loads(handle_request(backend, request, RuntimeOptions()))
    assert denied["error"]["code"] == "CAPABILITY_DISABLED"
    assert not backend.namespace().DefaultStore.drafts.Items.items
    items = backend.namespace().DefaultStore.drafts.Items
    original_add = items.Add

    def broken_result(kind):
        draft = original_add(kind)

        def save():
            draft.Saved = True
            draft.EntryID = None

        draft.Save = save
        return draft

    items.Add = broken_result
    failed = json.loads(handle_request(backend, request, RuntimeOptions(enable_write_tools=True)))
    assert failed["error"]["code"] == "WRITE_OUTCOME_UNKNOWN"
    assert not failed["error"]["retryable"]


def test_explicit_preview_account_survives_null_native_reference_and_own_assignment(backend):
    item, arguments = make_draft(backend)
    item.SendUsingAccount = None
    preview = prepare_send(backend, arguments)
    original_put = item._oleobj_.Invoke

    def dirty_assignment(*args):
        original_put(*args)
        item.Saved = False

    item._oleobj_.Invoke = dirty_assignment
    assert send_draft(backend, confirmed(arguments, preview)).status == "submitted_to_outlook"
    assert item.events == ["save", "send"]


@pytest.mark.parametrize("mode", ["failure", "content_change", "account_change"])
def test_assignment_failure_or_side_effect_never_sends_and_consumes_token(backend, mode):
    item, arguments = make_draft(backend)
    preview = prepare_send(backend, arguments)
    original_put = item._oleobj_.Invoke

    def changed_assignment(*args):
        if mode == "failure":
            raise RuntimeError("synthetic-private-assignment")
        original_put(*args)
        if mode == "content_change":
            item.Body = "changed"
        else:
            item.SendUsingAccount = backend.namespace().Accounts.Item(2)

    item._oleobj_.Invoke = changed_assignment
    backend.mutation_attempted = False
    with pytest.raises((OutlookError, RuntimeError)):
        send_draft(backend, confirmed(arguments, preview))
    assert backend.mutation_attempted and not backend.previews.previews
    assert "send" not in item.events


def test_saved_external_edit_invalidates_preview_even_when_visible_content_matches(backend):
    from datetime import timedelta

    item, arguments = make_draft(backend)
    preview = prepare_send(backend, arguments)
    item.LastModificationTime += timedelta(seconds=1)
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.DRAFT_CHANGED):
        send_draft(backend, confirmed(arguments, preview))
    assert "send" not in item.events


def test_send_failure_is_unknown_consumes_preview_and_does_not_repeat(backend):
    item, arguments = make_draft(backend)
    preview = prepare_send(backend, arguments)

    def failed_send():
        item.events.append("send-attempt")
        raise RuntimeError("synthetic-private-send-error")

    item.Send = failed_send
    request = json.dumps(
        {"operation": "send_draft", "arguments": confirmed(arguments, preview).model_dump()}
    ).encode()
    options = RuntimeOptions(enable_write_tools=True, enable_send=True)
    response = json.loads(handle_request(backend, request, options))
    assert response["error"]["code"] == "WRITE_OUTCOME_UNKNOWN"
    assert not response["error"]["retryable"] and not backend.previews.previews
    assert item.events.count("send-attempt") == 1


def test_account_lost_after_assignment_readback_never_sends(backend, monkeypatch):
    item, arguments = make_draft(backend)
    preview = prepare_send(backend, arguments)

    def transient_account(draft):
        account = draft.__dict__["SendUsingAccount"]
        draft.__dict__["SendUsingAccount"] = None
        return account

    monkeypatch.setattr(
        Draft,
        "SendUsingAccount",
        property(
            transient_account,
            lambda draft, account: draft.__dict__.__setitem__("SendUsingAccount", account),
        ),
        raising=False,
    )
    with pytest.raises(OutlookError):
        send_draft(backend, confirmed(arguments, preview))
    assert "send" not in item.events and not backend.previews.previews


@pytest.mark.parametrize(
    "property_value,allowed",
    [
        (ComFailure(), False),
        (ComFailure(0x8004010F), True),
        (ComFailure(0x80040107), False),
        (ComFailure(0x80070002), False),
        (ComFailure(0x80004005), False),
        ("unresolved-internal-name", False),
        ("bad@@example.com", False),
        (123, False),
        (None, False),
    ],
)
def test_represented_from_only_falls_back_when_property_is_absent(backend, property_value, allowed):
    item, arguments = make_draft(backend)
    item.SentOnBehalfOfName = "First"

    def read_property(name):
        if isinstance(property_value, Exception):
            raise property_value
        return property_value

    item.PropertyAccessor.GetProperty = read_property
    if allowed:
        assert prepare_send(backend, arguments).sender.email == "first@example.com"
    else:
        with pytest.raises(OutlookError):
            prepare_send(backend, arguments)
        assert not backend.previews.previews
    assert item.events == ["save"]


@pytest.mark.parametrize("represented_name", ["", "First"])
@pytest.mark.parametrize(
    "status,details,allowed",
    [
        (-2147221233, None, True),
        (-2147352567, (0, None, None, None, 0, -2147221233), True),
        (0x8004010F, (0, None, None, None, 0, 0x80070005), False),
        (0x80070005, (0, None, None, None, 0, 0x8004010F), False),
        (0x80040107, (0, None, None, None, 0, 0x8004010F), False),
        (0x80020009, (0, None, None, None, 0, 0x80070002), False),
        (0x80020009, (0, None), False),
        (0x8004010F, "malformed", False),
        (0x8004010F, (0, None, None, None, 0, None), False),
        (None, (0, None, None, None, 0, 0x8004010F), False),
    ],
)
def test_property_absence_requires_unambiguous_numeric_status(
    backend, represented_name, status, details, allowed
):
    item, arguments = make_draft(backend)
    item.SentOnBehalfOfName = represented_name
    error = ComFailure(status)
    error.excepinfo = details

    def read_property(name):
        raise error

    item.PropertyAccessor.GetProperty = read_property
    if allowed:
        assert prepare_send(backend, arguments).sender.email == "first@example.com"
    else:
        with pytest.raises(OutlookError):
            prepare_send(backend, arguments)
        assert not backend.previews.previews
    assert item.events == ["save"]

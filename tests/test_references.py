import json

import pytest

from outlook_local_mcp.enums import EErrorCode, EReferenceKind
from outlook_local_mcp.errors import OutlookError
from outlook_local_mcp.models import RuntimeOptions
from outlook_local_mcp.outlook import Outlook
from outlook_local_mcp.references import ReferenceStore
from outlook_local_mcp.worker import handle_request

from .fakes import Application, Mail


def test_short_references_bind_kind_and_store_with_native_input_compatibility():
    refs = ReferenceStore()
    first = refs.output(
        {
            "entry_id": "same",
            "store_id": "one",
            "folder_id": "inbox",
            "body": "olm_not_an_identifier",
        }
    )
    second = refs.output({"entry_id": "same", "store_id": "two"})
    assert first["entry_id"] != second["entry_id"]
    assert first["body"] == "olm_not_an_identifier"
    resolved = refs.arguments({"entry_id": first["entry_id"], "store_id": "one"})
    assert resolved == {"entry_id": "same", "store_id": "one"}
    with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.INVALID_ARGUMENT):
        refs.arguments({"entry_id": first["entry_id"], "store_id": second["store_id"]})
    with pytest.raises(OutlookError):
        refs.arguments({"entry_id": first["folder_id"], "store_id": first["store_id"]})


def test_reuse_refreshes_refs_without_rebinding_and_reset_expires_them():
    now = [0.0]
    refs = ReferenceStore(lambda: now[0])
    token = refs.put(EReferenceKind.STORE, "store", None)
    now[0] = 590
    assert refs.resolve(token, EReferenceKind.STORE, None) == "store"
    now[0] = 601
    refs.prune()
    assert refs.put(EReferenceKind.STORE, "store", None) == token
    now[0] = 1202
    refs.prune()
    with pytest.raises(
        OutlookError, check=lambda error: error.code == EErrorCode.REFERENCE_EXPIRED
    ):
        refs.resolve(token, EReferenceKind.STORE, None)
    fresh = refs.put(EReferenceKind.STORE, "store", None)
    refs.clear()
    with pytest.raises(OutlookError):
        refs.resolve(fresh, EReferenceKind.STORE, None)


def test_worker_native_cursor_signature_survives_short_and_native_ids():
    backend = Outlook(Application([Mail(), Mail(EntryID="second")]))

    def call(operation, arguments):
        return json.loads(
            handle_request(
                backend,
                json.dumps({"operation": operation, "arguments": arguments}).encode(),
                RuntimeOptions(),
            )
        )

    first = call("recent_emails", {"store_id": "store", "limit": 1})["result"]
    second = call(
        "recent_emails", {"store_id": first["store_id"], "limit": 1, "cursor": first["next_cursor"]}
    )["result"]
    assert first["items"][0]["entry_id"] != second["items"][0]["entry_id"]
    item = first["items"][0]
    assert "result" in call(
        "read_email", {"entry_id": item["entry_id"], "store_id": item["store_id"]}
    )
    backend.close()
    error = call("read_email", {"entry_id": item["entry_id"], "store_id": item["store_id"]})[
        "error"
    ]
    assert error["code"] == "REFERENCE_EXPIRED"


def test_capacity_failure_clears_unpublished_cursor_and_preview(monkeypatch):
    backend = Outlook(Application([Mail(), Mail(EntryID="second")]))
    monkeypatch.setattr("outlook_local_mcp.references.MAX_SESSION_IDS", 1)
    request = json.dumps({"operation": "recent_emails", "arguments": {"limit": 1}}).encode()
    response = json.loads(handle_request(backend, request, RuntimeOptions()))
    assert response["error"]["code"] == "REFERENCE_LIMIT"
    assert (
        not backend.cursors.sessions
        and not backend.references.entries
        and not backend.previews.previews
    )


def test_synthetic_id_heavy_json_size_reduction():
    store = "AB" * 350
    payload = {
        "store_id": store,
        "folder_id": "CD" * 45,
        "items": [
            {
                "entry_id": f"{index:04X}" + "EF" * 70,
                "store_id": store,
                "folder_id": "CD" * 45,
                "subject": "Synthetic email",
                "received_at": "2026-01-01T00:00:00Z",
            }
            for index in range(20)
        ],
    }
    before = len(json.dumps(payload).encode())
    after = len(json.dumps(ReferenceStore().output(payload)).encode())
    assert after < before * 0.25


def test_post_save_reference_failure_is_unknown_and_clears_backend(monkeypatch):
    from .action_fakes import Application as ActionApplication

    backend = Outlook(ActionApplication())
    folder = backend.namespace().DefaultStore.drafts
    monkeypatch.setattr("outlook_local_mcp.references.MAX_SESSION_IDS", 0)
    request = json.dumps(
        {
            "operation": "create_draft",
            "arguments": {"to": ["to@example.com"], "subject": "Synthetic", "body": "Body"},
        }
    ).encode()
    response = json.loads(handle_request(backend, request, RuntimeOptions(enable_write_tools=True)))
    assert folder.Items.items[0].events == ["save"]
    assert response["error"]["code"] == "WRITE_OUTCOME_UNKNOWN"
    assert not response["error"]["retryable"]
    assert not backend.references.entries and not backend.previews.previews

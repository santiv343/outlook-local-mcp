import pytest

from outlook_local_mcp import pagination
from outlook_local_mcp.enums import EErrorCode, EToolName
from outlook_local_mcp.errors import OutlookError
from outlook_local_mcp.pagination import CursorStore, SearchSession, scan_page, signature

from .fakes import Collection, ComFailure, Mail


def session(items, created=0):
    return SearchSession(Collection(items), "store", "inbox", "filters", created)


def test_no_matches_at_budget_is_not_exhausted_and_continues():
    cursors = CursorStore(clock=lambda: 0)
    traversal = session([Mail(EntryID=str(index)) for index in range(3)])
    first = scan_page(cursors, traversal, 20, lambda _: None, scan_limit=2)
    assert first.items == []
    assert first.coverage.scanned == 2
    assert first.coverage.stop_reason == "scan_limit"
    assert not first.coverage.exhausted
    resumed = cursors.take(first.next_cursor, "filters")
    last = scan_page(cursors, resumed, 20, lambda item: item.EntryID)
    assert last.items == ["2"]
    assert last.coverage.evaluation_complete
    assert traversal.collection.starts == 1


def test_signature_mismatch_preserves_cursor_and_success_consumes_it():
    cursors = CursorStore(clock=lambda: 0)
    token = cursors.put(session([]))
    with pytest.raises(OutlookError, check=lambda error: error.code == "INVALID_ARGUMENT"):
        cursors.take(token, "other")
    cursors.take(token, "filters")
    with pytest.raises(OutlookError, check=lambda error: error.code == "CURSOR_EXPIRED"):
        cursors.take(token, "filters")


def test_expiration_clear_and_capacity(monkeypatch):
    now = [0]
    cursors = CursorStore(clock=lambda: now[0])
    token = cursors.put(session([]))
    now[0] = pagination.CURSOR_TTL_SECONDS
    with pytest.raises(OutlookError, check=lambda error: error.code == "CURSOR_EXPIRED"):
        cursors.take(token, "filters")
    monkeypatch.setattr(pagination, "MAX_SESSIONS", 1)
    token = cursors.put(session([], now[0]))
    with pytest.raises(OutlookError, check=lambda error: error.code == "SEARCH_SESSION_LIMIT"):
        cursors.check_capacity()
    cursors.clear()
    with pytest.raises(OutlookError, check=lambda error: error.code == "CURSOR_EXPIRED"):
        cursors.take(token, "filters")


def test_duplicates_deletions_and_omissions_persist_across_pages():
    cursors = CursorStore(clock=lambda: 0)
    traversal = session([ComFailure(0x8004010F), Mail(), Mail(), Mail(EntryID="second")])
    first = scan_page(cursors, traversal, 1, lambda item: item.EntryID)
    assert first.omitted == 1
    resumed = cursors.take(first.next_cursor, "filters")
    last = scan_page(cursors, resumed, 20, lambda item: item.EntryID)
    assert last.items == ["second"]
    assert last.coverage.exhausted
    assert not last.coverage.evaluation_complete
    assert last.warnings


def test_widespread_denial_stops_operation():
    cursors = CursorStore(clock=lambda: 0)
    traversal = session([ComFailure() for _ in range(5)])
    with pytest.raises(OutlookError, check=lambda error: error.code == "ACCESS_DENIED"):
        scan_page(cursors, traversal, 20, lambda item: item.EntryID)
    assert cursors.sessions == {}


def test_time_budget_and_seen_memory_limits(monkeypatch):
    ticks = iter([0, 11])
    cursors = CursorStore(clock=lambda: next(ticks))
    traversal = session([Mail()])
    page = scan_page(cursors, traversal, 20, lambda item: item.EntryID)
    assert page.coverage.stop_reason == "time_budget"
    assert page.coverage.scanned == 0
    monkeypatch.setattr(pagination, "MAX_SESSION_ID_BYTES", 1)
    with pytest.raises(OutlookError) as error:
        traversal.remember("too-long")
    assert error.value.code == EErrorCode.SEARCH_SESSION_LIMIT


def test_cursor_signature_binds_tool_filters_but_not_page_size():
    args = {"query": "synthetic", "limit": 1, "cursor": None}
    original = signature(EToolName.SEARCH_EMAILS, args)
    assert signature(EToolName.SEARCH_EMAILS, {**args, "limit": 5, "cursor": "token"}) == original
    assert signature(EToolName.RECENT_EMAILS, args) != original
    assert signature(EToolName.SEARCH_EMAILS, {**args, "query": "different"}) != original

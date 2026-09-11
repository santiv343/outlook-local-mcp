"""Bounded, one-use cursors whose COM collections stay inside the worker."""

import json
import secrets
import time
from collections.abc import Callable
from typing import Generic, TypeVar

from pydantic import JsonValue

from .com_types import ICursorCollection, IFolder, IIdentifiedItem, IOutlookItem
from .config import (
    CURSOR_RANDOM_BYTES,
    CURSOR_TTL_SECONDS,
    MAX_CONSECUTIVE_DENIALS,
    MAX_SESSION_ID_BYTES,
    MAX_SESSION_IDS,
    MAX_SESSIONS,
    SCAN_LIMIT,
    SCAN_SECONDS,
    SEEN_ID_OVERHEAD_BYTES,
)
from .enums import EErrorCode, EStopReason, EToolName
from .errors import OutlookError, com_error
from .models import Coverage, Page, WarningInfo

TItem = TypeVar("TItem")
TComItem = TypeVar("TComItem", bound=IIdentifiedItem, covariant=True)


def signature(operation: EToolName, arguments: dict[str, JsonValue]) -> str:
    return json.dumps(
        [
            operation,
            {key: value for key, value in arguments.items() if key not in {"limit", "cursor"}},
        ],
        sort_keys=True,
    )


class SearchSession(Generic[TComItem]):
    def __init__(
        self,
        collection: ICursorCollection[TComItem],
        store_id: str,
        folder_id: str,
        query_signature: str,
        created: float,
    ) -> None:
        self.collection = collection
        self.store_id = store_id
        self.folder_id = folder_id
        self.signature = query_signature
        self.created = created
        self.started = False
        self.seen: set[str] = set()
        self.seen_bytes = 0
        self.omitted = 0

    def next_item(self) -> TComItem | None:
        if not self.started:
            self.started = True
            return self.collection.GetFirst()
        return self.collection.GetNext()

    def remember(self, entry_id: str) -> bool:
        if entry_id in self.seen:
            return False
        size = len(entry_id.encode("utf-8")) + SEEN_ID_OVERHEAD_BYTES
        if len(self.seen) >= MAX_SESSION_IDS or self.seen_bytes + size > MAX_SESSION_ID_BYTES:
            raise OutlookError(EErrorCode.SEARCH_SESSION_LIMIT)
        self.seen.add(entry_id)
        self.seen_bytes += size
        return True


class MailSession(SearchSession[IOutlookItem]):
    """A mailbox traversal; its Python type avoids probing live COM objects for types."""


class FolderSession(SearchSession[IFolder]):
    """A folder traversal with the same bounded lifecycle."""


class CursorStore:
    def __init__(self, clock: Callable[[], float] = time.monotonic) -> None:
        self.clock = clock
        self.sessions: dict[str, SearchSession[IIdentifiedItem]] = {}

    def prune(self) -> None:
        now = self.clock()
        self.sessions = {
            token: session
            for token, session in self.sessions.items()
            if now - session.created < CURSOR_TTL_SECONDS
        }

    def check_capacity(self) -> None:
        self.prune()
        if len(self.sessions) >= MAX_SESSIONS:
            raise OutlookError(EErrorCode.SEARCH_SESSION_LIMIT)

    def take(self, token: str, query_signature: str) -> SearchSession[IIdentifiedItem]:
        self.prune()
        session = self.sessions.get(token)
        if session is None:
            raise OutlookError(EErrorCode.CURSOR_EXPIRED)
        if session.signature != query_signature:
            raise OutlookError(
                EErrorCode.INVALID_ARGUMENT, "A cursor must retain its original tool and filters."
            )
        del self.sessions[token]
        return session

    def put(self, session: SearchSession[IIdentifiedItem]) -> str:
        token = secrets.token_urlsafe(CURSOR_RANDOM_BYTES)
        self.sessions[token] = session
        return token

    def clear(self) -> None:
        self.sessions.clear()


def scan_page(
    cursors: CursorStore,
    session: SearchSession[TComItem],
    limit: int,
    project: Callable[[TComItem], TItem | None],
    *,
    scan_limit: int = SCAN_LIMIT,
    time_budget: float = SCAN_SECONDS,
) -> Page[TItem]:
    started = cursors.clock()
    items: list[TItem] = []
    scanned = omitted = denied = 0
    reason = EStopReason.EXHAUSTED
    while True:
        if len(items) >= limit:
            reason = EStopReason.PAGE_LIMIT
            break
        if scanned >= scan_limit:
            reason = EStopReason.SCAN_LIMIT
            break
        if cursors.clock() - started >= time_budget:
            reason = EStopReason.TIME_BUDGET
            break
        scanned += 1
        try:
            item = session.next_item()
            if item is None:
                scanned -= 1
                break
            if not session.remember(item.EntryID):
                continue
            projected = project(item)
            if projected is not None:
                items.append(projected)
            denied = 0
        except Exception as error:
            mapped = (
                error
                if isinstance(error, OutlookError)
                else com_error(error, EErrorCode.ITEM_NOT_FOUND)
            )
            if mapped.code not in {
                EErrorCode.ITEM_NOT_FOUND,
                EErrorCode.ACCESS_DENIED,
                EErrorCode.BODY_UNAVAILABLE,
            }:
                raise mapped from None
            omitted += 1
            session.omitted += 1
            denied = denied + 1 if mapped.code == EErrorCode.ACCESS_DENIED else 0
            if denied >= MAX_CONSECUTIVE_DENIALS:
                raise OutlookError(EErrorCode.ACCESS_DENIED) from None
    exhausted = reason == EStopReason.EXHAUSTED
    warnings = []
    if session.omitted:
        warnings.append(
            WarningInfo(
                code="ITEMS_OMITTED",
                message="Some candidates were inaccessible during this search. "
                "An exhausted traversal may still have incomplete evaluation.",
            )
        )
    return Page[TItem](
        items=items,
        store_id=session.store_id,
        folder_id=session.folder_id,
        next_cursor=None if exhausted else cursors.put(session),
        omitted=omitted,
        warnings=warnings,
        coverage=Coverage(
            scanned=scanned,
            exhausted=exhausted,
            stop_reason=reason,
            evaluation_complete=exhausted and session.omitted == 0,
        ),
    )

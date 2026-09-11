"""Bounded, store-bound short identifiers at the private JSON boundary."""

import secrets
import time
from collections.abc import Callable

from pydantic import JsonValue

from .config import (
    CURSOR_TTL_SECONDS,
    MAX_SESSION_ID_BYTES,
    MAX_SESSION_IDS,
    SEEN_ID_OVERHEAD_BYTES,
)
from .enums import EErrorCode, EReferenceKind
from .errors import OutlookError
from .reference_constants import REFERENCE_FIELDS, REFERENCE_PREFIX, REFERENCE_RANDOM_BYTES
from .reference_models import Reference, TReferenceKey


class ReferenceStore:
    def __init__(self, clock: Callable[[], float] = time.monotonic) -> None:
        self.clock = clock
        self.entries: dict[str, Reference] = {}
        self.reverse: dict[TReferenceKey, str] = {}
        self.size = 0

    def clear(self) -> None:
        self.entries.clear()
        self.reverse.clear()
        self.size = 0

    def prune(self) -> None:
        expired = [
            token
            for token, entry in self.entries.items()
            if self.clock() - entry.touched >= CURSOR_TTL_SECONDS
        ]
        for token in expired:
            entry = self.entries.pop(token)
            del self.reverse[entry.key]
            self.size -= entry.size

    def put(self, kind: EReferenceKind, native_id: str, store_id: str | None) -> str:
        if kind != EReferenceKind.STORE and store_id is None:
            raise OutlookError(EErrorCode.INTERNAL_ERROR)
        key = (kind, store_id, native_id)
        existing = self.reverse.get(key)
        if existing is not None:
            self.entries[existing].touched = self.clock()
            return existing
        size = len(native_id.encode()) + len((store_id or "").encode()) + SEEN_ID_OVERHEAD_BYTES
        if len(self.entries) >= MAX_SESSION_IDS or self.size + size > MAX_SESSION_ID_BYTES:
            raise OutlookError(EErrorCode.REFERENCE_LIMIT)
        token = REFERENCE_PREFIX + kind + "_" + secrets.token_urlsafe(REFERENCE_RANDOM_BYTES)
        while token in self.entries:
            token = REFERENCE_PREFIX + kind + "_" + secrets.token_urlsafe(REFERENCE_RANDOM_BYTES)
        self.entries[token] = Reference(key, self.clock(), size)
        self.reverse[key] = token
        self.size += size
        return token

    def resolve(self, value: str, kind: EReferenceKind, store_id: str | None) -> str:
        if not value.startswith(REFERENCE_PREFIX):
            return value
        entry = self.entries.get(value)
        if entry is None or self.clock() - entry.touched >= CURSOR_TTL_SECONDS:
            raise OutlookError(EErrorCode.REFERENCE_EXPIRED)
        if entry.key[0] != kind or (kind != EReferenceKind.STORE and entry.key[1] != store_id):
            raise OutlookError(
                EErrorCode.INVALID_ARGUMENT, "Reference kind or mailbox does not match."
            )
        entry.touched = self.clock()
        return entry.key[2]

    def arguments(self, arguments: dict[str, JsonValue]) -> dict[str, JsonValue]:
        result = dict(arguments)
        store = result.get("store_id")
        if isinstance(store, str):
            store = self.resolve(store, EReferenceKind.STORE, None)
            result["store_id"] = store
        for field, kind in REFERENCE_FIELDS.items():
            value = result.get(field)
            if field != "store_id" and isinstance(value, str):
                result[field] = self.resolve(value, kind, store if isinstance(store, str) else None)
        return result

    def output(
        self, payload: dict[str, JsonValue], store_id: str | None = None
    ) -> dict[str, JsonValue]:
        local_store = payload.get("store_id", store_id)
        if isinstance(local_store, str):
            store_id = local_store
        return {key: self.value(key, value, store_id) for key, value in payload.items()}

    def value(self, key: str, value: JsonValue, store_id: str | None) -> JsonValue:
        if key in REFERENCE_FIELDS and isinstance(value, str):
            kind = REFERENCE_FIELDS[key]
            return self.put(kind, value, None if kind == EReferenceKind.STORE else store_id)
        if isinstance(value, dict):
            return self.output(value, store_id)
        if isinstance(value, list):
            return [self.value("", item, store_id) for item in value]
        return value

"""Worker-only identities. Never contain message content or COM references."""

from dataclasses import dataclass

from .enums import EReferenceKind

TReferenceKey = tuple[EReferenceKind, str | None, str]


@dataclass
class Reference:
    key: TReferenceKey
    touched: float
    size: int

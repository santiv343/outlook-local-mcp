"""Public tool contracts. No COM objects belong in these models."""

from typing import Annotated, Generic, Literal, TypeVar

from pydantic import BaseModel, ConfigDict, Field, JsonValue, model_validator

from .config import (
    DEFAULT_BODY_CHARACTERS,
    DEFAULT_PAGE_SIZE,
    MAX_BODY_CHARACTERS,
    MAX_CURSOR_CHARACTERS,
    MAX_DATE_CHARACTERS,
    MAX_IDENTIFIER_CHARACTERS,
    MAX_PAGE_SIZE,
    MAX_QUERY_CHARACTERS,
)
from .enums import EErrorCode, ERecipientKind, EStopReason, EToolName

TIdentifier = Annotated[str, Field(min_length=1, max_length=MAX_IDENTIFIER_CHARACTERS)]
TCursor = Annotated[str, Field(min_length=1, max_length=MAX_CURSOR_CHARACTERS)]
TLimit = Annotated[int, Field(ge=1, le=MAX_PAGE_SIZE)]
TItem = TypeVar("TItem")


class Arguments(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)


class FolderArguments(Arguments):
    store_id: TIdentifier
    parent_folder_id: TIdentifier | None = None
    limit: TLimit = DEFAULT_PAGE_SIZE
    cursor: TCursor | None = None


class RecentArguments(Arguments):
    store_id: TIdentifier | None = None
    folder_id: TIdentifier | None = None
    limit: TLimit = DEFAULT_PAGE_SIZE
    cursor: TCursor | None = None

    @model_validator(mode="after")
    def require_store(self) -> "RecentArguments":
        if self.folder_id is not None and self.store_id is None:
            raise ValueError("folder_id requires store_id")
        return self


class SearchArguments(RecentArguments):
    query: Annotated[str, Field(max_length=MAX_QUERY_CHARACTERS)] = ""
    query_in: Literal["subject", "body", "subject_body"] = "subject"
    sender: Annotated[str, Field(max_length=MAX_QUERY_CHARACTERS)] | None = None
    after: Annotated[str, Field(max_length=MAX_DATE_CHARACTERS)] | None = None
    before: Annotated[str, Field(max_length=MAX_DATE_CHARACTERS)] | None = None
    unread: bool | None = None
    has_attachments: bool | None = None


class ReadArguments(Arguments):
    entry_id: TIdentifier
    store_id: TIdentifier
    body_offset: Annotated[int, Field(ge=0)] = 0
    body_limit: Annotated[int, Field(ge=1, le=MAX_BODY_CHARACTERS)] = DEFAULT_BODY_CHARACTERS


class ErrorInfo(BaseModel):
    code: EErrorCode
    message: str
    retryable: bool


class WarningInfo(BaseModel):
    code: str
    message: str


class OutlookStatus(BaseModel):
    available: bool
    version: str | None = None
    store_count: int | None = None
    error: ErrorInfo | None = None


class Mailbox(BaseModel):
    store_id: str
    name: str
    is_default: bool


class Mailboxes(BaseModel):
    items: list[Mailbox]
    warnings: list[WarningInfo] = Field(default_factory=list)
    omitted: int = 0


class Folder(BaseModel):
    folder_id: str
    store_id: str
    name: str
    has_children: bool


class Person(BaseModel):
    name: str | None
    email: str | None


class EmailSummary(BaseModel):
    entry_id: str
    store_id: str
    folder_id: str
    subject: str
    sender: Person
    received_at: str
    unread: bool
    has_attachments: bool
    warnings: list[WarningInfo] = Field(default_factory=list)


class Coverage(BaseModel):
    scanned: int
    exhausted: bool
    stop_reason: EStopReason
    consistency: Literal["best_effort"] = "best_effort"
    evaluation_complete: bool


class Page(BaseModel, Generic[TItem]):
    items: list[TItem]
    store_id: str
    folder_id: str
    next_cursor: str | None
    coverage: Coverage
    warnings: list[WarningInfo] = Field(default_factory=list)
    omitted: int = 0


class Recipient(Person):
    kind: ERecipientKind


class Attachment(BaseModel):
    name: str
    size: int


class EmailDetail(EmailSummary):
    recipients: list[Recipient]
    sent_at: str | None
    body: str
    body_offset: int
    next_body_offset: int | None
    body_truncated: bool
    body_length: int
    attachments: list[Attachment]
    omitted_recipients: int = 0
    omitted_attachments: int = 0


class WorkerRequest(Arguments):
    operation: EToolName
    arguments: dict[str, JsonValue]


class WorkerSuccess(Arguments):
    result: dict[str, JsonValue]


class WorkerFailure(Arguments):
    error: ErrorInfo


TWorkerResponse = WorkerSuccess | WorkerFailure


class ConfigurationResult(BaseModel):
    changed: bool
    backup_created: bool

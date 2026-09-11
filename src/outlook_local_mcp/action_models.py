"""Strict input/output contracts for explicit Outlook interactions."""

from typing import Annotated, Literal

from pydantic import AfterValidator, BaseModel, Field, model_validator

from .addresses import explicit_smtp
from .config import (
    CONVERSATION_BODY_CHARACTERS,
    CONVERSATION_PAGE_SIZE,
    MAX_BODY_CHARACTERS,
    MAX_CONVERSATION_PAGE_SIZE,
    MAX_DRAFT_BODY_CHARACTERS,
    MAX_DRAFT_RECIPIENTS,
    MAX_DRAFT_SUBJECT_CHARACTERS,
    MAX_SMTP_CHARACTERS,
)
from .models import Arguments, BodyPage, Person, Recipient, TCursor, TIdentifier, WarningInfo

TSmtp = Annotated[
    str, Field(min_length=1, max_length=MAX_SMTP_CHARACTERS), AfterValidator(explicit_smtp)
]
TRecipientAddresses = Annotated[list[TSmtp], Field(max_length=MAX_DRAFT_RECIPIENTS)]
TDraftBody = Annotated[str, Field(max_length=MAX_DRAFT_BODY_CHARACTERS)]


class ItemArguments(Arguments):
    entry_id: TIdentifier
    store_id: TIdentifier


class DraftArguments(Arguments):
    to: TRecipientAddresses
    subject: Annotated[str, Field(max_length=MAX_DRAFT_SUBJECT_CHARACTERS)]
    body: TDraftBody
    cc: TRecipientAddresses = Field(default_factory=list)
    bcc: TRecipientAddresses = Field(default_factory=list)
    store_id: TIdentifier | None = None
    account_email: TSmtp | None = None

    @model_validator(mode="after")
    def require_recipients(self) -> "DraftArguments":
        if not 1 <= len(self.to) + len(self.cc) + len(self.bcc) <= MAX_DRAFT_RECIPIENTS:
            raise ValueError("Provide at least one recipient within the total recipient limit.")
        return self


class ReplyArguments(ItemArguments):
    body: TDraftBody
    reply_all: bool = False
    account_email: TSmtp | None = None


class ConversationArguments(ItemArguments):
    limit: Annotated[int, Field(ge=1, le=MAX_CONVERSATION_PAGE_SIZE)] = CONVERSATION_PAGE_SIZE
    body_limit: Annotated[int, Field(ge=1, le=MAX_BODY_CHARACTERS)] = CONVERSATION_BODY_CHARACTERS
    cursor: TCursor | None = None


class SendArguments(ItemArguments):
    confirmation_token: TCursor


class PreviewArguments(ItemArguments):
    account_email: TSmtp


class OpenResult(BaseModel):
    entry_id: str
    store_id: str
    opened: Literal[True] = True


class DraftResult(BodyPage):
    entry_id: str
    store_id: str
    folder_id: str
    account: Person
    recipients: list[Recipient]
    subject: str
    warnings: list[WarningInfo] = Field(default_factory=list)


class DraftReview(BaseModel):
    account: Person
    sender: Person
    recipients: list[Recipient]
    subject: str
    body: str
    draft_modified_at: str


class SendPreview(DraftReview):
    entry_id: str
    store_id: str
    confirmation_token: str
    digest: str
    expires_in_seconds: int


class SendResult(BaseModel):
    status: Literal["submitted_to_outlook"] = "submitted_to_outlook"


class SendConfirmation(BaseModel):
    entry_id: str
    store_id: str
    digest: str
    created: float
    account_email: str

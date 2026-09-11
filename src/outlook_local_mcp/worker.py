"""Private JSON pipe worker. All COM work runs on this process's main STA thread."""

import gc
import sys

from pydantic import JsonValue, ValidationError

from .action_models import (
    ConversationArguments,
    DraftArguments,
    ItemArguments,
    PreviewArguments,
    ReplyArguments,
    SendArguments,
)
from .capabilities import require_capability
from .config import MAX_IPC_BYTES
from .conversation import read_conversation
from .drafts import create_draft, open_email, reply_to_email
from .enums import EErrorCode, EToolName
from .errors import OutlookError, com_error
from .models import (
    FolderArguments,
    ReadArguments,
    RecentArguments,
    RuntimeOptions,
    SearchArguments,
    WorkerFailure,
    WorkerRequest,
    WorkerSuccess,
)
from .outlook import Outlook
from .sending import prepare_send, send_draft


def execute(
    backend: Outlook,
    operation: EToolName,
    arguments: dict[str, JsonValue],
    options: RuntimeOptions | None = None,
) -> dict[str, JsonValue]:
    require_capability(operation, options or RuntimeOptions())
    if operation == EToolName.OUTLOOK_STATUS:
        return backend.status().model_dump(mode="json")
    if operation == EToolName.LIST_MAILBOXES:
        return backend.mailboxes().model_dump(mode="json")
    if operation == EToolName.LIST_FOLDERS:
        return backend.folders(FolderArguments.model_validate(arguments)).model_dump(mode="json")
    if operation == EToolName.RECENT_EMAILS:
        return backend.emails(RecentArguments.model_validate(arguments)).model_dump(mode="json")
    if operation == EToolName.SEARCH_EMAILS:
        return backend.emails(SearchArguments.model_validate(arguments)).model_dump(mode="json")
    if operation == EToolName.READ_EMAIL:
        return backend.read(ReadArguments.model_validate(arguments)).model_dump(mode="json")
    if operation == EToolName.READ_CONVERSATION:
        return read_conversation(
            backend, ConversationArguments.model_validate(arguments)
        ).model_dump(mode="json")
    if operation == EToolName.OPEN_EMAIL:
        return open_email(backend, ItemArguments.model_validate(arguments)).model_dump(mode="json")
    if operation == EToolName.CREATE_DRAFT:
        return create_draft(backend, DraftArguments.model_validate(arguments)).model_dump(
            mode="json"
        )
    if operation == EToolName.REPLY_TO_EMAIL:
        return reply_to_email(backend, ReplyArguments.model_validate(arguments)).model_dump(
            mode="json"
        )
    if operation == EToolName.PREPARE_SEND:
        return prepare_send(backend, PreviewArguments.model_validate(arguments)).model_dump(
            mode="json"
        )
    if operation == EToolName.SEND_DRAFT:
        return send_draft(backend, SendArguments.model_validate(arguments)).model_dump(mode="json")
    raise OutlookError(EErrorCode.INVALID_ARGUMENT)


def handle_request(backend: Outlook, line: bytes, options: RuntimeOptions) -> str:
    backend.mutation_attempted = False
    try:
        request = WorkerRequest.model_validate_json(line)
        backend.references.prune()
        arguments = backend.references.arguments(request.arguments)
        result = execute(backend, request.operation, arguments, options)
        result = backend.references.output(result)
        return WorkerSuccess(result=result).model_dump_json()
    except Exception as error:
        if backend.mutation_attempted:
            mapped = OutlookError(EErrorCode.WRITE_OUTCOME_UNKNOWN)
        elif isinstance(error, OutlookError):
            mapped = error
        elif isinstance(error, ValidationError):
            mapped = OutlookError(EErrorCode.INVALID_ARGUMENT)
        else:
            mapped = (
                com_error(error)
                if hasattr(error, "hresult") or isinstance(error, AttributeError)
                else OutlookError(EErrorCode.INTERNAL_ERROR)
            )
        if mapped.code in {
            EErrorCode.OUTLOOK_UNAVAILABLE,
            EErrorCode.INTERNAL_ERROR,
            EErrorCode.WRITE_OUTCOME_UNKNOWN,
            EErrorCode.REFERENCE_LIMIT,
        }:
            backend.close()
        return WorkerFailure(error=mapped.info()).model_dump_json()


def run_worker(options: RuntimeOptions | None = None) -> None:
    initialized = False
    if sys.platform == "win32":
        import pythoncom

        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        initialized = True
    backend = Outlook()
    try:
        while line := sys.stdin.buffer.readline(MAX_IPC_BYTES + 1):
            if len(line) > MAX_IPC_BYTES:
                break
            response = handle_request(backend, line, options or RuntimeOptions())
            sys.stdout.buffer.write((response + "\n").encode("utf-8"))
            sys.stdout.buffer.flush()
    finally:
        backend.close()
        gc.collect()
        if initialized:
            pythoncom.CoUninitialize()

"""Private JSON pipe worker. All COM work runs on this process's main STA thread."""

import gc
import sys

from pydantic import JsonValue, ValidationError

from .config import MAX_IPC_BYTES
from .enums import EErrorCode, EToolName
from .errors import OutlookError, com_error
from .models import (
    FolderArguments,
    ReadArguments,
    RecentArguments,
    SearchArguments,
    TWorkerResponse,
    WorkerFailure,
    WorkerRequest,
    WorkerSuccess,
)
from .outlook import Outlook


def execute(
    backend: Outlook, operation: EToolName, arguments: dict[str, JsonValue]
) -> dict[str, JsonValue]:
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
    raise OutlookError(EErrorCode.INVALID_ARGUMENT)


def run_worker() -> None:
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
            response: TWorkerResponse
            try:
                request = WorkerRequest.model_validate_json(line)
                result = execute(backend, request.operation, request.arguments)
                response = WorkerSuccess(result=result)
            except ValidationError:
                response = WorkerFailure(error=OutlookError(EErrorCode.INVALID_ARGUMENT).info())
            except OutlookError as error:
                response = WorkerFailure(error=error.info())
                if error.code == EErrorCode.OUTLOOK_UNAVAILABLE:
                    backend.close()
            except Exception as error:
                mapped = (
                    com_error(error)
                    if hasattr(error, "hresult") or isinstance(error, AttributeError)
                    else OutlookError(EErrorCode.INTERNAL_ERROR)
                )
                response = WorkerFailure(error=mapped.info())
                backend.close()
            sys.stdout.buffer.write((response.model_dump_json() + "\n").encode("utf-8"))
            sys.stdout.buffer.flush()
    finally:
        backend.close()
        gc.collect()
        if initialized:
            pythoncom.CoUninitialize()

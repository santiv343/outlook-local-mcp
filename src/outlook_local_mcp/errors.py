"""Stable, deliberately content-free operational errors."""

from pydantic import JsonValue

from .enums import EErrorCode
from .error_constants import ERRORS
from .models import ErrorInfo
from .outlook_constants import (
    ACCESS_DENIED_HRESULTS,
    EXCEPINFO_STATUS_INDEX,
    HRESULT_MASK,
    NOT_FOUND_HRESULTS,
)


class OutlookError(Exception):
    def __init__(self, code: EErrorCode, message: str | None = None) -> None:
        default, self.retryable = ERRORS[code]
        self.code = code
        self.message = message or default
        super().__init__(self.message)

    def payload(self) -> dict[str, JsonValue]:
        return {"code": self.code, "message": self.message, "retryable": self.retryable}

    def info(self) -> ErrorInfo:
        return ErrorInfo(code=self.code, message=self.message, retryable=self.retryable)


def com_error(
    error: Exception, missing: EErrorCode = EErrorCode.OUTLOOK_UNAVAILABLE
) -> OutlookError:
    """Inspect numeric HRESULTs only; never forward an exception's text."""
    if isinstance(error, OutlookError):
        return error
    codes = {getattr(error, "hresult", None)}
    details = getattr(error, "excepinfo", None)
    if isinstance(details, tuple) and len(details) > EXCEPINFO_STATUS_INDEX:
        codes.add(details[EXCEPINFO_STATUS_INDEX])
    unsigned = {code & HRESULT_MASK for code in codes if isinstance(code, int)}
    if unsigned & ACCESS_DENIED_HRESULTS:
        return OutlookError(EErrorCode.ACCESS_DENIED)
    if unsigned & NOT_FOUND_HRESULTS:
        return OutlookError(missing)
    return OutlookError(EErrorCode.OUTLOOK_UNAVAILABLE)

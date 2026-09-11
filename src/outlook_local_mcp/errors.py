"""Stable, deliberately content-free operational errors."""

from pydantic import JsonValue

from .enums import EErrorCode
from .error_constants import ERRORS
from .models import ErrorInfo
from .outlook_constants import (
    ACCESS_DENIED_HRESULTS,
    DISPATCH_EXCEPTION_HRESULT,
    EXCEPINFO_STATUS_INDEX,
    HRESULT_MASK,
    MAPI_NOT_FOUND_HRESULT,
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


def com_hresult_codes(error: Exception) -> set[int]:
    """Inspect numeric HRESULTs only; never inspect an exception's text."""
    codes = [getattr(error, "hresult", None)]
    details = getattr(error, "excepinfo", None)
    if isinstance(details, tuple) and len(details) > EXCEPINFO_STATUS_INDEX:
        codes.append(details[EXCEPINFO_STATUS_INDEX])
    return {code & HRESULT_MASK for code in codes if isinstance(code, int)}


def is_missing_property(error: Exception) -> bool:
    """Require precise property absence, not the broader item-locator category."""
    status: object = getattr(error, "hresult", None)
    details: object = getattr(error, "excepinfo", None)
    if not isinstance(status, int):
        return False
    if details is None:
        return status & HRESULT_MASK == MAPI_NOT_FOUND_HRESULT
    if not isinstance(details, tuple) or len(details) != EXCEPINFO_STATUS_INDEX + 1:
        return False
    inner: object = details[EXCEPINFO_STATUS_INDEX]
    if not isinstance(inner, int) or inner & HRESULT_MASK != MAPI_NOT_FOUND_HRESULT:
        return False
    return com_hresult_codes(error) <= {MAPI_NOT_FOUND_HRESULT, DISPATCH_EXCEPTION_HRESULT}


def com_error(
    error: Exception, missing: EErrorCode = EErrorCode.OUTLOOK_UNAVAILABLE
) -> OutlookError:
    """Map operational errors without forwarding any raw COM text."""
    if isinstance(error, OutlookError):
        return error
    unsigned = com_hresult_codes(error)
    if unsigned & ACCESS_DENIED_HRESULTS:
        return OutlookError(EErrorCode.ACCESS_DENIED)
    if unsigned & NOT_FOUND_HRESULTS:
        return OutlookError(missing)
    return OutlookError(EErrorCode.OUTLOOK_UNAVAILABLE)

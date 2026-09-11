"""Explicit capability selection, shared by both process boundaries."""

from .enums import EErrorCode, EToolName
from .errors import OutlookError
from .models import RuntimeOptions

WRITE_TOOLS = frozenset({EToolName.OPEN_EMAIL, EToolName.CREATE_DRAFT, EToolName.REPLY_TO_EMAIL})
SEND_TOOLS = frozenset({EToolName.PREPARE_SEND, EToolName.SEND_DRAFT})
MUTATING_TOOLS = WRITE_TOOLS | {EToolName.SEND_DRAFT}


def enabled(operation: EToolName, options: RuntimeOptions) -> bool:
    if operation in SEND_TOOLS:
        return options.enable_write_tools and options.enable_send
    return operation not in WRITE_TOOLS or options.enable_write_tools


def require_capability(operation: EToolName, options: RuntimeOptions) -> None:
    if not enabled(operation, options):
        raise OutlookError(EErrorCode.CAPABILITY_DISABLED)

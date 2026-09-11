"""The same bounded connection check exposed through MCP and the package CLI."""

from .enums import EToolName
from .errors import OutlookError
from .models import OutlookStatus
from .supervisor import Supervisor


async def diagnose() -> OutlookStatus:
    supervisor = Supervisor()
    try:
        return OutlookStatus.model_validate(await supervisor.request(EToolName.OUTLOOK_STATUS, {}))
    except OutlookError as error:
        return OutlookStatus(available=False, error=error.info())
    finally:
        await supervisor.close()

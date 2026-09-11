import asyncio
import sys
import time

import pytest

from outlook_local_mcp.enums import EErrorCode, EToolName
from outlook_local_mcp.errors import OutlookError
from outlook_local_mcp.supervisor import Supervisor


def supervisor(timeout=3):
    return Supervisor(timeout=timeout, command=[sys.executable, "-m", "tests.fake_worker"])


async def wait_until(predicate):
    async with asyncio.timeout(5):
        while not predicate():
            await asyncio.sleep(0.01)


async def test_hard_timeout_reaps_worker_and_next_call_recovers(tmp_path):
    owner = supervisor(timeout=1)
    marker = tmp_path / "started"
    started = time.monotonic()
    task = asyncio.create_task(
        owner.request(EToolName.OUTLOOK_STATUS, {"delay": 60, "marker": str(marker)})
    )
    try:
        await wait_until(marker.exists)
        child = owner.process
        with pytest.raises(
            OutlookError, check=lambda error: error.code == EErrorCode.OUTLOOK_TIMEOUT
        ):
            await task
        assert time.monotonic() - started < 4
        assert child.returncode is not None and owner.process is None
        result = await owner.request(EToolName.OUTLOOK_STATUS, {})
        assert result["pid"] != child.pid
    finally:
        await owner.close()


async def test_queue_is_bounded_and_queued_timeout_does_not_kill_active_worker(tmp_path):
    owner = supervisor()
    marker = tmp_path / "started"
    active = asyncio.create_task(
        owner.request(EToolName.OUTLOOK_STATUS, {"delay": 0.6, "marker": str(marker)})
    )
    try:
        await wait_until(marker.exists)
        child = owner.process
        owner.timeout = 0.15
        queued = asyncio.create_task(owner.request(EToolName.OUTLOOK_STATUS, {}))
        await wait_until(lambda: owner.pending == 2)
        with pytest.raises(OutlookError, check=lambda error: error.code == EErrorCode.SERVER_BUSY):
            await owner.request(EToolName.OUTLOOK_STATUS, {})
        with pytest.raises(
            OutlookError, check=lambda error: error.code == EErrorCode.OUTLOOK_TIMEOUT
        ):
            await queued
        assert child.returncode is None
        assert (await active)["pid"] == child.pid
    finally:
        await owner.close()


async def test_cancellation_reaps_worker_and_clears_admission(tmp_path):
    owner = supervisor()
    marker = tmp_path / "started"
    task = asyncio.create_task(
        owner.request(EToolName.OUTLOOK_STATUS, {"delay": 60, "marker": str(marker)})
    )
    try:
        await wait_until(marker.exists)
        child = owner.process
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task
        assert child.returncode is not None
        assert owner.pending == 0 and not owner.lock.locked()
        assert (await owner.request(EToolName.OUTLOOK_STATUS, {}))["available"]
    finally:
        await owner.close()


async def test_worker_crash_is_sanitized_and_recoverable():
    owner = supervisor()
    try:
        with pytest.raises(
            OutlookError, check=lambda error: error.code == EErrorCode.INTERNAL_ERROR
        ):
            await owner.request(EToolName.OUTLOOK_STATUS, {"crash": True})
        assert (await owner.request(EToolName.OUTLOOK_STATUS, {}))["available"]
    finally:
        await owner.close()

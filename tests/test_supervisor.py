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
    owner = supervisor(timeout=5)
    await owner.request(EToolName.OUTLOOK_STATUS, {})
    owner.timeout = 0.25
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
        owner.timeout = 5
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


async def test_close_during_process_creation_reaps_new_worker(monkeypatch):
    owner = supervisor()
    spawn_started = asyncio.Event()
    release_spawn = asyncio.Event()
    original_spawn = asyncio.create_subprocess_exec
    created = []

    async def delayed_spawn(*arguments, **options):
        spawn_started.set()
        await release_spawn.wait()
        child = await original_spawn(*arguments, **options)
        created.append(child)
        return child

    monkeypatch.setattr(asyncio, "create_subprocess_exec", delayed_spawn)
    task = asyncio.create_task(owner.request(EToolName.OUTLOOK_STATUS, {}))
    await spawn_started.wait()
    await owner.close()
    release_spawn.set()
    with pytest.raises(
        OutlookError, check=lambda error: error.code == EErrorCode.OUTLOOK_UNAVAILABLE
    ):
        await task
    assert owner.process is None and owner.pending == 0
    assert len(created) == 1 and created[0].returncode is not None


@pytest.mark.parametrize("failure", ["timeout", "crash", "cancel"])
async def test_dispatched_mutation_failure_is_unknown_and_never_retried(tmp_path, failure):
    owner = supervisor(timeout=5)
    await owner.request(EToolName.OUTLOOK_STATUS, {})
    child = owner.process
    marker = tmp_path / "dispatched"
    arguments = {"marker": str(marker), "crash": failure == "crash", "delay": 60}
    if failure == "timeout":
        owner.timeout = 0.3
    task = asyncio.create_task(owner.request(EToolName.SEND_DRAFT, arguments))
    try:
        await wait_until(marker.exists)
        if failure == "cancel":
            task.cancel()
        with pytest.raises(
            OutlookError,
            check=lambda error: (
                error.code == EErrorCode.WRITE_OUTCOME_UNKNOWN and not error.info().retryable
            ),
        ):
            await task
        assert owner.process is None and child.returncode is not None
        assert owner.pending == 0
        owner.timeout = 5
        assert (await owner.request(EToolName.OUTLOOK_STATUS, {}))["available"]
    finally:
        await owner.close()

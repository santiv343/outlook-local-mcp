"""Bounded admission and hard process isolation for blocking COM calls."""

import asyncio
import os
import subprocess
import sys
from collections.abc import Sequence
from contextlib import suppress

from pydantic import JsonValue, TypeAdapter

from .capabilities import MUTATING_TOOLS
from .config import (
    MAX_IPC_BYTES,
    MAX_PENDING_OPERATIONS,
    OPERATION_TIMEOUT,
    WORKER_EXIT_TIMEOUT,
)
from .enums import EErrorCode, EToolName
from .errors import OutlookError
from .models import RuntimeOptions, TWorkerResponse, WorkerFailure, WorkerRequest


class Supervisor:
    def __init__(
        self,
        timeout: float = OPERATION_TIMEOUT,
        command: Sequence[str] | None = None,
        options: RuntimeOptions | None = None,
    ) -> None:
        self.timeout = timeout
        self.options = options or RuntimeOptions()
        self.command = list(
            command
            or [sys.executable, "-m", "outlook_local_mcp", "--worker", *self.options.flags()]
        )
        self.process: asyncio.subprocess.Process | None = None
        self.lock = asyncio.Lock()
        self.pending = 0
        self.closed = False

    async def _start(self) -> asyncio.subprocess.Process:
        if self.process is None or self.process.returncode is not None:
            flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            command = list(self.command)
            environment = {**os.environ, "PYTHONUTF8": "1"}
            if sys.platform == "win32" and command[0] == sys.executable:
                # CPython's Windows venv executable is a subprocess launcher.
                # Start its real interpreter directly, preserving venv discovery,
                # so the owned process is the one that actually executes COM.
                base_executable: object = getattr(sys, "_base_executable", None)
                if not isinstance(base_executable, str) or not base_executable:
                    raise OutlookError(
                        EErrorCode.INTERNAL_ERROR, "Cannot locate the CPython worker interpreter."
                    )
                command[0] = base_executable
                environment["__PYVENV_LAUNCHER__"] = sys.executable
            spawning = asyncio.create_task(
                asyncio.create_subprocess_exec(
                    *command,
                    stdin=asyncio.subprocess.PIPE,
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.DEVNULL,
                    env=environment,
                    creationflags=flags,
                    limit=MAX_IPC_BYTES,
                )
            )
            try:
                self.process = await asyncio.shield(spawning)
            except asyncio.CancelledError:
                # Retain ownership if cancellation races with process creation.
                self.process = await spawning
                raise
        return self.process

    async def _stop(self) -> None:
        process = self.process
        if process is None:
            return
        if process.returncode is None:
            with suppress(ProcessLookupError):
                process.kill()
        try:
            await asyncio.wait_for(process.wait(), WORKER_EXIT_TIMEOUT)
        except TimeoutError:
            raise OutlookError(
                EErrorCode.INTERNAL_ERROR, "The worker did not exit within the cleanup deadline."
            ) from None
        if self.process is process:
            self.process = None

    async def _stop_after_failure(self) -> None:
        try:
            await self._stop()
        except (Exception, asyncio.CancelledError):
            # Preserve the operation's outcome and retain ownership for close().
            # Never start another COM worker while cleanup is unconfirmed.
            self.closed = True

    async def close(self) -> None:
        self.closed = True
        await self._stop()

    async def request(
        self, operation: EToolName, arguments: dict[str, JsonValue]
    ) -> dict[str, JsonValue]:
        if self.closed:
            raise OutlookError(EErrorCode.OUTLOOK_UNAVAILABLE)
        if self.pending >= MAX_PENDING_OPERATIONS:
            raise OutlookError(EErrorCode.SERVER_BUSY)
        self.pending += 1
        acquired = False
        dispatched_mutation = False
        try:
            async with asyncio.timeout(self.timeout):
                await self.lock.acquire()
                acquired = True
                if self.closed:
                    raise OutlookError(EErrorCode.OUTLOOK_UNAVAILABLE)
                process = await self._start()
                if self.closed:
                    await self._stop()
                    raise OutlookError(EErrorCode.OUTLOOK_UNAVAILABLE)
                writer, reader = process.stdin, process.stdout
                if writer is None or reader is None:
                    raise OutlookError(EErrorCode.INTERNAL_ERROR, "Worker pipes were not created.")
                message = (
                    WorkerRequest(operation=operation, arguments=arguments).model_dump_json() + "\n"
                )
                dispatched_mutation = operation in MUTATING_TOOLS
                writer.write(message.encode("utf-8"))
                await writer.drain()
                line = await reader.readline()
                if not line:
                    raise BrokenPipeError
                response: TWorkerResponse = TypeAdapter(TWorkerResponse).validate_json(line)
                if isinstance(response, WorkerFailure):
                    raise OutlookError(response.error.code, response.error.message)
                return response.result
        except TimeoutError:
            if acquired:
                await self._stop_after_failure()
            raise OutlookError(
                EErrorCode.WRITE_OUTCOME_UNKNOWN
                if dispatched_mutation
                else EErrorCode.OUTLOOK_TIMEOUT
            ) from None
        except asyncio.CancelledError:
            if acquired:
                await self._stop_after_failure()
            if dispatched_mutation:
                raise OutlookError(EErrorCode.WRITE_OUTCOME_UNKNOWN) from None
            raise
        except OutlookError:
            raise
        except Exception:
            if acquired:
                await self._stop_after_failure()
            raise OutlookError(
                EErrorCode.WRITE_OUTCOME_UNKNOWN
                if dispatched_mutation
                else EErrorCode.INTERNAL_ERROR
            ) from None
        finally:
            if acquired:
                self.lock.release()
            self.pending -= 1

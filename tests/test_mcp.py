import asyncio
import json
import os
import sys
import time

import pytest
from mcp import Client, MCPError
from mcp.client.stdio import StdioServerParameters, stdio_client

from .test_supervisor import wait_until


async def test_stdio_discovery_schemas_errors_and_private_logs(tmp_path):
    private = "synthetic-query-that-must-never-be-logged"
    error_log = tmp_path / "stderr.txt"
    with error_log.open("w+") as error_stream:
        async with Client(
            stdio_client(
                StdioServerParameters(command=sys.executable, args=["-m", "tests.mcp_host"]),
                errlog=error_stream,
            )
        ) as client:
            listed = await client.list_tools()
            assert {tool.name for tool in listed.tools} == {
                "outlook_status",
                "list_mailboxes",
                "list_folders",
                "recent_emails",
                "search_emails",
                "read_email",
                "read_conversation",
            }
            assert all(tool.annotations.read_only_hint for tool in listed.tools)
            assert all(tool.input_schema and tool.output_schema for tool in listed.tools)
            status = await client.call_tool("outlook_status", {})
            assert not status.is_error and status.structured_content["available"]
            for tool, arguments in [
                ("search_emails", {"query": private, "limit": 0}),
                ("search_emails", {"query": private, "after": "not-a-date"}),
                ("recent_emails", {"folder_id": private}),
                ("outlook_status", {"unexpected": private}),
            ]:
                result = await client.call_tool(tool, arguments)
                assert result.is_error
                assert result.structured_content["code"] == "INVALID_ARGUMENT"
                assert private not in str(result)
            with pytest.raises(MCPError):
                await client.call_tool(private, {})
    logs = error_log.read_text()
    assert "operation=outlook_status" in logs and "code=INVALID_ARGUMENT" in logs
    assert private not in logs


async def test_invalid_worker_output_is_tool_error_and_sanitized(tmp_path):
    path = tmp_path / "errors.txt"
    with path.open("w+") as error_stream:
        async with Client(
            stdio_client(
                StdioServerParameters(
                    command=sys.executable, args=["-m", "tests.mcp_host", "invalid"]
                ),
                errlog=error_stream,
            )
        ) as client:
            result = await client.call_tool("outlook_status", {})
            assert result.is_error and result.structured_content["code"] == "INTERNAL_ERROR"
            assert "synthetic-private-response" not in str(result)
    assert "synthetic-private-response" not in path.read_text()


@pytest.mark.parametrize("flags,count", [([], 7), (["writes"], 10), (["writes", "send"], 12)])
async def test_capability_discovery_and_mutation_output_errors(flags, count):
    async with Client(
        StdioServerParameters(command=sys.executable, args=["-m", "tests.mcp_host", *flags])
    ) as client:
        listed = await client.list_tools()
        assert len(listed.tools) == count
        names = {tool.name: tool for tool in listed.tools}
        if "writes" in flags:
            assert not names["open_email"].annotations.read_only_hint
            failed = await client.call_tool(
                "open_email", {"entry_id": "synthetic", "store_id": "store"}
            )
            assert failed.is_error and failed.structured_content["code"] == "WRITE_OUTCOME_UNKNOWN"
            assert not failed.structured_content["retryable"]
        else:
            denied = await client.call_tool(
                "open_email", {"entry_id": "synthetic", "store_id": "store"}
            )
            assert denied.is_error and denied.structured_content["code"] == "CAPABILITY_DISABLED"
        if "send" in flags:
            assert names["send_draft"].annotations.destructive_hint
        else:
            denied = await client.call_tool("send_draft", {})
            assert denied.is_error and denied.structured_content["code"] == "CAPABILITY_DISABLED"


async def test_client_eof_reaps_active_worker_promptly(tmp_path):
    import win32api
    import win32con
    import win32event

    handle = None
    marker = tmp_path / "worker-pid"
    # Bypass the Windows venv launcher for the test host as the production
    # supervisor does for its worker. This makes process ownership observable.
    executable = sys._base_executable if sys.platform == "win32" else sys.executable
    environment = {**os.environ, "__PYVENV_LAUNCHER__": sys.executable, "PYTHONUTF8": "1"}
    process = await asyncio.create_subprocess_exec(
        executable,
        "-m",
        "tests.mcp_host",
        "block",
        str(marker),
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.DEVNULL,
        env=environment,
    )

    async def send(payload):
        process.stdin.write((json.dumps({"jsonrpc": "2.0", **payload}) + "\n").encode())
        await process.stdin.drain()

    try:
        await send(
            {
                "id": 1,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2025-06-18",
                    "capabilities": {},
                    "clientInfo": {"name": "synthetic-eof-test", "version": "1"},
                },
            }
        )
        initialized = json.loads(await asyncio.wait_for(process.stdout.readline(), 10))
        assert "result" in initialized
        await send({"method": "notifications/initialized"})
        await send({"id": 2, "method": "tools/call", "params": {"name": "outlook_status"}})
        await wait_until(marker.exists)
        worker_pid = int(marker.read_text())
        handle = win32api.OpenProcess(
            win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, worker_pid
        )
        started = time.monotonic()
        process.stdin.close()
        await asyncio.wait_for(process.wait(), 4)
        assert time.monotonic() - started < 4
        assert win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_OBJECT_0
    finally:
        if process.returncode is None:
            process.kill()
            await process.wait()
        if handle is not None:
            if win32event.WaitForSingleObject(handle, 0) != win32event.WAIT_OBJECT_0:
                win32api.TerminateProcess(handle, 1)
            handle.Close()

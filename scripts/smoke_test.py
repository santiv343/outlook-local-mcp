"""Real MCP integration; never print or persist mailbox data."""

import argparse
import asyncio
import sys
from typing import Any

from mcp import Client
from mcp.client.stdio import StdioServerParameters

EXPECTED_TOOLS = {
    "outlook_status",
    "list_mailboxes",
    "list_folders",
    "recent_emails",
    "search_emails",
    "read_email",
}


async def run(command: list[str]) -> None:
    async with Client(
        StdioServerParameters(command=command[0], args=command[1:]), read_timeout_seconds=45
    ) as client:

        async def call(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
            result = await client.call_tool(name, arguments)
            if result.is_error:
                payload = result.structured_content
                code = payload.get("code", "UNKNOWN") if isinstance(payload, dict) else "UNKNOWN"
                raise RuntimeError(f"{name}:{code}")
            assert isinstance(result.structured_content, dict)
            return result.structured_content

        tools = await client.list_tools()
        assert {tool.name for tool in tools.tools} == EXPECTED_TOOLS
        assert all(tool.annotations and tool.annotations.read_only_hint for tool in tools.tools)
        print("PASS discovery tools=6")
        status = await call("outlook_status", {})
        assert status["available"]
        print("PASS outlook_status")
        mailboxes = await call("list_mailboxes", {})
        assert mailboxes["items"]
        selected = next(
            (item for item in mailboxes["items"] if item["is_default"]), mailboxes["items"][0]
        )
        folders = await call("list_folders", {"store_id": selected["store_id"], "limit": 5})
        print(f"PASS mailbox_and_folder_listing folders={len(folders['items'])}")
        recent = await call("recent_emails", {"limit": 5})
        assert recent["items"], "A nonempty Inbox is required for the real smoke test"
        print(f"PASS recent_emails count={len(recent['items'])}")
        reference = recent["items"][0]
        search = await call(
            "search_emails",
            {
                "query": reference["subject"],
                "store_id": reference["store_id"],
                "folder_id": reference["folder_id"],
                "limit": 5,
            },
        )
        assert any(item["entry_id"] == reference["entry_id"] for item in search["items"])
        print("PASS known_subject_search")
        identifiers = {"entry_id": reference["entry_id"], "store_id": reference["store_id"]}
        first = await call("read_email", {**identifiers, "body_limit": 128})
        assert first["unread"] == reference["unread"]
        if first["body_truncated"]:
            following = await call(
                "read_email",
                {**identifiers, "body_offset": first["next_body_offset"], "body_limit": 128},
            )
            assert following["body_offset"] == first["next_body_offset"]
            assert following["unread"] == reference["unread"]
            print("PASS body_continuation")
        else:
            print("SKIP body_continuation reason=short_body")
        reread = await call("read_email", {**identifiers, "body_limit": 1})
        assert reread["unread"] == reference["unread"]
        print("PASS read_and_preserve_unread")
        invalid = await client.call_tool("recent_emails", {"limit": 0})
        assert invalid.is_error
        assert invalid.structured_content["code"] == "INVALID_ARGUMENT"
        missing = await client.call_tool(
            "read_email", {"entry_id": "0" * 64, "store_id": reference["store_id"]}
        )
        assert missing.is_error
        assert missing.structured_content["code"] == "ITEM_NOT_FOUND"
        print("PASS invalid_arguments_and_missing_item")
    print("PASS clean_client_shutdown")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--server-command", nargs=argparse.REMAINDER)
    options = parser.parse_args()
    try:
        asyncio.run(run(options.server_command or [sys.executable, "-m", "outlook_local_mcp"]))
    except Exception as error:
        # Only our own fixed-name/code errors are printed, never SDK or COM exception text.
        safe = (
            str(error)
            if type(error) is RuntimeError and str(error).split(":")[0] in EXPECTED_TOOLS
            else type(error).__name__
        )
        print(f"FAIL {safe}")
        sys.exit(1)

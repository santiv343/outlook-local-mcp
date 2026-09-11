"""Real MCP checks. Creates one synthetic draft and opens it; never transmits mail.

The draft remains in Outlook for review. Do not blindly rerun after an unknown
mutation outcome: inspect Drafts first. No mailbox data is printed or persisted.
"""

import argparse
import asyncio
import sys

from mcp import Client
from mcp.client.stdio import StdioServerParameters


async def run(command):
    async with Client(
        StdioServerParameters(command=command[0], args=command[1:]), read_timeout_seconds=45
    ) as client:

        async def call(name, arguments):
            response = await client.call_tool(name, arguments)
            if response.is_error:
                code = response.structured_content.get("code", "UNKNOWN")
                print(f"FAIL operation={name} code={code}", flush=True)
                raise RuntimeError
            return response.structured_content

        tools = await client.list_tools()
        assert len(tools.tools) == 12
        print("PASS discovery tools=12", flush=True)
        recent = await call("recent_emails", {"limit": 5})
        reference = recent["items"][0]
        identifiers = {"entry_id": reference["entry_id"], "store_id": reference["store_id"]}
        conversation = await call(
            "read_conversation", {**identifiers, "limit": 1, "body_limit": 64}
        )
        assert conversation["items"]
        if conversation["next_cursor"]:
            await call(
                "read_conversation",
                {
                    **identifiers,
                    "limit": 1,
                    "body_limit": 64,
                    "cursor": conversation["next_cursor"],
                },
            )
        print("PASS native_conversation_and_continuation", flush=True)
        detail = await call("read_email", {**identifiers, "body_limit": 1})
        arguments = {
            "query": reference["subject"],
            "importance": reference["importance"],
            "limit": 5,
        }
        if detail["recipients"]:
            person = detail["recipients"][0]
            arguments["recipient"] = person["email"] or person["name"]
        if reference["categories"]:
            arguments["category"] = reference["categories"][0]
        if detail["attachments"]:
            arguments["attachment_name"] = detail["attachments"][0]["name"]
        matches = await call("search_emails", arguments)
        assert any(item["entry_id"] == reference["entry_id"] for item in matches["items"])
        print("PASS combined_available_metadata_filters", flush=True)
        draft = await call(
            "create_draft",
            {
                "to": ["recipient@example.com"],
                "subject": "Outlook MCP synthetic draft verification - do not send",
                "body": "Synthetic local MCP verification. This draft must not be sent.",
            },
        )
        draft_ids = {"entry_id": draft["entry_id"], "store_id": draft["store_id"]}
        print("PASS saved_synthetic_draft drafts_created=1", flush=True)
        preview = await call(
            "prepare_send", {**draft_ids, "account_email": draft["account"]["email"]}
        )
        assert preview["subject"] == draft["subject"] and preview["body"] == draft["body"]
        assert preview["recipients"][0]["email"] == "recipient@example.com"
        print("PASS complete_send_preview no_submission", flush=True)
        opened = await call("open_email", draft_ids)
        assert opened["opened"]
        print("PASS open_draft_in_outlook", flush=True)
        reread = await call("read_email", {**identifiers, "body_limit": 1})
        assert reread["unread"] == reference["unread"]
        print("PASS original_read_state_preserved", flush=True)
    print("PASS clean_shutdown synthetic_drafts_retained=1 live_send=not_exercised", flush=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--server-command", nargs=argparse.REMAINDER)
    options = parser.parse_args()
    try:
        asyncio.run(
            run(
                options.server_command
                or [
                    sys.executable,
                    "-m",
                    "outlook_local_mcp",
                    "--enable-write-tools",
                    "--enable-send",
                ]
            )
        )
    except Exception:
        print("FAIL action_smoke; inspect previous step and Outlook before repeating")
        sys.exit(1)

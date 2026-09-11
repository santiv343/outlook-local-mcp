"""Package entry points. The default command speaks MCP stdio only."""

import argparse
import asyncio
import json
import sys
from pathlib import Path

from . import __version__
from .config import CLIENT_SERVER_NAME, ENABLE_SEND_FLAG, ENABLE_WRITE_FLAG
from .errors import OutlookError
from .models import RuntimeOptions


def main() -> None:
    parser = argparse.ArgumentParser(description="Classic Outlook MCP server; read-only by default")
    parser.add_argument("command", nargs="?", choices=["doctor", "config", "configure-claude"])
    parser.add_argument("--config-path", type=Path, help="Explicit Claude configuration file")
    parser.add_argument(
        "--replace", action="store_true", help="Replace a different outlook-local entry"
    )
    parser.add_argument(
        "--dry-run", action="store_true", help="Check configuration without writing"
    )
    parser.add_argument("--version", action="version", version=__version__)
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument(
        ENABLE_WRITE_FLAG, action="store_true", help="Enable opening mail and saving drafts/replies"
    )
    parser.add_argument(
        ENABLE_SEND_FLAG,
        action="store_true",
        help="Also enable previewed draft submission; requires --enable-write-tools",
    )
    arguments = parser.parse_args()
    if arguments.enable_send and not arguments.enable_write_tools:
        parser.error("--enable-send requires --enable-write-tools")
    options = RuntimeOptions(
        enable_write_tools=arguments.enable_write_tools, enable_send=arguments.enable_send
    )
    if arguments.worker:
        from .worker import run_worker

        run_worker(options)
    elif arguments.command == "doctor":
        from .doctor import diagnose

        status = asyncio.run(diagnose())
        print(status.model_dump_json(indent=2))
        sys.exit(0 if status.available else 1)
    elif arguments.command in {"config", "configure-claude"}:
        from .client_config import claude_config_path, configure_claude, server_entry

        try:
            entry = server_entry(options=options)
            if arguments.command == "config":
                print(json.dumps({"mcpServers": {CLIENT_SERVER_NAME: entry}}, indent=2))
            else:
                result = configure_claude(
                    arguments.config_path or claude_config_path(),
                    entry,
                    replace=arguments.replace,
                    dry_run=arguments.dry_run,
                )
                print(result.model_dump_json(indent=2))
        except OutlookError as error:
            print(json.dumps(error.payload()), file=sys.stderr)
            sys.exit(1)
        except OSError:
            print(
                "Configuration could not be read or written. Check file permissions.",
                file=sys.stderr,
            )
            sys.exit(1)
    else:
        from .server import serve

        asyncio.run(serve(options=options))


if __name__ == "__main__":
    main()

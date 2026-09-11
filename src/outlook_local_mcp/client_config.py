"""Optional client configuration; the MCP server itself is client-independent."""

import json
import os
import shutil
import tempfile
from pathlib import Path

from pydantic import JsonValue, TypeAdapter

from . import __version__
from .config import CLIENT_SERVER_NAME, RELEASE_BASE, RUNTIME_PYTHON
from .enums import EErrorCode
from .errors import OutlookError
from .models import ConfigurationResult, RuntimeOptions


def server_entry(
    launcher: str | None = None, options: RuntimeOptions | None = None
) -> dict[str, JsonValue]:
    command = launcher or shutil.which("uvx")
    if command is None:
        raise OutlookError(
            EErrorCode.INVALID_ARGUMENT, "Install uv and make uvx available on PATH."
        )
    release = f"{RELEASE_BASE}/v{__version__}"
    return {
        "command": str(Path(command).resolve()),
        "args": [
            "--python",
            RUNTIME_PYTHON,
            "--constraints",
            f"{release}/constraints.txt",
            "--from",
            f"{release}/outlook_local_mcp-{__version__}-py3-none-any.whl",
            "outlook-local-mcp",
            *(options or RuntimeOptions()).flags(),
        ],
    }


def claude_config_path() -> Path:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        raise OutlookError(EErrorCode.UNSUPPORTED_PLATFORM)
    conventional = Path(appdata) / "Claude" / "claude_desktop_config.json"
    candidates = [conventional]
    local_appdata = os.environ.get("LOCALAPPDATA")
    if local_appdata:
        packages = Path(local_appdata) / "Packages"
        if packages.is_dir():
            candidates.extend(
                path / "LocalCache" / "Roaming" / "Claude" / "claude_desktop_config.json"
                for path in packages.iterdir()
                if path.is_dir() and path.name.startswith("Claude_")
            )
    existing = [path for path in candidates if path.is_file()]
    if len(existing) > 1:
        raise OutlookError(
            EErrorCode.INVALID_ARGUMENT,
            "Multiple Claude configurations exist. Select the active file using --config-path.",
        )
    if existing:
        return existing[0]
    installed = [path for path in candidates if path.parent.is_dir()]
    if len(installed) == 1:
        return installed[0]
    return conventional


def configure_claude(
    path: Path,
    entry: dict[str, JsonValue],
    *,
    replace: bool = False,
    dry_run: bool = False,
) -> ConfigurationResult:
    original = path.read_bytes() if path.exists() else None
    try:
        configuration = (
            TypeAdapter(dict[str, JsonValue]).validate_json(original)
            if original is not None
            else {}
        )
    except ValueError:
        raise OutlookError(
            EErrorCode.INVALID_ARGUMENT,
            "Claude configuration is not a valid JSON object; it was not changed.",
        ) from None
    servers = configuration.get("mcpServers", {})
    if not isinstance(servers, dict):
        raise OutlookError(
            EErrorCode.INVALID_ARGUMENT,
            "mcpServers must be a JSON object; configuration was not changed.",
        )
    current = servers.get(CLIENT_SERVER_NAME)
    if current == entry:
        return ConfigurationResult(changed=False, backup_created=False)
    if CLIENT_SERVER_NAME in servers and not replace:
        raise OutlookError(
            EErrorCode.INVALID_ARGUMENT,
            "A different outlook-local entry exists. Review it and use --replace explicitly.",
        )
    servers[CLIENT_SERVER_NAME] = entry
    configuration["mcpServers"] = servers
    if dry_run:
        return ConfigurationResult(changed=True, backup_created=False)
    path.parent.mkdir(parents=True, exist_ok=True)
    if original is not None:
        backup_fd, _ = tempfile.mkstemp(prefix=path.name + ".backup-", dir=path.parent)
        with os.fdopen(backup_fd, "wb") as backup:
            backup.write(original)
            backup.flush()
            os.fsync(backup.fileno())
    descriptor, temporary_name = tempfile.mkstemp(prefix=path.name + ".tmp-", dir=path.parent)
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as target:
            json.dump(configuration, target, ensure_ascii=False, indent=2)
            target.write("\n")
            target.flush()
            os.fsync(target.fileno())
        current_bytes = path.read_bytes() if path.exists() else None
        if current_bytes != original:
            raise OutlookError(
                EErrorCode.INVALID_ARGUMENT,
                "Configuration changed during the update; retry after closing its editor.",
            )
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
    return ConfigurationResult(changed=True, backup_created=original is not None)

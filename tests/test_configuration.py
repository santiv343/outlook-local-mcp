import json
from pathlib import Path

import pytest

from outlook_local_mcp import client_config
from outlook_local_mcp.client_config import claude_config_path, configure_claude, server_entry
from outlook_local_mcp.errors import OutlookError


def test_preserve_backup_idempotence_and_dry_run(tmp_path):
    path = tmp_path / "settings.json"
    original = b'{"theme":"dark","mcpServers":{"existing":{"command":"example"}}}'
    path.write_bytes(original)
    entry = {"command": "uvx", "args": ["synthetic-package"]}
    preview = configure_claude(path, entry, dry_run=True)
    assert preview.changed and not preview.backup_created
    assert path.read_bytes() == original
    assert list(tmp_path.iterdir()) == [path]
    result = configure_claude(path, entry)
    assert result.changed and result.backup_created
    updated = json.loads(path.read_bytes())
    assert updated["theme"] == "dark"
    assert updated["mcpServers"]["existing"] == {"command": "example"}
    backups = list(tmp_path.glob("*.backup-*"))
    assert len(backups) == 1 and backups[0].read_bytes() == original
    assert not configure_claude(path, entry).changed
    assert len(list(tmp_path.glob("*.backup-*"))) == 1


@pytest.mark.parametrize("original", [b"not json", b"[]", b'{"mcpServers":null}'])
def test_invalid_json_does_not_overwrite_or_backup(tmp_path, original):
    path = tmp_path / "settings.json"
    path.write_bytes(original)
    with pytest.raises(OutlookError):
        configure_claude(path, {"command": "uvx"})
    assert path.read_bytes() == original
    assert list(tmp_path.iterdir()) == [path]


def test_conflicting_entry_requires_explicit_replace(tmp_path):
    path = tmp_path / "settings.json"
    original = b'{"mcpServers":{"outlook-local":{"command":"different"}}}'
    path.write_bytes(original)
    with pytest.raises(OutlookError):
        configure_claude(path, {"command": "uvx"})
    assert path.read_bytes() == original
    assert configure_claude(path, {"command": "uvx"}, replace=True).changed


def test_atomic_replace_failure_keeps_original(tmp_path, monkeypatch):
    path = tmp_path / "settings.json"
    path.write_bytes(b"{}")

    def fail_replace(source, target):
        raise PermissionError("synthetic file lock")

    monkeypatch.setattr(client_config.os, "replace", fail_replace)
    with pytest.raises(PermissionError):
        configure_claude(path, {"command": "uvx"})
    assert path.read_bytes() == b"{}"
    assert not list(tmp_path.glob("*.tmp-*"))


def test_packaged_claude_path_and_ambiguous_configs(tmp_path, monkeypatch):
    roaming = tmp_path / "roaming"
    local = tmp_path / "local"
    monkeypatch.setenv("APPDATA", str(roaming))
    monkeypatch.setenv("LOCALAPPDATA", str(local))
    packaged = (
        local / "Packages/Claude_example/LocalCache/Roaming/Claude/claude_desktop_config.json"
    )
    packaged.parent.mkdir(parents=True)
    packaged.write_text("{}")
    assert claude_config_path() == packaged
    classic = roaming / "Claude/claude_desktop_config.json"
    classic.parent.mkdir(parents=True)
    classic.write_text("{}")
    with pytest.raises(OutlookError):
        claude_config_path()


def test_entry_uses_absolute_launcher_and_versioned_release_constraints(tmp_path):
    entry = server_entry(str(tmp_path / "uvx.exe"))
    assert Path(entry["command"]).is_absolute()
    assert "--constraints" in entry["args"]
    assert "--from" in entry["args"]
    assert any(argument.endswith(".whl") for argument in entry["args"])

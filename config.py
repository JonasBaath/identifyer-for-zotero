"""
config.py
Persistent user settings for Identifyer for Zotero.

Stored as a single JSON file in the platform-appropriate user config
directory:
  - macOS:   ~/Library/Application Support/identifyer-for-zotero/settings.json
  - Windows: %APPDATA%/identifyer-for-zotero/settings.json
  - Linux:   $XDG_CONFIG_HOME/identifyer-for-zotero/settings.json
             (or ~/.config/identifyer-for-zotero/settings.json)

Readers (``load_settings``) never raise — a missing or malformed file
returns an empty dict so the app can fall back to platform defaults.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any, Dict

_APP_NAME = "identifyer-for-zotero"


def _config_dir() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / _APP_NAME
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        return Path(base) / _APP_NAME
    base = os.environ.get("XDG_CONFIG_HOME") or str(Path.home() / ".config")
    return Path(base) / _APP_NAME


def config_path() -> Path:
    return _config_dir() / "settings.json"


def load_settings() -> Dict[str, Any]:
    """Return the saved settings dict, or {} if the file is missing/invalid."""
    path = config_path()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(settings: Dict[str, Any]) -> None:
    """Write the settings dict atomically. Silently ignores write failures
    so a read-only config directory cannot crash the app."""
    path = config_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".json.tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
        os.replace(tmp, path)
    except OSError:
        pass


def update_setting(key: str, value: Any) -> None:
    """Load current settings, update one key, and save."""
    settings = load_settings()
    settings[key] = value
    save_settings(settings)

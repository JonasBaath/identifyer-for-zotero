"""Tests for config.py — persistent user settings."""
import json
import os
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

import config


@pytest.fixture
def tmp_config_dir(tmp_path, monkeypatch):
    """Redirect config_path() to a temporary directory for each test."""
    tmp_file = tmp_path / "settings.json"
    monkeypatch.setattr(config, "config_path", lambda: tmp_file)
    return tmp_file


class TestLoadSettings:
    def test_missing_file_returns_empty_dict(self, tmp_config_dir):
        assert config.load_settings() == {}

    def test_valid_json_returned(self, tmp_config_dir):
        tmp_config_dir.write_text(
            '{"db_path": "/foo/bar.sqlite"}', encoding="utf-8"
        )
        assert config.load_settings() == {"db_path": "/foo/bar.sqlite"}

    def test_malformed_json_returns_empty_dict(self, tmp_config_dir):
        tmp_config_dir.write_text("{not valid json", encoding="utf-8")
        assert config.load_settings() == {}

    def test_non_dict_json_returns_empty_dict(self, tmp_config_dir):
        tmp_config_dir.write_text("[1, 2, 3]", encoding="utf-8")
        assert config.load_settings() == {}


class TestSaveSettings:
    def test_save_creates_parent_dir(self, tmp_path, monkeypatch):
        nested = tmp_path / "deep" / "nested" / "settings.json"
        monkeypatch.setattr(config, "config_path", lambda: nested)
        config.save_settings({"db_path": "/a.sqlite"})
        assert nested.exists()
        assert json.loads(nested.read_text(encoding="utf-8")) == {"db_path": "/a.sqlite"}

    def test_save_overwrites(self, tmp_config_dir):
        config.save_settings({"db_path": "/a.sqlite"})
        config.save_settings({"db_path": "/b.sqlite"})
        assert json.loads(tmp_config_dir.read_text(encoding="utf-8")) == {"db_path": "/b.sqlite"}

    def test_save_write_failure_is_silent(self, tmp_path, monkeypatch):
        # Point config_path at a path whose parent is a file, so mkdir raises.
        blocker = tmp_path / "blocker"
        blocker.write_text("file")
        monkeypatch.setattr(
            config, "config_path", lambda: blocker / "sub" / "settings.json"
        )
        # Must not raise.
        config.save_settings({"k": "v"})


class TestUpdateSetting:
    def test_update_preserves_other_keys(self, tmp_config_dir):
        config.save_settings({"db_path": "/a.sqlite", "other": 42})
        config.update_setting("db_path", "/b.sqlite")
        loaded = config.load_settings()
        assert loaded["db_path"] == "/b.sqlite"
        assert loaded["other"] == 42

    def test_update_creates_new_key(self, tmp_config_dir):
        config.update_setting("collection_key", "ABCD1234")
        assert config.load_settings() == {"collection_key": "ABCD1234"}

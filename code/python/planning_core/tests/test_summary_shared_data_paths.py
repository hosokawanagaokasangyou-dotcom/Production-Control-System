from __future__ import annotations

from pathlib import Path

from planning_core.core.summary_shared_data_paths import (
    SUMMARY_AI_DISPATCH_XLSX,
    normalize_summary_shared_data_dir,
    resolve_summary_ai_dispatch_workbook_path,
    resolve_summary_shared_data_dir,
    resolve_summary_shared_data_dir_from_override,
    summary_shared_data_sibling_path,
)


def test_resolve_shared_data_dir_defaults_to_repo_code(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.delenv("PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK", raising=False)
    monkeypatch.setenv("PM_AI_REPO_ROOT", str(tmp_path))
    assert resolve_summary_shared_data_dir() == str((tmp_path / "code").resolve())


def test_legacy_workbook_path_normalizes_to_parent(tmp_path, monkeypatch):
    legacy = tmp_path / "DATA" / SUMMARY_AI_DISPATCH_XLSX
    legacy.parent.mkdir(parents=True)
    legacy.touch()
    monkeypatch.setenv("PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK", str(legacy))
    assert resolve_summary_shared_data_dir() == str(legacy.parent.resolve())
    assert resolve_summary_ai_dispatch_workbook_path() == str(
        (legacy.parent / SUMMARY_AI_DISPATCH_XLSX).resolve()
    )


def test_directory_override_is_used_directly(tmp_path, monkeypatch):
    data_dir = tmp_path / "shared-data"
    data_dir.mkdir()
    monkeypatch.setenv("PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK", str(data_dir))
    assert resolve_summary_shared_data_dir() == str(data_dir.resolve())
    assert summary_shared_data_sibling_path("attendance-data.json") == str(
        (data_dir / "attendance-data.json").resolve()
    )


def test_normalize_from_override_string(tmp_path):
    legacy = tmp_path / "old.xlsm"
    legacy.touch()
    assert normalize_summary_shared_data_dir(legacy) == tmp_path.resolve()
    assert resolve_summary_shared_data_dir_from_override(str(legacy)) == str(tmp_path.resolve())

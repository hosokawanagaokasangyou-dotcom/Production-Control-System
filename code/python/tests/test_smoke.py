"""Lightweight layout smoke tests (no Excel / xlwings required)."""

from __future__ import annotations

from pathlib import Path


def _code_python_root() -> Path:
    return Path(__file__).resolve().parent.parent


def test_planning_core_package_dir_exists():
    root = _code_python_root()
    assert (root / "planning_core").is_dir()
    assert (root / "planning_core" / "__init__.py").is_file()


def test_requirements_file_exists():
    assert (_code_python_root() / "requirements.txt").is_file()

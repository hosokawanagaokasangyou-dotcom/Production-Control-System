# -*- coding: utf-8 -*-
"""Suppress openpyxl.reader.drawings DrawingML incomplete-support UserWarning."""
from __future__ import annotations

import warnings

_installed = False


def suppress_openpyxl_drawingsml_userwarning() -> None:
    """Register warnings filter once (safe to call multiple times)."""
    global _installed
    if _installed:
        return
    warnings.filterwarnings(
        "ignore",
        message=r"DrawingML support is incomplete.*",
        category=UserWarning,
        module=r"openpyxl\.reader\.drawings",
    )
    _installed = True

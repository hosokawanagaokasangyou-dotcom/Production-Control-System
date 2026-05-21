# -*- coding: utf-8 -*-
"""Suppress openpyxl.reader.drawings UserWarning (DrawingML / unsupported image formats)."""
from __future__ import annotations

import warnings

_installed = False


def suppress_openpyxl_drawingsml_userwarning() -> None:
    """Register warnings filter once (safe to call multiple times)."""
    global _installed
    if _installed:
        return
    _drawings_module = r"openpyxl\.reader\.drawings"
    warnings.filterwarnings(
        "ignore",
        message=r"DrawingML support is incomplete.*",
        category=UserWarning,
        module=_drawings_module,
    )
    warnings.filterwarnings(
        "ignore",
        message=r".*wmf image format is not supported.*",
        category=UserWarning,
        module=_drawings_module,
    )
    _installed = True

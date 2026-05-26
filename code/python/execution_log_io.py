# -*- coding: utf-8 -*-
"""log/execution_log.txt のサイズ制限（planning_core import 前も利用可）。"""

from __future__ import annotations

import os

EXECUTION_LOG_MAX_BYTES = 10 * 1024 * 1024  # 10 MiB
_EXECUTION_LOG_TRIM_TARGET_BYTES = int(EXECUTION_LOG_MAX_BYTES * 0.85)
_UTF8_SIG_BOM = b"\xef\xbb\xbf"


def trim_execution_log_if_oversized(
    path: str, *, max_bytes: int = EXECUTION_LOG_MAX_BYTES
) -> bool:
    """上限超過時は先頭行を捨て末尾のみ残す。UTF-8 BOM 付きで書き戻す。トリムしたら True。"""
    try:
        size = os.path.getsize(path)
    except OSError:
        return False
    if size <= max_bytes:
        return False

    keep = min(_EXECUTION_LOG_TRIM_TARGET_BYTES, max_bytes)
    try:
        with open(path, "rb") as f:
            f.seek(max(0, size - keep))
            tail = f.read()
    except OSError:
        return False

    nl = tail.find(b"\n")
    if nl >= 0:
        tail = tail[nl + 1 :]

    out = tail if tail.startswith(_UTF8_SIG_BOM) else _UTF8_SIG_BOM + tail
    try:
        with open(path, "wb") as f:
            f.write(out)
            f.flush()
            try:
                os.fsync(f.fileno())
            except OSError:
                pass
    except OSError:
        return False
    return True

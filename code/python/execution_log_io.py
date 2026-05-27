# -*- coding: utf-8 -*-
"""log/execution_log.txt の行数制限（planning_core import 前も利用可）。"""

from __future__ import annotations

import os
from collections import deque

EXECUTION_LOG_MAX_LINES = 2000
_UTF8_SIG_BOM = b"\xef\xbb\xbf"


def trim_execution_log_if_oversized(
    path: str, *, max_lines: int = EXECUTION_LOG_MAX_LINES
) -> bool:
    """行数が上限超過のとき末尾 max_lines 行だけ残す。UTF-8 BOM 付きで書き戻す。トリムしたら True。"""
    if max_lines < 1:
        return False
    try:
        with open(path, "rb") as f:
            line_count = 0
            tail: deque[bytes] = deque(maxlen=max_lines)
            for line in f:
                line_count += 1
                tail.append(line)
    except OSError:
        return False

    if line_count <= max_lines:
        return False

    body = b"".join(tail)
    if body.startswith(_UTF8_SIG_BOM):
        out = body
    else:
        out = _UTF8_SIG_BOM + body
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

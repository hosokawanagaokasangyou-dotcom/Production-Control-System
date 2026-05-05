# -*- coding: utf-8 -*-

"""
スプラッシュ UserForm への xlwings 経由ログ連携は廃止した。
ログは planning_core.bootstrap の execution_log.txt（および StreamHandler）を参照すること。
環境変数 PM_AI_SPLASH_XLWINGS は互換のため残すが、常に無効として扱う。
"""

from __future__ import annotations


def enabled() -> bool:
    return False


def append_formatted_line(text: str) -> None:
    return


def shutdown() -> None:
    return

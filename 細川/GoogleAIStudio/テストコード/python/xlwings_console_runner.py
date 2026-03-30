# -*- coding: utf-8 -*-
"""
xlwings の「Show Console」＋ RunPython で段階1/2を動かす（cmd.exe 不要）。

前提
----
- Excel に **xlwings アドイン**を入れ、リボンの **Show Console** にチェック。
- マクロブック（.xlsm）と同じフォルダに **xlwings.conf.json** を置き、PYTHONPATH に
  本ファイルがある ``python`` フォルダを含める（リポジトリ同梱の例を参照）。
- VBA の標準モジュールで **xlwings への参照**を有効にし、次のように呼ぶ::

    RunPython "import xlwings_console_runner as _xw; _xw.run_stage1_for_xlwings()"
    RunPython "import xlwings_console_runner as _xw; _xw.run_stage2_for_xlwings()"

logging は planning_core が ``StreamHandler(sys.stdout)`` を付けているため、
print と同様に xlwings のコンソールへ流れる。

注意
----
- ``planning_core`` は import 時に TASK_INPUT_WORKBOOK 依存の初期化があるため、
  本モジュールでは **都度 ``planning_core`` を sys.modules から外して再 import** する。
- 既存の ``段階1_コア実行``（LOG シート埋め・plan 取り込み等）とは別経路。
  試験なら上記 RunPython のみ、本番運用は従来の cmd 経路か、VBA 側で runner 呼び出し後に
  後処理を続けるよう組み合わせてください。
"""
from __future__ import annotations

import logging
import os
import sys


def _purge_planning_core_modules() -> None:
    for k in list(sys.modules.keys()):
        if k == "planning_core" or k.startswith("planning_core."):
            del sys.modules[k]


def _prepare_from_caller_book() -> str:
    import xlwings as xw

    wb = xw.Book.caller()
    path = os.path.abspath(str(wb.fullname))
    root = os.path.dirname(path)
    os.chdir(root)
    os.environ["TASK_INPUT_WORKBOOK"] = path
    return path


def run_stage1_for_xlwings() -> int:
    """
    段階1（run_stage1_extract）。戻り値: 0=成功, 1=失敗, 2=caller 不備。
    """
    try:
        _prepare_from_caller_book()
    except Exception:
        logging.exception(
            "xlwings: Book.caller() を取得できません。"
            " マクロブック上のボタンから RunPython してください。"
        )
        return 2
    _purge_planning_core_modules()
    try:
        import planning_core as pc

        ok = pc.run_stage1_extract()
        return 0 if ok else 1
    except SystemExit as e:
        c = e.code
        if c is None:
            return 0
        if isinstance(c, int):
            return 0 if c == 0 else c
        return 1
    except Exception:
        logging.exception("xlwings: 段階1で未捕捉例外")
        return 1


def run_stage2_for_xlwings() -> int:
    """
    段階2（generate_plan）。戻り値: 0=正常終了, 1=例外, 2=caller 不備。
    """
    try:
        _prepare_from_caller_book()
    except Exception:
        logging.exception("xlwings: Book.caller() を取得できません。")
        return 2
    _purge_planning_core_modules()
    try:
        import planning_core as pc

        pc.generate_plan()
        return 0
    except SystemExit as e:
        c = e.code
        if c is None:
            return 0
        if isinstance(c, int):
            return 0 if c == 0 else c
        return 1
    except Exception:
        logging.exception("xlwings: 段階2で未捕捉例外")
        return 1

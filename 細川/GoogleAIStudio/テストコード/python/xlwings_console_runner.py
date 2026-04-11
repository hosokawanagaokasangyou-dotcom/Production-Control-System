# -*- coding: utf-8 -*-
"""
xlwings の RunPython から段階1／段階2 等を呼ぶときのエントリ（cmd.exe 起動版の代替・補助）。

概要
----
- Excel 側 **xlwings アドイン** の **Show Console** 有無とは別に、マクロブックと **同じフォルダ** の
  ``python/`` 配下に本ファイルを置き、VBA から ``runpy.run_path`` で読み込む想定。
  例: ``.../マクロブック.xlsm`` と同階層の ``python/xlwings_console_runner.py``
- ``import xlwings_console_runner`` だけでは ``sys.path`` に ``python`` が入らない場合があるため、
  **VBA から ``runpy.run_path`` で本ファイルを実行する**か、``RunPython`` 用に
  ``xlwings.conf.json`` の PYTHONPATH 等で ``python`` を通す。
- VBA から xlwings RunPython を ON にする例::

    xlwings.RunPython "import os, runpy, xlwings as xw; wb=xw.Book.caller(); p=os.path.join(os.path.dirname(str(wb.fullname)), 'python', 'xlwings_console_runner.py'); ns=runpy.run_path(p); ns['run_stage1_for_xlwings']()"

  段階1／段階2 の本体は ``task_extract_stage1.py`` / ``plan_simulation_stage2.py`` 等と
  ``STAGE12_USE_XLWINGS_RUNPYTHON`` の組み合わせで切り替え。
  終了コードは ``log/stage_vba_exitcode.txt`` に 1 行書き、VBA や cmd 側が参照可能。

logging は planning_core 側の ``StreamHandler(sys.stdout)`` 等に依存。
print や logging は xlwings コンソールに出る場合がある。

注意
----
- マクロブックの ``設定_環境変数`` 等は ``_apply_workbook_env_overrides`` 経由で
  ``import planning_core`` より前に ``os.environ`` に載る（``workbook_env_bootstrap``）。
- ``planning_core`` を import する前に TASK_INPUT_WORKBOOK をセットする。
  本モジュールでは **都度 ``planning_core`` を ``sys.modules`` から外してから import** する。
- cmd 経由の ``task_extract_stage1.py`` は planning_core  import 後に ``execution_log.txt`` へ出力するが、
  **本ランナー**の ``run_stage1_for_xlwings`` / ``run_stage2_for_xlwings`` は import の都度クリーンに近づけるため
  モジュールを消してから読み直す。VBA のログ枠との同期挙動は cmd 版と異なる場合がある。
- ``runpy.run_path`` は **sys.path の先頭に本ファイルのある ``python`` を足す**（``_ensure_this_python_dir_on_syspath``）。
  ``import planning_core`` で ``ModuleNotFoundError`` になる場合はここを確認。
"""
from __future__ import annotations

import logging
import os
import sys
import traceback
from datetime import datetime

STAGE_VBA_EXIT_CODE_FILE = "stage_vba_exitcode.txt"


def _ensure_this_python_dir_on_syspath() -> None:
    """planning_core を解決できるよう、本ファイルと同じ ``python`` を sys.path に追加（run_path 用）。"""
    here = os.path.dirname(os.path.abspath(__file__))
    here_n = os.path.normcase(here)
    for _p in sys.path:
        try:
            if os.path.normcase(os.path.abspath(_p)) == here_n:
                return
        except (OSError, ValueError):
            continue
    sys.path.insert(0, here)


_ensure_this_python_dir_on_syspath()


def _apply_workbook_env_overrides() -> None:
    """マクロブックの環境変数シート等を反映し、TASK_INPUT_WORKBOOK を含め planning_core 読込前の状態にする。"""
    try:
        import workbook_env_bootstrap as _wbe

        _wbe.apply_from_task_input_workbook()
    except Exception:
        pass


def _write_stage_vba_exit_code(code: int) -> None:
    """VBA の ReadStageVbaExitCodeFromFile 等が読む 1 行の終了コードを書く。"""
    try:
        logd = os.path.join(os.getcwd(), "log")
        os.makedirs(logd, exist_ok=True)
        p = os.path.join(logd, STAGE_VBA_EXIT_CODE_FILE)
        with open(p, "w", encoding="utf-8", newline="\n") as f:
            f.write(str(int(code)))
            f.flush()
            os.fsync(f.fileno())
    except OSError:
        pass


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


def _execution_log_path() -> str:
    return os.path.join(os.getcwd(), "log", "execution_log.txt")


def _append_execution_log_line(level: str, msg: str) -> None:
    """
    cmd 版の task_extract_stage1 と同様、planning_core の log に加え、VBA のログ枠へ送る行を追記する。
    """
    path = _execution_log_path()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts} - {level} - {msg}\n"
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "a", encoding="utf-8-sig", newline="\n") as f:
            f.write(line)
            f.flush()
            try:
                os.fsync(f.fileno())
            except OSError:
                pass
    except OSError:
        return
    try:
        import xlwings_splash_log as xsl

        if xsl.enabled():
            xsl.append_formatted_line(line)
    except Exception:
        pass


def _append_execution_log_traceback(title: str) -> None:
    """except ブロックから traceback.format_exc を execution_log に追記する。"""
    _append_execution_log_line("ERROR", title)
    tb = traceback.format_exc()
    try:
        path = _execution_log_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "a", encoding="utf-8-sig", newline="\n") as f:
            f.write(tb)
            if not tb.endswith("\n"):
                f.write("\n")
            f.flush()
            try:
                os.fsync(f.fileno())
            except OSError:
                pass
    except OSError:
        pass


def run_stage1_for_xlwings() -> int:
    """
    段階1: ``run_stage1_extract`` を実行。戻り値: 0=成功, 1=失敗, 2=caller 取得失敗。
    終了コードは ``log/stage_vba_exitcode.txt`` にも書く。
    """
    rc = 1
    try:
        try:
            _prepare_from_caller_book()
        except Exception:
            logging.exception(
                "xlwings: Book.caller() の取得に失敗しました。"
                " ブックが呼び出し元として開かれているか、RunPython の指定を確認してください。"
            )
            rc = 2
            return rc
        _apply_workbook_env_overrides()
        _append_execution_log_line(
            "INFO",
            "段階1: xlwings run_stage1_for_xlwings から planning_core を実行します。",
        )
        _purge_planning_core_modules()
        try:
            import planning_core as pc

            ok = pc.run_stage1_extract()
            rc = 0 if ok else 1
        except SystemExit as e:
            c = e.code
            if c is None:
                rc = 0
            elif isinstance(c, int):
                rc = 0 if c == 0 else c
            else:
                rc = 1
        except Exception:
            logging.exception("xlwings: 段階1の実行で例外が発生しました。")
            _append_execution_log_traceback("xlwings: 段階1の実行で例外が発生しました。")
            rc = 1
        return rc
    finally:
        _write_stage_vba_exit_code(rc)


def run_refresh_plan_input_dispatch_trial_order_for_xlwings() -> int:
    """
    配台計画_タスク入力の配台試行順番を再計算（refresh_plan_input_dispatch_trial_order）。
    VBA: XwRunConsoleRunner "run_refresh_plan_input_dispatch_trial_order_for_xlwings"
    """
    rc = 1
    try:
        try:
            _prepare_from_caller_book()
        except Exception:
            logging.exception("xlwings: Book.caller() の取得に失敗しました。")
            rc = 2
            return rc
        _apply_workbook_env_overrides()
        _append_execution_log_line(
            "INFO",
            "配台試行順: xlwings run_refresh_plan_input_dispatch_trial_order_for_xlwings 開始",
        )
        _purge_planning_core_modules()
        try:
            import planning_core as pc

            ok = pc.refresh_plan_input_dispatch_trial_order_only()
            rc = 0 if ok else 1
        except SystemExit as e:
            c = e.code
            if c is None:
                rc = 0
            elif isinstance(c, int):
                rc = 0 if c == 0 else c
            else:
                rc = 1
        except Exception:
            logging.exception("xlwings: 配台試行順の更新で失敗しました。")
            _append_execution_log_traceback("xlwings: 配台試行順の更新で失敗しました。")
            rc = 1
        return rc
    finally:
        _write_stage_vba_exit_code(rc)


def run_sort_plan_input_dispatch_trial_order_by_float_keys_for_xlwings() -> int:
    """
    配台計画_タスク入力: 配台試行順番を小数キーで並べ替え 1..n（xlwings 用）。
    VBA: XwRunConsoleRunner "run_sort_plan_input_dispatch_trial_order_by_float_keys_for_xlwings"
    """
    rc = 1
    try:
        try:
            _prepare_from_caller_book()
        except Exception:
            logging.exception("xlwings: Book.caller() の取得に失敗しました。")
            rc = 2
            return rc
        _apply_workbook_env_overrides()
        _append_execution_log_line(
            "INFO",
            "配台試行順: xlwings run_sort_plan_input_dispatch_trial_order_by_float_keys_for_xlwings 開始",
        )
        _purge_planning_core_modules()
        try:
            import planning_core as pc

            ok = pc.sort_plan_input_dispatch_trial_order_by_float_keys_only()
            rc = 0 if ok else 1
        except SystemExit as e:
            c = e.code
            if c is None:
                rc = 0
            elif isinstance(c, int):
                rc = 0 if c == 0 else c
            else:
                rc = 1
        except Exception:
            logging.exception("xlwings: 配台試行順（小数キー並べ）で失敗しました。")
            _append_execution_log_traceback("xlwings: 配台試行順（小数キー並べ）で失敗しました。")
            rc = 1
        return rc
    finally:
        _write_stage_vba_exit_code(rc)


def run_stage2_for_xlwings() -> int:
    """
    段階2: ``generate_plan`` を実行。戻り値: 0=成功, 1=失敗, 2=caller 取得失敗。
    終了コードは ``log/stage_vba_exitcode.txt`` にも書く。
    """
    rc = 1
    try:
        try:
            _prepare_from_caller_book()
        except Exception:
            logging.exception("xlwings: Book.caller() の取得に失敗しました。")
            rc = 2
            return rc
        _apply_workbook_env_overrides()
        _append_execution_log_line(
            "INFO",
            "段階2: xlwings run_stage2_for_xlwings から planning_core を実行します。",
        )
        _purge_planning_core_modules()
        try:
            import planning_core as pc

            pc.generate_plan()
            rc = 0
        except SystemExit as e:
            c = e.code
            if c is None:
                rc = 0
            elif isinstance(c, int):
                rc = 0 if c == 0 else c
            else:
                rc = 1
        except Exception:
            logging.exception("xlwings: 段階2の実行で例外が発生しました。")
            _append_execution_log_traceback("xlwings: 段階2の実行で例外が発生しました。")
            rc = 1
        return rc
    finally:
        _write_stage_vba_exit_code(rc)

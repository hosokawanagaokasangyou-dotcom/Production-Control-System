# -*- coding: utf-8 -*-
"""
planning_core パッケージの import 時ブートストラップ。

カレントディレクトリ・ルートロガー・output/log 等のディレクトリを確保する。
``import planning_core`` の最初に本モジュールが読み込まれること。
"""

import ctypes
import fnmatch
import logging
import os
import shutil
import sys
import time as time_module

# =========================================================
# 【重要】カレントをマクロブック（TASK_INPUT_WORKBOOK）と同じフォルダに設定する。
# log/execution_log.txt・output/ を VBA の ThisWorkbook.Path 基準と一致させる。
# 環境変数が無い・chdir 不可のときは従来どおり「python の親」またはパッケージの親階層へ。
# （VBA は …\output\plan_input_tasks.xlsx を参照するため output はブックと同じ階層が前提）
# =========================================================
_env_wb_for_cwd = (os.environ.get("TASK_INPUT_WORKBOOK") or "").strip()
_cwd_from_workbook = False
if _env_wb_for_cwd:
    try:
        os.chdir(os.path.dirname(os.path.abspath(_env_wb_for_cwd)))
        _cwd_from_workbook = True
    except OSError:
        _cwd_from_workbook = False
if not _cwd_from_workbook:
    _here = os.path.dirname(os.path.abspath(__file__))
    _base = os.path.basename(_here).lower()
    if _base == "planning_core":
        _python_dir = os.path.dirname(_here)
        if os.path.basename(_python_dir).lower() == "python":
            os.chdir(os.path.dirname(_python_dir))
        else:
            os.chdir(_python_dir)
    elif _base == "python":
        os.chdir(os.path.dirname(_here))
    else:
        os.chdir(_here)

# cmd で chcp 65001 時にログの日本語をコンソールへ出しやすくする（段階2でリダイレクト無し実行時向け）
if os.name == "nt" and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ---------------------------------------------------------
# コマンドプロンプト画面を「最前面（トップレベル）」にする処理
# ---------------------------------------------------------
if os.name == "nt":
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        # HWND_TOPMOST = -1, SWP_NOMOVE = 0x0002, SWP_NOSIZE = 0x0001
        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 3)

# ---------------------------------------------------------
# ログを「画面（コンソール）」と「一時ファイル」の両方に出力する
# ---------------------------------------------------------
logger = logging.getLogger()
logger.setLevel(logging.INFO)
# 既存のハンドラをクリア
if logger.hasHandlers():
    logger.handlers.clear()

# _core.py のコメント・文字列は UTF-8 正字に修正済み。旧 execution_log 向けの誤 Unicode 置換は撤廃した。
_LOG_MOJIBAKE_PAIRS: tuple[tuple[str, str], ...] = ()
_SINGLE_MOJIBAKE: tuple[tuple[str, str], ...] = ()


def _normalize_log_line(text: str) -> str:
    """execution_log / コンソール向け（互換のため残置。現状は変換なし）。"""
    return text


class _MojibakeFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        return _normalize_log_line(super().format(record))


formatter = _MojibakeFormatter("%(asctime)s - %(levelname)s - %(message)s")


class _FlushingFileHandler(logging.Handler):
    """1レコードごとに追記後 fsync し、都度ファイルを閉じる。

    logging.FileHandler はファイルを開きっぱなしにするため、Windows では VBA 側の
    ADODB.LoadFromFile / FileCopy が共有違反で失敗し、スプラッシュのポーリングが
    先頭1行のまま止まることがある。開閉 per emit で他プロセスが読めるようにする。
    """

    terminator = "\n"

    def __init__(self, filename: str, encoding: str = "utf-8-sig") -> None:
        super().__init__()
        self.baseFilename = os.path.abspath(filename)
        self.encoding = encoding

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record) + self.terminator
            with open(self.baseFilename, "a", encoding=self.encoding, newline="\n") as f:
                f.write(msg)
                f.flush()
                try:
                    os.fsync(f.fileno())
                except OSError:
                    pass
        except Exception:
            self.handleError(record)


# 1. 画面(コンソール)用ハンドラ
stream_handler = logging.StreamHandler(sys.stdout)
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)

# 2. 成果物用 output / 実行ログ用 log（いずれも常に最新1ファイルのみ上書き）
output_dir = os.path.join(os.getcwd(), "output")
os.makedirs(output_dir, exist_ok=True)
log_dir = os.path.join(os.getcwd(), "log")
os.makedirs(log_dir, exist_ok=True)
# 段階2がユーザー検証で中断したときの1行メッセージ（VBA が log から読み MsgBox 用）
STAGE2_BLOCKING_MESSAGE_FILE = "stage2_blocking_message.txt"
# Gemini 利用・推定料金の累計 JSON（log ではなく専用フォルダ）
api_payment_dir = os.path.join(os.getcwd(), "API_Payment")
os.makedirs(api_payment_dir, exist_ok=True)
# 上書き保存するアプリ用 JSON（API 料金累計は API_Payment のみ）
json_data_dir = os.path.join(os.getcwd(), "json")
os.makedirs(json_data_dir, exist_ok=True)
# 旧仕様: ログが output 直下にあった名残を削除（常に log 側の1ファイルのみが最新）
for _legacy_name in (
    "execution_log.txt",
    "ai_task_special_remark_last.txt",
    "ai_task_special_last_prompt.txt",
    "planning_conflict_highlight.tsv",
    "cmd_stage2.log",
):
    try:
        os.remove(os.path.join(output_dir, _legacy_name))
    except OSError:
        pass


def _try_remove_path_with_retries(
    path: str, *, attempts: int = 5, delay_s: float = 0.35
) -> tuple[bool, OSError | None]:
    """Excel 占有中のファイルが一瞬解放されることがあるため短いリトライを行う。"""
    last_ex: OSError | None = None
    for _ in range(max(1, attempts)):
        try:
            os.remove(path)
            return True, None
        except OSError as ex:
            last_ex = ex
            time_module.sleep(delay_s)
    return False, last_ex


def _remove_prior_stage2_workbooks_and_prune_empty_dirs(output_root: str) -> None:
    """
    ``production_plan_multi_day_*.xlsx`` / ``member_schedule_*.xlsx`` を output 配下からすべて削除し、
    空になったサブフォルダを削除する（日付階層の旧出力を含む）。
    段階2の直前に呼び、常に最新1組の成果物だけを残す土台にする。
    """
    if not output_root or not os.path.isdir(output_root):
        return
    patterns = ("production_plan_multi_day_*.xlsx", "member_schedule_*.xlsx")
    root_abs = os.path.normcase(os.path.abspath(output_root))
    removed = 0
    failed_paths: list[str] = []
    for dirpath, _dirnames, filenames in os.walk(output_root, topdown=False):
        for name in filenames:
            if name.startswith("~$"):
                continue
            for pat in patterns:
                if fnmatch.fnmatch(name, pat):
                    fp = os.path.join(dirpath, name)
                    ok, last_ex = _try_remove_path_with_retries(fp)
                    if ok:
                        removed += 1
                    else:
                        failed_paths.append(fp)
                        logging.warning(
                            "段階2旧出力の削除に失敗: %s (%s)",
                            fp,
                            last_ex,
                        )
                    break
        dir_abs = os.path.normcase(os.path.abspath(dirpath))
        if dir_abs == root_abs:
            continue
        try:
            if not os.listdir(dirpath):
                os.rmdir(dirpath)
        except OSError:
            pass
    if removed:
        logging.info(
            "段階2出力の整理: production_plan_multi_day_*.xlsx / member_schedule_*.xlsx を %s 件削除しました。",
            removed,
        )
    if failed_paths:
        preview = "; ".join(failed_paths[:3])
        if len(failed_paths) > 3:
            preview += " …"
        logging.error(
            "段階2: 旧出力 %s 件を削除できませんでした（多くの場合、Excel で当該 .xlsx を開いたままです）。"
            " ファイルを閉じてから再実行してください。対象例: %s",
            len(failed_paths),
            preview,
        )


# 3. ファイル用ハンドラ（VBAで後から読み取るため UTF-8 で保存）
log_file_path = os.path.join(log_dir, "execution_log.txt")
# BOM 付き UTF-8（Excel / VBA の ADODB.Stream が文字化けしにくい）
file_handler = _FlushingFileHandler(log_file_path, encoding="utf-8-sig")
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)


def _maybe_register_xlwings_splash_logging() -> None:
    """PM_AI_SPLASH_XLWINGS=1 のとき xlwings で UserForm txtExecutionLog にログ行を送る。"""
    try:
        import atexit

        import xlwings_splash_log as xsl

        if not xsl.enabled():
            return

        class _XlwingsSplashLogHandler(logging.Handler):
            terminator = "\n"

            def emit(self, record: logging.LogRecord) -> None:
                try:
                    import xlwings_splash_log as xsl2

                    msg = self.format(record)
                    if self.terminator and not msg.endswith(self.terminator):
                        msg += self.terminator
                    xsl2.append_formatted_line(msg)
                except Exception:
                    self.handleError(record)

        _xh = _XlwingsSplashLogHandler()
        _xh.setFormatter(formatter)
        _xh.setLevel(logging.INFO)
        logger.addHandler(_xh)

        def _atexit_flush_splash() -> None:
            try:
                import xlwings_splash_log as xsl3

                xsl3.shutdown()
            except Exception:
                pass

        atexit.register(_atexit_flush_splash)
    except Exception:
        pass


_maybe_register_xlwings_splash_logging()

stage2_blocking_message_path = os.path.join(log_dir, STAGE2_BLOCKING_MESSAGE_FILE)


class PlanningValidationError(Exception):
    """配台計画シートの検証エラー（段階2を続行しない）。メッセージは log にも書く。"""


def _clear_stage2_blocking_message_file() -> None:
    try:
        if os.path.isfile(stage2_blocking_message_path):
            os.remove(stage2_blocking_message_path)
    except OSError:
        pass


def _write_stage2_blocking_message(message: str) -> None:
    s = (message or "").strip()
    if not s:
        return
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(stage2_blocking_message_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(s)
    except OSError as ex:
        logging.warning("stage2_blocking_message の書き込みに失敗: %s", ex)

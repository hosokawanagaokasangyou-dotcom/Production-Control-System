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

# _core.py 由来の誤 Unicode（U+374x 系・〝・酝坰 等）をログ表示用に正す（UTF-8 保存は正しいが文言が破損している場合）
_LOG_MOJIBAKE_PAIRS: tuple[tuple[str, str], ...] = tuple(
    sorted(
        (
            ("酝坰_酝坰丝覝工程", "配台_配台不要工程"),
            ("設定_酝坰丝覝工程", "設定_配台試行工程"),
            ("酝坰丝覝工程", "配台試行工程"),
            ("酝坰トレース", "配台トレース"),
            ("酝坰計画", "配台計画"),
            ("酝坰試行順番", "配台試行順番"),
            ("酝坰丝覝ルール", "配台試行ルール"),
            ("酝坰丝覝", "配台試行"),
            ("酝坰丝能", "配台試行"),
            ("酝坰丝坂加", "配台試行加"),
            ("酝坰時追加人数", "配台時追加人数"),
            ("酝坰対象", "配台対象"),
            ("デポッグ酝坰", "デバッグ配台"),
            ("全件酝坰", "全件配台"),
            ("酝完_", "配台済_"),
            ("酝坰", "配台"),
            ("暗坷化誝証", "暗号化証明書"),
            ("誝証", "証明書"),
            ("暗坷化", "暗号化"),
            ("読㝿込㝿㝾㝗㝟", "読み込みました"),
            ("読㝿込㝿", "読み込み"),
            ("見㝤㝋り㝾㝛ん", "見つかりません"),
            ("㝌㝂り㝾㝛ん", "はありません"),
            ("読ゝ㝾㝛ん", "読めません"),
            ("読ゝ㝪㝄㝟ゝ", "読めないため"),
            ("使ゝ㝪㝄", "使えない"),
            ("使ゝ㝚〝", "使うので、"),
            ("書㝄㝦㝝㝠㝕㝄", "書いてください"),
            ("書㝝", "書き"),
            ("作り直㝙", "作り直し"),
            ("作戝", "作成"),
            ("生戝", "生成"),
            ("戝坦", "追跡"),
            ("絝果_", "結果_"),
            ("絝果", "結果"),
            ("グローポル", "グローバル"),
            ("メンポー", "メンバー"),
            ("稼僝", "稼働"),
            ("管睆", "管理"),
            ("用愝", "用意"),
            ("坂照", "参照"),
            ("一括坝映", "一括反映"),
            ("坝映", "反映"),
            ("坠有", "占有"),
            ("ポェック", "チェック"),
            ("ポイライト", "ハイライト"),
            ("解枝", "解析"),
            ("以陝", "以降"),
            ("上陝", "上限"),
            ("下陝", "下限"),
            ("必覝", "必須"),
            ("確誝", "確認"),
            ("冝実行", "再実行"),
            ("冝記載", "再記載"),
            ("冝", "再"),
            ("革用", "適用"),
            ("欠蝽", "欠損"),
            ("欠杝", "欠損"),
            ("杛算", "残作"),
            ("数針", "数値"),
            ("実出来高杛算", "実出来高残作"),
            ("ポーム", "フォーム"),
            ("酝置", "配置"),
            ("坯能", "可能"),
            ("投入坯能日", "投入可能日"),
            ("回答紝期", "回答納期"),
            ("指定紝期", "指定納期"),
            ("計画基準紝期", "計画基準納期"),
            ("原坝投入日", "原板投入日"),
            ("靅れ", "遅れ"),
            ("处睆", "処理"),
            ("処睆", "処理"),
            ("ロジック弝", "ロジック式"),
            ("〝設定】", "【設定】"),
            ("既定 9〝2", "既定 9、2"),
            ("確定後〝下記", "確定後、下記"),
            ("退靿㝗〝次回", "退避し、次回"),
            ("呼出㝗〝S", "呼出し、S"),
            ("MessageBox〝㝝れ以外", "MessageBox。それ以外"),
            ("㝟㝠㝗", "として"),
            ("㝯〝備考", "は「備考"),
            ("㝟ゝ㝮", "ための"),
            ("㝋ら", "から"),
            ("㝨㝗㝦", "として"),
            ("㝾㝙", "した"),
            ("㝾㝟㝯", "または"),
            ("㝾㝗㝟", "ました"),
            ("㝧㝙", "です"),
            ("指定㝕れ㝟", "指定された"),
            ("指定㝕れ", "指定され"),
            ("スキップ㝕れした", "スキップされました"),
            ("㝕れした", "されました"),
            ("㝕れ", "され"),
            ("無㝑れ㝰", "無ければ"),
            ("打㝡切り", "打ち切り"),
            ("列坝", "列名"),
            ("シート坝", "シート名"),
            ("ファイル坝", "ファイル名"),
            ("工程坝", "工程名"),
            ("機械坝", "機械名"),
            ("依頼坝", "依頼名"),
            ("用愝㝗", "用意し"),
            ("管睆者", "管理者"),
            ("従㝄", "従い"),
            ("失敗㝙る", "失敗する"),
            ("指定㝙る", "指定する"),
            ("試行㝗㝪㝄", "試行する"),
            ("試行㝗㝾㝙", "試行した"),
            ("入力㝗㝟", "入力した"),
            ("利用㝧㝝㝾㝛ん", "利用できません"),
            ("利用㝙る", "利用する"),
            ("作戝・保存㝮", "作成・保存の"),
            ("㝨併用", "と併用"),
            ("デポッグ", "デバッグ"),
            ("平文㝾㝟㝯", "平文または"),
            ("㝲㝪型", "参考型"),
            ("㝙㝹㝦", "まとめて"),
            ("坫む", "含む"),
            ("坫㝾れる", "含まれる"),
            ("坖る", "得る"),
            ("啝", "相"),
            ("坘佝", "短縮"),
            ("靿㝑る", "避ける"),
            ("靿", "避"),
            ("続㝑㝦", "続けて"),
            ("割坈", "割引"),
            ("外坴", "外側"),
            ("丝覝", "試行"),
            ("丝能", "試行"),
        ),
        key=lambda x: -len(x[0]),
    )
)

_SINGLE_MOJIBAKE = (
    ("㝕", "さ"),
    ("〝", "」"),
    ("㝮", "の"),
    ("㝯", "は"),
    ("㝫", "に"),
    ("㝧", "で"),
    ("㝨", "と"),
    ("㝌", "は"),
    ("㝓", "こ"),
    ("㝋", "か"),
    ("㝿", "み"),
    ("㝾", "ま"),
    ("㝛", "せ"),
    ("㝙", "れ"),
    ("㝟", "た"),
    ("㝝", "し"),
    ("㝪", "な"),
    ("㝄", "い"),
    ("㝂", "あ"),
    ("㝤", "つ"),
    ("㝸", "へ"),
    ("㝲", "参"),
    ("㝠", "て"),
)


def _normalize_log_line(text: str) -> str:
    """execution_log / コンソール向けに誤字置換（ソース定数は変更しない）。"""
    for old, new in _LOG_MOJIBAKE_PAIRS:
        text = text.replace(old, new)
    for old, new in _SINGLE_MOJIBAKE:
        text = text.replace(old, new)
    text = text.replace("# 」設定】", "# 【設定】")
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

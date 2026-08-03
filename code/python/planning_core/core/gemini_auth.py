# -*- coding: utf-8 -*-
# planning_core.core.gemini_auth — body only (loaded via _core exec chain)
GEMINI_MODEL_IDS_BY_QUALITY: tuple[str, ...] = (
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-3.1-flash-lite-preview",
    "gemini-flash-latest",
)
GEMINI_MODEL_FLASH = "gemini-3.5-flash"
# 無料枠の割り当ては世代ごとに打ち切られる。gemini-1.x / 2.x と pro 系は未使用のキーでも
# 429 を返すため、待機付き再試行を誘発するだけの候補として試行列から除外する。
GEMINI_EXHAUSTED_FREE_TIER_PREFIXES: tuple[str, ...] = ("gemini-1.", "gemini-2.")
def _gemini_model_has_free_tier_allocation(model_id: str) -> bool:
    """モデル ID が無料枠の割り当てを持つ世代・系統か（``models/`` 接頭辞は無視）。"""
    mid = str(model_id or "").strip().lower()
    if mid.startswith("models/"):
        mid = mid[len("models/") :]
    if not mid:
        return False
    if "pro" in mid:
        return False
    return not mid.startswith(GEMINI_EXHAUSTED_FREE_TIER_PREFIXES)
_GEMINI_FLASH_IN_PER_M = float(
    os.environ.get("GEMINI_PRICE_USD_IN_PER_M", "0.075") or 0.075
)
_GEMINI_FLASH_OUT_PER_M = float(
    os.environ.get("GEMINI_PRICE_USD_OUT_PER_M", "0.30") or 0.30
)
GEMINI_JPY_PER_USD = float(os.environ.get("GEMINI_JPY_PER_USD", "150") or 150)
_GEMINI_PRE_REQUEST_JITTER_MAX = float(
    os.environ.get("GEMINI_PRE_REQUEST_JITTER_MAX_SEC", "0.75") or 0.75
)
_GEMINI_RETRY_BACKOFF_BASE = float(
    os.environ.get("GEMINI_RETRY_BACKOFF_BASE_SEC", "2.0") or 2.0
)
_GEMINI_RETRY_MAX_ATTEMPTS = max(
    1, int(os.environ.get("GEMINI_RETRY_MAX_ATTEMPTS", "3") or 3)
)
def _gemini_request_timeout_sec() -> float:
    raw = (os.environ.get("GEMINI_REQUEST_TIMEOUT_SEC") or "").strip()
    if not raw:
        return 60.0
    try:
        v = float(raw)
    except (TypeError, ValueError):
        return 60.0
    return max(0.0, v)
def _gemini_client(api_key: str) -> genai.Client:
    """API キー付き Client。可能なら HttpOptions で読み取りタイムアウトを付与する。"""
    sec = _gemini_request_timeout_sec()
    if sec > 0:
        try:
            from google.genai import types as genai_types

            ms = max(1000, int(round(sec * 1000.0)))
            return genai.Client(
                api_key=api_key,
                http_options=genai_types.HttpOptions(timeout=ms),
            )
        except Exception:
            logging.debug(
                "Gemini Client: HttpOptions によるタイムアウト設定に失敗したため、既定クライアントを使用します。",
                exc_info=True,
            )
    return genai.Client(api_key=api_key)
def master_workbook_filename() -> str:
    """マスタブックの basename。環境変数 ``PM_AI_MASTER_WORKBOOK`` から取得（空なら ``master.xlsm``）。"""
    alt = (os.environ.get("PM_AI_MASTER_WORKBOOK") or "").strip()
    if alt:
        return os.path.basename(alt)
    return "master.xlsm"
MASTER_FILE = master_workbook_filename()
def _master_workbook_path_resolved() -> str:
    """
    マスタブックの絶対パス。環境変数 ``PM_AI_MASTER_WORKBOOK`` のみ（必須・実在ファイル）。

    ``PM_AI_MASTER_WORKBOOK`` 未設定・未存在時は ``PlanningValidationError``。
    """
    alt = (os.environ.get("PM_AI_MASTER_WORKBOOK") or "").strip()
    if not alt:
        raise PlanningValidationError(
            "環境変数 PM_AI_MASTER_WORKBOOK が未設定です。"
            "JavaFX の環境変数タブでマスタブック（.xlsm）の絶対パスを設定してください。"
        )
    if not os.path.isfile(alt):
        raise PlanningValidationError(
            "PM_AI_MASTER_WORKBOOK で指定したマスタブックが見つかりません: "
            f"{alt!r}"
        )
    return os.path.normpath(os.path.abspath(alt))
def _require_master_workbook_path_exists() -> str:
    """``_master_workbook_path_resolved`` と同じ（未設定・未存在時は ``PlanningValidationError``）。"""
    return _master_workbook_path_resolved()
MACHINE_CALENDAR_SLOT_MINUTES = 30
_MACHINE_CALENDAR_BLOCKS_BY_DATE: dict[
    date, dict[str, list[tuple[datetime, datetime]]]
] = {}
_MACHINE_CALENDAR_INTERACTIVE_DEFINED_SLOTS_BY_DATE: dict[
    date, list[tuple[datetime, datetime]]
] = {}
_STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE: dict[str, int] = {}
_STAGE2_MACHINE_DAILY_STARTUP_REQ_BY_MACHINE: dict[str, int] = {}
_STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE: dict[tuple[str, str], int] = {}
_STAGE2_REQUEST_SWITCH_PREP_BY_MACHINE: dict[str, int] = {}
_STAGE2_BREAK_RESUME_PREP_BY_PROC_MACHINE: dict[tuple[str, str], int] = {}
_STAGE2_BREAK_RESUME_PREP_BY_MACHINE: dict[str, int] = {}
_STAGE2_POST_MACHINING_CLEANUP_BY_PROC_MACHINE: dict[tuple[str, str], int] = {}
_STAGE2_POST_MACHINING_CLEANUP_BY_MACHINE: dict[str, int] = {}
_STAGE2_REQUEST_INTERVAL_BUFFER_BY_PROC_MACHINE: dict[tuple[str, str], int] = {}
_STAGE2_REQUEST_INTERVAL_BUFFER_BY_MACHINE: dict[str, int] = {}
_STAGE2_REGULAR_SHIFT_START: time | None = None
_STAGE2_DATA_EXTRACTION_DATETIME: datetime | None = None
SURPLUS_TEAM_MAX_SPEEDUP_RATIO = 0.05
TASKS_SHEET_NAME = "加工計画DATA"
TASKS_SHEET_NAME_FOR_ACTUAL_GANTT_PLAN = "加工計画DATA_実績比較用"
OPENPYXL_INCOMPATIBLE_SHEET_MARKER = "配台_配台不要工程"
def _excel_plan_input_wb() -> str:
    """専用 UI の ``PM_AI_PLAN_INPUT_PATH`` が Excel ブック実ファイルのときの絶対パス（それ以外は空）。"""
    return plan_input_workbook_path_for_excel_ops()
def _ooxml_workbook_sheet_names(wb_path: str) -> list[str] | None:
    """ZIP 内 xl/workbook.xml からシート名一覧を得る（openpyxl を使えない）。"""
    import zipfile
    import xml.etree.ElementTree as ET

    if not wb_path or not os.path.isfile(wb_path):
        return None
    low = wb_path.lower()
    if not low.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        return None
    try:
        with zipfile.ZipFile(wb_path, "r") as zf:
            if "xl/workbook.xml" not in zf.namelist():
                return None
            raw = zf.read("xl/workbook.xml")
    except (OSError, zipfile.BadZipFile, KeyError):
        return None
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return None
    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    tag_sheet = "{%s}sheet" % ns_main
    names: list[str] = []
    for el in root.iter():
        if el.tag == tag_sheet or el.tag.endswith("}sheet"):
            n = el.get("name")
            if n:
                names.append(n)
    return names
def _ooxml_workbook_missing_shared_strings(wb_path: str) -> bool:
    """
    OOXML ブックに xl/sharedStrings.xml が無いとき True。

    openpyxl の load_workbook / read_only 読込はメモリ急増・応答停止の原因になりやすい（専用 UI 等の xlsx）。
    """
    p = (wb_path or "").strip()
    if not p or not os.path.isfile(p):
        return False
    low = p.lower()
    if not low.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        return False
    try:
        import zipfile

        with zipfile.ZipFile(p, "r") as zf:
            missing = "xl/sharedStrings.xml" not in zf.namelist()
        return missing
    except Exception:
        return False
def normalize_ooxml_shared_strings_if_missing(path: str) -> bool:
    """
    xl/sharedStrings.xml が無い xlsx を calamine 読込 + openpyxl 書込で再保存する。
    段階1出力 plan_input_tasks.xlsx 等で openpyxl 系後処理がスキップされるのを防ぐ。
    """
    p = (path or "").strip()
    if not p or not os.path.isfile(p) or not _ooxml_workbook_missing_shared_strings(p):
        return False
    try:
        xf = pd.ExcelFile(p, engine="calamine")
    except Exception as ex:
        logging.debug("sharedStrings 正規化: calamine 読込失敗 path=%r err=%s", p, ex)
        return False
    import shutil
    import tempfile

    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
            for name in xf.sheet_names:
                df = pd.read_excel(xf, sheet_name=name, header=0)
                safe = str(name)[:31] if name else "Sheet1"
                df.to_excel(writer, sheet_name=safe, index=False)
        shutil.move(tmp, p)
        logging.info(
            "Excel: xl/sharedStrings.xml を追加するため %r を openpyxl で再保存しました。",
            p,
        )
        return True
    except Exception as ex:
        if os.path.isfile(tmp):
            try:
                os.unlink(tmp)
            except OSError:
                pass
        logging.warning("Excel: sharedStrings 正規化に失敗 path=%r err=%s", p, ex)
        return False
_MASTER_PD_EXCEL_CACHE: dict[str, tuple[int, int, pd.ExcelFile]] = {}
_TABULAR_DF_LOAD_CACHE: dict[str, tuple[int, int, pd.DataFrame]] = {}
def _workbook_file_stat_sig(path: str) -> tuple[int, int] | None:
    try:
        st = os.stat(path)
        return int(st.st_mtime), int(st.st_size)
    except OSError:
        return None
def _cached_master_pd_excel_file(path: str | None = None) -> pd.ExcelFile | None:
    """同一 master.xlsm を段階2内で何度も pd.ExcelFile する I/O を避ける。"""
    global _MASTER_PD_EXCEL_CACHE
    p = (path or "").strip() or _master_workbook_path_resolved()
    if not p or not os.path.isfile(p):
        return None
    sig = _workbook_file_stat_sig(p)
    key = os.path.abspath(p)
    if sig is not None:
        prev = _MASTER_PD_EXCEL_CACHE.get(key)
        if prev and prev[0] == sig[0] and prev[1] == sig[1]:
            return prev[2]
        if prev:
            try:
                prev[2].close()
            except Exception:
                pass
    try:
        xf = pd.ExcelFile(p)
    except Exception:
        return None
    if sig is not None:
        _MASTER_PD_EXCEL_CACHE[key] = (sig[0], sig[1], xf)
    return xf
def _cached_tabular_dataframe(
    cache_namespace: str, path: str, loader
) -> pd.DataFrame:
    """ネットワーク上の加工計画DATA・実績明細等の再読込を同一プロセス内で抑える。"""
    global _TABULAR_DF_LOAD_CACHE
    p = os.path.abspath((path or "").strip())
    sig = _workbook_file_stat_sig(p) if p and os.path.isfile(p) else None
    key = f"{cache_namespace}:{p}"
    if sig is not None:
        prev = _TABULAR_DF_LOAD_CACHE.get(key)
        if prev and prev[0] == sig[0] and prev[1] == sig[1]:
            return prev[2].copy()
    df = loader()
    if sig is not None:
        _TABULAR_DF_LOAD_CACHE[key] = (sig[0], sig[1], df.copy())
    return df
def _stage2_path_debug_meta(path: str | None) -> dict:
    """段階2計測用: パス種別（UNC/ローカル等）のみ。秘密情報は含めない。"""
    p = (path or "").strip()
    if not p:
        return {"path": "", "kind": "empty"}
    try:
        ap = os.path.abspath(p)
    except OSError:
        return {"path": p[:200], "kind": "invalid"}
    norm = ap.replace("/", "\\")
    kind = "local"
    if norm.startswith("\\\\") or norm.startswith("//"):
        kind = "unc"
    elif len(ap) >= 2 and ap[1] == ":":
        kind = "drive"
    meta: dict = {"path": ap, "kind": kind}
    parent = os.path.dirname(ap)
    if parent:
        meta["parent_kind"] = (
            "unc"
            if parent.replace("/", "\\").startswith("\\\\")
            else "drive"
            if len(parent) >= 2 and parent[1] == ":"
            else "local"
        )
    try:
        if os.path.isfile(ap):
            st = os.stat(ap)
            meta["size_bytes"] = int(st.st_size)
    except OSError as ex:
        meta["stat_error"] = type(ex).__name__
    return meta
def _log_stage2_phase_timing(
    label: str, t_prev: float, *, extra: dict | None = None
) -> float:
    now = time_module.perf_counter()
    elapsed = now - t_prev
    logging.info("段階2 計測: %s %.2f秒", label, elapsed)
    return now
_STAGE2_DISPATCH_LOOP_PROFILE_DAY: dict | None = None
_STAGE2_DISPATCH_LOOP_PROFILE_RUN: dict = {}
def _reset_dispatch_loop_profile() -> None:
    global _STAGE2_DISPATCH_LOOP_PROFILE_DAY, _STAGE2_DISPATCH_LOOP_PROFILE_RUN
    _STAGE2_DISPATCH_LOOP_PROFILE_DAY = None
    _STAGE2_DISPATCH_LOOP_PROFILE_RUN = {
        "day_records": [],
        "run_buckets": {},
    }
def _dispatch_loop_profile_begin_day(iso_date: str) -> None:
    global _STAGE2_DISPATCH_LOOP_PROFILE_DAY
    _STAGE2_DISPATCH_LOOP_PROFILE_DAY = {
        "date": iso_date,
        "buckets": {},
        "trial_pass_count": 0,
        "sched_pass_secs": [],
    }
def _dispatch_loop_profile_add(bucket: str, sec: float, count: int = 1) -> None:
    if sec < 0.0:
        sec = 0.0
    day = _STAGE2_DISPATCH_LOOP_PROFILE_DAY
    if day is not None:
        ent = day["buckets"].setdefault(bucket, {"sec": 0.0, "count": 0})
        ent["sec"] += sec
        ent["count"] += int(count)
def _dispatch_loop_profile_top_buckets(
    buckets: dict, limit: int = 8
) -> list[dict]:
    rows = [
        {"bucket": k, "sec": round(v.get("sec", 0.0), 3), "count": v.get("count", 0)}
        for k, v in buckets.items()
        if v.get("sec", 0.0) > 1e-6 or v.get("count", 0) > 0
    ]
    rows.sort(key=lambda r: r["sec"], reverse=True)
    return rows[:limit]
def _dispatch_loop_profile_finish_day(
    day_elapsed_sec: float, **meta
) -> dict:
    """暦日計測を run 集計にマージし、日別サマリ dict を返す。"""
    global _STAGE2_DISPATCH_LOOP_PROFILE_DAY
    day = _STAGE2_DISPATCH_LOOP_PROFILE_DAY
    _STAGE2_DISPATCH_LOOP_PROFILE_DAY = None
    if day is None:
        return {}
    buckets = day.get("buckets") or {}
    run = _STAGE2_DISPATCH_LOOP_PROFILE_RUN
    for k, v in buckets.items():
        ent = run["run_buckets"].setdefault(k, {"sec": 0.0, "count": 0})
        ent["sec"] += v.get("sec", 0.0)
        ent["count"] += v.get("count", 0)
    pass_secs = day.get("sched_pass_secs") or []
    summary = {
        "date": day.get("date"),
        "elapsed_sec": round(day_elapsed_sec, 3),
        "trial_pass_count": day.get("trial_pass_count", 0),
        "top_buckets": _dispatch_loop_profile_top_buckets(buckets, 8),
        **meta,
    }
    if pass_secs:
        summary["sched_pass_count"] = len(pass_secs)
        summary["sched_pass_sec_max"] = round(max(pass_secs), 3)
        summary["sched_pass_sec_sum"] = round(sum(pass_secs), 3)
        if len(pass_secs) > 1:
            summary["sched_pass_sec_avg"] = round(
                sum(pass_secs) / len(pass_secs), 3
            )
    run["day_records"].append(summary)
    return summary
def _dispatch_loop_profile_emit_run_summary() -> None:
    run = _STAGE2_DISPATCH_LOOP_PROFILE_RUN
    day_records = run.get("day_records") or []
def _stage2_skip_actual_detail_gantt_prep(
    stage2_output_root, publish_plan_xlsx: bool
) -> bool:
    if stage2_output_root:
        return True
    if not publish_plan_xlsx:
        return True
    if _interactive_stage2_parity_active():
        return True
    return False
def _stage2_skip_heavy_workbook_sidecars() -> bool:
    """段階3配台試行: 論理ビュー JSON 等、UI 必須でない重い副産物を省略。"""
    return _interactive_stage2_parity_active()
def _stage2_skip_member_schedule_output() -> bool:
    """段階3配台試行: 人員*.xlsx/json は試行後 UI が必須としないため省略。"""
    return _interactive_stage2_parity_active()
def _workbook_should_skip_openpyxl_io(wb_path: str) -> bool:
    """当該パスは OOXML でシート「配台_配台不要工程」を含むとする True（openpyxl 利用を避ける）。"""
    p = (wb_path or "").strip()
    if not p:
        return False
    names = _ooxml_workbook_sheet_names(p)
    if not names:
        return False
    return OPENPYXL_INCOMPATIBLE_SHEET_MARKER in names
GEMINI_CREDENTIALS_ENCRYPTED_FILENAME = "gemini_credentials.encrypted.json"
APP_CONFIG_SHEET_NAME = "設定"
CONFIG_SHEET_TASK_IDS_FIRST_ROW = 3
CONFIG_SHEET_TASK_IDS_LAST_ROW = 26
_GEMINI_CREDENTIALS_PASSPHRASE_FIXED = "nagaoka1234"
_GEMINI_CREDENTIALS_PBKDF2_ITERATIONS_DEFAULT = 480_000
def _config_cell_text(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()
def _resolve_path_relative_to_workbook(wb_path: str, user_path: str) -> str:
    p = (user_path or "").strip().strip('"')
    if not p:
        return ""
    if os.path.isabs(p):
        return os.path.normpath(p)
    base = os.path.dirname(os.path.abspath(wb_path))
    return os.path.normpath(os.path.join(base, p))
def _gemini_credentials_json_path_next_to_workbook(wb_path: str) -> str | None:
    """マクロ実行ブックと同一フォルダの GEMINI_CREDENTIALS_ENCRYPTED_FILENAME。ブックパスが有効なファイルでない場合は None。"""
    if not wb_path or not os.path.isfile(wb_path):
        return None
    return os.path.normpath(
        os.path.join(
            os.path.dirname(os.path.abspath(wb_path)),
            GEMINI_CREDENTIALS_ENCRYPTED_FILENAME,
        )
    )
def _normalize_gemini_credentials_json_env_value(raw: str) -> str:
    """GEMINI_CREDENTIALS_JSON の値から前後空白・先頭 @・外側の引用符を除く（UI 貼り付け用）。"""
    s = (raw or "").strip()
    if s.startswith("@"):
        s = s[1:].strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
        s = s[1:-1].strip()
    return s
def _resolve_explicit_gemini_credentials_path(normalized_explicit: str) -> str:
    """
    環境変数で明示された証明書 JSON の絶対パス（無い場合も返す）。

    bootstrap が PM_AI_WORKSPACE へ chdir した後に import されるため、相対パスは cwd 基準で外れやすい。
    ファイルが見つからないときは PM_AI_REPO_ROOT 配下を追加で試す。
    """
    s = (normalized_explicit or "").strip()
    if not s:
        return ""
    primary = os.path.normpath(os.path.abspath(s))
    if os.path.isfile(primary):
        return primary
    repo = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
    if not repo:
        return primary
    try:
        repo_abs = os.path.normpath(os.path.abspath(repo))
    except Exception:
        return primary
    if not os.path.isabs(s):
        under_repo = os.path.normpath(os.path.join(repo_abs, s))
        if os.path.isfile(under_repo):
            return under_repo
        base_only = os.path.normpath(os.path.join(repo_abs, os.path.basename(s)))
        if base_only != under_repo and os.path.isfile(base_only):
            return base_only
    return primary
def _resolve_gemini_credentials_json_path() -> str | None:
    """
    Gemini 証明書 JSON の候補パス。

    1) GEMINI_CREDENTIALS_JSON … 暗号化/平文の証明書ファイルへの絶対または相対パス（最優先。先頭 @ や外側引用符を除去。相対は cwd のあと PM_AI_REPO_ROOT 基準でも探索）。
    2) PM_AI_PLAN_INPUT_PATH ブック（Excel）と同階層の GEMINI_CREDENTIALS_ENCRYPTED_FILENAME
    3) PM_AI_WORKSPACE 直下の同ファイル名（JavaFX ランチャー向け）
    """
    explicit_in = (os.environ.get("GEMINI_CREDENTIALS_JSON") or "").strip()
    explicit = _normalize_gemini_credentials_json_env_value(explicit_in)
    if explicit:
        return _resolve_explicit_gemini_credentials_path(explicit)
    plan_wb = plan_input_workbook_path_for_excel_ops()
    if plan_wb:
        return os.path.normpath(
            os.path.join(
                os.path.dirname(os.path.abspath(plan_wb)),
                GEMINI_CREDENTIALS_ENCRYPTED_FILENAME,
            )
        )
    ws = (os.environ.get("PM_AI_WORKSPACE") or "").strip()
    if ws and os.path.isdir(ws):
        return os.path.normpath(
            os.path.join(os.path.abspath(ws), GEMINI_CREDENTIALS_ENCRYPTED_FILENAME)
        )
    return None
def _read_task_ids_from_config_sheet_column(
    wb_path: str,
    column_index: int,
    log_label: str,
    column_letter_desc: str,
    *,
    openpyxl_skip_hint: str | None = None,
) -> list[str]:
    """
    マクロブック「設定」シートの指定列（1=A, 2=B）3〜26 行目から依頼NOを読む。
    空セルはスキップ。連続 30 セル空で打ち切り。カンマ区切りで複数件。
    """
    out: list[str] = []
    if not wb_path or not os.path.isfile(wb_path):
        return out
    try:
        sn = _ooxml_workbook_sheet_names(wb_path)
        if sn is not None and APP_CONFIG_SHEET_NAME not in sn:
            return out
    except Exception:
        pass
    if _workbook_should_skip_openpyxl_io(wb_path):
        msg = (
            f"{log_label}: ブックに「{OPENPYXL_INCOMPATIBLE_SHEET_MARKER}」があるため、"
            f"「{APP_CONFIG_SHEET_NAME}」!{column_letter_desc}{CONFIG_SHEET_TASK_IDS_FIRST_ROW}:"
            f"{column_letter_desc}{CONFIG_SHEET_TASK_IDS_LAST_ROW} は openpyxl で読めません。"
        )
        if openpyxl_skip_hint:
            msg += " " + openpyxl_skip_hint.strip()
        logging.info(msg)
        return out
    try:
        keep_vba = str(wb_path).lower().endswith(".xlsm")
        wb = load_workbook(
            wb_path, read_only=True, data_only=True, keep_vba=keep_vba
        )
        try:
            if APP_CONFIG_SHEET_NAME not in wb.sheetnames:
                return out
            ws = wb[APP_CONFIG_SHEET_NAME]
            consecutive_empty = 0
            for r in range(
                CONFIG_SHEET_TASK_IDS_FIRST_ROW,
                CONFIG_SHEET_TASK_IDS_LAST_ROW + 1,
            ):
                t = _config_cell_text(ws.cell(row=r, column=column_index).value)
                if not t:
                    consecutive_empty += 1
                    if consecutive_empty >= 30:
                        break
                    continue
                consecutive_empty = 0
                if "," in t:
                    for part in t.split(","):
                        p = part.strip()
                        if p:
                            out.append(p)
                else:
                    out.append(t)
        finally:
            wb.close()
    except Exception as ex:
        logging.warning(
            "%s: 「%s」!%s%d:%s%d の依頼NOを読めません（無視）: %s",
            log_label,
            APP_CONFIG_SHEET_NAME,
            column_letter_desc,
            CONFIG_SHEET_TASK_IDS_FIRST_ROW,
            column_letter_desc,
            CONFIG_SHEET_TASK_IDS_LAST_ROW,
            ex,
        )
        return []
    return out
def _read_trace_schedule_task_ids_from_config_sheet(wb_path: str) -> list[str]:
    """
    マクロブック「設定」シート A 列の 3〜26 行目を配台トレース対象の依頼NOとして読む。
    空セルはスキップ。連続 30 セル空なら打ち切り。
    """
    return _read_task_ids_from_config_sheet_column(
        wb_path,
        1,
        "配台トレース",
        "A",
        openpyxl_skip_hint="配台トレースは「設定」シート A 列を openpyxl で読めないため無効です。",
    )
def _read_debug_dispatch_task_ids_from_config_sheet(wb_path: str) -> list[str]:
    """
    マクロブック「設定」シート B 列の 3〜26 行目を段階2デバッグ配台の対象依頼NOとして読む。
    1 件も無い場合は段階2は通常モード（全件配台）。空セル・打ち切り等は A 列トレースと同様。
    """
    return _read_task_ids_from_config_sheet_column(
        wb_path,
        2,
        "デバッグ配台",
        "B",
        openpyxl_skip_hint="デバッグ配台は「設定」シート B 列を openpyxl で読めないため無効（全件配台）です。",
    )
def _show_stage2_debug_dispatch_mode_dialog(task_ids_sorted: list[str]) -> None:
    """設定シート B3:B26 が空でないことが前提。Windows では MessageBox。それ以外は WARNING ログ。"""
    if not task_ids_sorted:
        return
    preview_lines = task_ids_sorted[:30]
    preview = "\n".join(preview_lines)
    if len(task_ids_sorted) > 30:
        preview += "\n…"
    body = (
        "デバッグモードで実行した。\n\n"
        "「設定」シート B3:B26 に入力した依頼NOのみを配台対象とした。\n\n"
        "対象依頼NO:\n"
        + preview
    )
    title = "段階2（配台）— デバッグモード"
    if sys.platform != "win32":
        logging.warning("%s\n%s", title, body)
        return
    try:
        ctypes.windll.user32.MessageBoxW(0, body, title, 0x00000040)
    except Exception as ex:
        logging.warning(
            "デバッグ配台: メッセージボックスを表示できません (%s)。%s", ex, body
        )
def _extract_gemini_api_key_from_plain_dict(data: dict, json_path: str) -> str | None:
    key = data.get("gemini_api_key")
    if key is None or (isinstance(key, str) and not key.strip()):
        key = data.get("GEMINI_API_KEY")
    if key is None:
        logging.warning(
            "Gemini: 証明書データに gemini_api_key（または GEMINI_API_KEY）はありません（%s）。",
            json_path,
        )
        return None
    s = str(key).strip()
    return s or None
def _derive_fernet_key_from_passphrase(
    passphrase: str, salt: bytes, iterations: int
) -> bytes:
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.backends import default_backend
    from cryptography.hazmat.primitives import hashes

    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=iterations,
        backend=default_backend(),
    )
    return base64.urlsafe_b64encode(kdf.derive(passphrase.encode("utf-8")))
def _credentials_json_is_encrypted_v2(data: dict) -> bool:
    if data.get("format_version") == 2:
        return True
    return bool(
        data.get("kdf") == "pbkdf2_sha256" and (data.get("fernet_ciphertext") or "").strip()
    )
def _fernet_ciphertext_ascii_for_decrypt(token_s: str) -> str:
    """
    JSON に格納された Fernet トークン（ASCII）を cryptography 向けに正規化する。

    Java の Base64.getUrlEncoder().withoutPadding() で書いた ciphertext は長さが 4 の倍数でなく
    InvalidToken / Incorrect padding になり得るため、末尾に '=' を補う。
    """
    s = (token_s or "").strip().replace(" ", "").replace("\n", "").replace("\r", "")
    if not s:
        return s
    pad = (-len(s)) % 4
    if pad:
        s += "=" * pad
    return s
def _decrypt_gemini_credentials_v2(
    data: dict, passphrase: str, json_path: str
) -> dict | None:
    try:
        from cryptography.fernet import Fernet
    except ImportError:
        logging.warning(
            "Gemini: 暗号化証明書 JSON には cryptography は必須です（pip install cryptography）。"
        )
        return None
    token_s = (data.get("fernet_ciphertext") or "").strip()
    if not token_s:
        logging.warning(
            "Gemini: 暗号化証明書 JSON に fernet_ciphertext はありません（%s）。",
            json_path,
        )
        return None
    salt_b64 = (data.get("salt_b64") or "").strip()
    if not salt_b64:
        logging.warning(
            "Gemini: 暗号化証明書 JSON に salt_b64 はありません（%s）。",
            json_path,
        )
        return None
    try:
        salt = base64.standard_b64decode(salt_b64)
    except Exception as ex:
        logging.warning("Gemini: salt_b64 の解釈に失敗しました（%s）: %s", json_path, ex)
        return None
    iterations = int(data.get("iterations") or _GEMINI_CREDENTIALS_PBKDF2_ITERATIONS_DEFAULT)
    kdf_name = (data.get("kdf") or "pbkdf2_sha256").strip()
    if kdf_name != "pbkdf2_sha256":
        logging.warning("Gemini: 未対応の kdf（%s）: %s", kdf_name, json_path)
        return None
    try:
        fkey = _derive_fernet_key_from_passphrase(passphrase, salt, iterations)
        token_norm = _fernet_ciphertext_ascii_for_decrypt(token_s)
        plain = Fernet(fkey).decrypt(token_norm.encode("ascii"))
    except Exception:
        logging.debug("Gemini: 暗号化証明書の復号処理に失敗しました（%s）。", json_path)
        return None
    try:
        inner = json.loads(plain.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as ex:
        logging.warning(
            "Gemini: 復号後の JSON は正常です（%s）: %s",
            json_path,
            ex,
        )
        return None
    if not isinstance(inner, dict):
        logging.warning("Gemini: 復号後の JSON はオブジェクトである必須はありした（%s）。", json_path)
        return None
    return inner
def _load_gemini_api_key_from_credentials_json(
    json_path: str, workbook_path: str | None = None
) -> tuple[str | None, bool]:
    """戻り値: (api_key または None, 暗号化形式であったか)。暗号化時は _GEMINI_CREDENTIALS_PASSPHRASE_FIXED のみで復号。"""
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
    except OSError as ex:
        logging.warning("Gemini: 証明書 JSON を開きません: %s (%s)", json_path, ex)
        return None, False
    except json.JSONDecodeError as ex:
        logging.warning("Gemini: 証明書 JSON の形式は正常です: %s (%s)", json_path, ex)
        return None, False
    if not isinstance(data, dict):
        logging.warning("Gemini: 証明書 JSON はオブジェクト形式である必須はありした: %s", json_path)
        return None, False
    if _credentials_json_is_encrypted_v2(data):
        inner = _decrypt_gemini_credentials_v2(
            data, _GEMINI_CREDENTIALS_PASSPHRASE_FIXED, json_path
        )
        if inner is None:
            return None, True
        return _extract_gemini_api_key_from_plain_dict(inner, json_path), True
    return _extract_gemini_api_key_from_plain_dict(data, json_path), False
API_KEY = None
_cred_path = _resolve_gemini_credentials_json_path()
_used_encrypted_credentials = False
if _cred_path and os.path.isfile(_cred_path):
    API_KEY, _used_encrypted_credentials = _load_gemini_api_key_from_credentials_json(
        _cred_path,
        workbook_path=_excel_plan_input_wb()
        or (os.environ.get("PM_AI_WORKSPACE") or "").strip(),
    )
    if API_KEY:
        if _used_encrypted_credentials:
            logging.info("Gemini API キー: 暗号化証明書 JSON から読み込みました。")
        else:
            logging.info(
                "Gemini API キー: マクロブック同階層の「%s」から読み込みました。",
                GEMINI_CREDENTIALS_ENCRYPTED_FILENAME,
            )
elif _cred_path:
    logging.warning(
        "Gemini: マクロブック同階層の「%s」は見つかりません。",
        GEMINI_CREDENTIALS_ENCRYPTED_FILENAME,
    )
_encrypted_json_missing_key = (
    bool(_cred_path)
    and os.path.isfile(_cred_path)
    and _used_encrypted_credentials
    and not API_KEY
)
if _encrypted_json_missing_key:
    logging.error(
        "Gemini: マクロブック同階層の「%s」から API キーを利用できません。"
        " 社内手順に従い証明書を再設定れるか」管理者に相い合わせでしてさい。",
        GEMINI_CREDENTIALS_ENCRYPTED_FILENAME,
    )
if not API_KEY:
    logging.warning(
        "Gemini API キーは未設定です。GEMINI_CREDENTIALS_JSON で証明書ファイルを指定するか、"
        "マクロ実行ブックと同一フォルダ、または PM_AI_WORKSPACE に「%s」を配置してください。"
        " 備考の AI 解析等はスキップされした。"
        " 参考型: gemini_credentials.example.json / 参照用/python/encrypt_gemini_credentials.py（暗号化）。",
        GEMINI_CREDENTIALS_ENCRYPTED_FILENAME,
    )
GANTT_TIMELINE_SLOT_MINUTES = 10
GANTT_TIMELINE_COLUMN_WIDTH = 3
GANTT_LABEL_SHAPE_MIN_TIMELINE_COLUMNS = 3
GANTT_HDR_ROW_HEIGHT_PT = int(float(os.environ.get("GANTT_HDR_ROW_HEIGHT_PT", "38")))
GANTT_MACHINE_ROW_HEIGHT_PT = int(float(os.environ.get("GANTT_MACHINE_ROW_HEIGHT_PT", "60")))
def _effective_result_book_font_name() -> str:
    """環境変数 PM_AI_RESULT_BOOK_FONT があれば優先（JavaFX 実行タブのフォント選択など）。"""
    v = os.environ.get("PM_AI_RESULT_BOOK_FONT", "").strip()
    return v if v else RESULT_BOOK_FONT_NAME
def _gemini_err_text_for_exc(exc: BaseException) -> str:
    parts = [str(exc), repr(exc)]
    for attr in ("status_code", "code", "message"):
        v = getattr(exc, attr, None)
        if v is not None:
            parts.append(str(v))
    return " ".join(parts)
def _gemini_is_transient_api_error(err_text: str) -> bool:
    """503 / 過負荷 / 期限切れなど、待てば再試行に値する API 失敗。"""
    t = err_text.upper()
    if "429" in err_text:
        return True
    if "503" in err_text:
        return True
    if "504" in err_text:
        return True
    for needle in (
        "UNAVAILABLE",
        "RESOURCE_EXHAUSTED",
        "DEADLINE_EXCEEDED",
        "DEADLINE EXCEEDED",
        "SERVICE UNAVAILABLE",
        "INTERNAL ERROR",
        "UNRECOVERABLE",
    ):
        if needle in t:
            return True
    return False
def _gemini_is_quota_style_error(err_text: str) -> bool:
    t = err_text.upper()
    return ("429" in err_text) or ("RESOURCE_EXHAUSTED" in t)
def _gemini_is_timeout_error(exc: BaseException, err_text: str) -> bool:
    """HTTP 読み取りタイムアウト・接続タイムアウト等（応答が期限内に返らない）。"""
    if isinstance(exc, TimeoutError):
        return True
    tn = type(exc).__name__
    if tn in ("ReadTimeout", "ConnectTimeout", "WriteTimeout", "PoolTimeout"):
        return True
    u = err_text.upper()
    if "READ TIMEOUT" in u or "CONNECT TIMEOUT" in u or "WRITE TIMEOUT" in u:
        return True
    if "TIMED OUT" in u:
        return True
    # 504 等も TIMEOUT を含むが、DEADLINE_EXCEEDED 単体は一時エラー扱いに任せる
    if "TIMEOUT" in u and "DEADLINE_EXCEEDED" not in u:
        return True
    return False
def _gemini_try_order_from_env() -> tuple[str, ...] | None:
    raw = (os.environ.get("GEMINI_MODEL_TRY_ORDER") or "").strip()
    if not raw:
        return None
    parts = tuple(p.strip() for p in raw.split(",") if p.strip())
    return parts or None
def _gemini_effective_model_chain(model: str | None) -> tuple[str, ...]:
    """引数 model があればそれのみ。なければ GEMINI_MODEL、GEMINI_MODEL_TRY_ORDER、コード既定の順で決定。"""
    if model is not None and str(model).strip():
        return (str(model).strip(),)
    pinned = (os.environ.get("GEMINI_MODEL") or "").strip()
    if pinned:
        return (pinned,)
    ovr = _gemini_try_order_from_env()
    if ovr is not None:
        return ovr
    return GEMINI_MODEL_IDS_BY_QUALITY
def _gemini_is_model_endpoint_unavailable_error(err_text: str) -> bool:
    """モデル未提供・モデル名不正など、別モデルでの再試行が合理的な失敗。"""
    t = err_text.upper()
    u = err_text.lower()
    if "NOT_FOUND" in t and ("MODEL" in t or "MODELS/" in t):
        return True
    if "404" in err_text and "model" in u:
        return True
    if ("DOES NOT EXIST" in t or "WAS NOT FOUND" in t) and "model" in u:
        return True
    return False
def _gemini_pre_request_jitter_sleep() -> None:
    mx = max(0.0, _GEMINI_PRE_REQUEST_JITTER_MAX)
    if mx <= 0.0:
        return
    time_module.sleep(random.uniform(0.0, mx))
def _gemini_progress_log_interval_sec() -> float:
    """Gemini 応答待ち中に INFO を出す間隔（秒）。0 以下でハートビート無効（送信ログのみ）。"""
    try:
        return float((os.environ.get("GEMINI_PROGRESS_LOG_INTERVAL_SEC") or "12").strip())
    except (TypeError, ValueError):
        return 12.0
def _gemini_flush_log_handlers() -> None:
    try:
        for h in logging.getLogger().handlers:
            flush = getattr(h, "flush", None)
            if flush is not None:
                flush()
    except Exception:
        pass
    try:
        sys.stdout.flush()
    except Exception:
        pass
def _gemini_heartbeat_loop(
    stop: threading.Event, prefix: str, model_id: str, interval_sec: float
) -> None:
    """ブロッキング中でもターミナルが固まって見えないよう、一定間隔で待機ログを出す。"""
    start = time_module.monotonic()
    while True:
        if stop.wait(timeout=interval_sec):
            break
        elapsed = time_module.monotonic() - start
        logging.info(
            "%sGemini 応答待ち... 約%.0f秒経過（モデル: %s）",
            prefix,
            elapsed,
            model_id,
        )
        _gemini_flush_log_handlers()
def _gemini_thinking_budget() -> int:
    """``GEMINI_THINKING_BUDGET``: 0=思考無効（既定）、-1=モデル任せ、正数=思考トークン上限。

    抽出・転記系のプロンプトに思考は不要で、思考トークンはそのまま応答待ち時間になる。
    """
    raw = (os.environ.get("GEMINI_THINKING_BUDGET") or "").strip()
    if not raw:
        return 0
    try:
        return int(float(raw))
    except (TypeError, ValueError):
        return 0
def _gemini_generate_content_config(
    *,
    response_schema=None,
    max_output_tokens: int | None = None,
    thinking_budget: int | None = None,
):
    """``generate_content`` の設定。思考を既定で無効化し、スキーマ指定時は JSON 出力に固定する。

    SDK に ``types`` が無い等で組み立てられないときは ``None``（SDK 既定で送信）。
    """
    try:
        from google.genai import types as genai_types
    except Exception:
        logging.debug("Gemini: google.genai.types を読み込めないため生成設定を省略します。", exc_info=True)
        return None
    kwargs: dict = {}
    budget = _gemini_thinking_budget() if thinking_budget is None else int(thinking_budget)
    if budget >= 0:
        kwargs["thinking_config"] = genai_types.ThinkingConfig(thinking_budget=budget)
    if response_schema is not None:
        kwargs["response_mime_type"] = "application/json"
        kwargs["response_schema"] = response_schema
    if max_output_tokens is not None and int(max_output_tokens) > 0:
        kwargs["max_output_tokens"] = int(max_output_tokens)
    try:
        return genai_types.GenerateContentConfig(**kwargs)
    except Exception:
        logging.debug("Gemini: GenerateContentConfig の生成に失敗したため既定設定で送信します。", exc_info=True)
        return None
def _gemini_config_without_thinking(config):
    """思考設定だけを外した複製（モデルが thinkingBudget を受け付けないとき用）。"""
    if config is None:
        return None
    try:
        clone = config.model_copy(update={"thinking_config": None})
    except Exception:
        try:
            clone = copy.copy(config)
            clone.thinking_config = None
        except Exception:
            return None
    return clone
def _gemini_is_thinking_config_unsupported_error(err_text: str) -> bool:
    """thinkingBudget 非対応モデルが返す 400（思考設定を外せば通る）。

    gemini-3.5-flash-lite は理由を書かず ``Request contains an invalid argument.`` だけを返すため、
    400 INVALID_ARGUMENT なら思考設定が原因とみなして一度外して試す。原因が別なら同じ 400 で
    落ちるだけで、通常の再試行・モデル切り替えに戻る。
    """
    u = err_text.lower()
    if not u:
        return False
    if "invalid_argument" in u or "400" in u:
        return True
    return "thinking" in u and (("not supported" in u) or ("unsupported" in u))
_GEMINI_THINKING_REJECTED_LOCK = threading.Lock()
_gemini_models_rejecting_thinking_config: set[str] = set()
def _gemini_forget_thinking_config_rejections() -> None:
    """思考設定を拒んだモデルの記憶を消す（テスト用）。"""
    with _GEMINI_THINKING_REJECTED_LOCK:
        _gemini_models_rejecting_thinking_config.clear()
def _gemini_model_rejects_thinking_config(model_id: str) -> bool:
    with _GEMINI_THINKING_REJECTED_LOCK:
        return str(model_id or "") in _gemini_models_rejecting_thinking_config
def _gemini_remember_thinking_config_rejection(model_id: str) -> None:
    with _GEMINI_THINKING_REJECTED_LOCK:
        _gemini_models_rejecting_thinking_config.add(str(model_id or ""))
def _gemini_parse_json_object(text) -> dict | None:
    """応答テキストから JSON オブジェクトを取り出す（コードフェンス・前後の地の文を許容）。"""
    s = str(text or "").strip()
    if not s:
        return None
    try:
        parsed = json.loads(s)
    except (TypeError, ValueError):
        m = re.search(r"\{.*\}", s, re.DOTALL)
        if not m:
            return None
        try:
            parsed = json.loads(m.group(0))
        except (TypeError, ValueError):
            return None
    return parsed if isinstance(parsed, dict) else None
GEMINI_FREE_TIER_RPM_LIMIT = 15
class _GeminiRateLimiter:
    """直近 ``window_sec`` 秒あたり ``limit`` 件までに送信を絞る（無料枠 RPM 超過の予防）。

    429 を受けてからの再試行はサーバ指定の待機（30〜60 秒）に従うため、超過してから直すと
    並列化で稼いだ時間を失う。送信側で先に間隔を空けるほうが速い。
    """

    def __init__(self, limit: int, window_sec: float = 60.0):
        self._limit = max(1, int(limit))
        self._window = max(0.0, float(window_sec))
        self._lock = threading.Lock()
        self._sent: list[float] = []

    def acquire(self) -> float:
        """枠が空くまでブロックし、待った秒数を返す。"""
        waited = 0.0
        while True:
            with self._lock:
                now = time_module.monotonic()
                cutoff = now - self._window
                self._sent = [t for t in self._sent if t > cutoff]
                if len(self._sent) < self._limit:
                    self._sent.append(now)
                    return waited
                sleep_for = self._window - (now - self._sent[0])
            sleep_for = max(0.01, sleep_for)
            time_module.sleep(sleep_for)
            waited += sleep_for
def _gemini_requests_per_minute() -> int:
    """1 分あたりの送信上限（``GEMINI_REQUESTS_PER_MINUTE``、既定は無料枠 RPM）。"""
    raw = (os.environ.get("GEMINI_REQUESTS_PER_MINUTE") or "").strip()
    try:
        v = int(float(raw)) if raw else GEMINI_FREE_TIER_RPM_LIMIT
    except (TypeError, ValueError):
        v = GEMINI_FREE_TIER_RPM_LIMIT
    return max(1, v)
_GEMINI_RATE_LIMITER_LOCK = threading.Lock()
_gemini_rate_limiter: _GeminiRateLimiter | None = None
_gemini_rate_limiter_limit: int | None = None
def _gemini_acquire_request_slot(prefix: str = "") -> float:
    """送信直前に RPM 枠を確保する。プロセス内の全 Gemini 呼び出しで共有する。"""
    global _gemini_rate_limiter, _gemini_rate_limiter_limit
    limit = _gemini_requests_per_minute()
    with _GEMINI_RATE_LIMITER_LOCK:
        if _gemini_rate_limiter is None or _gemini_rate_limiter_limit != limit:
            _gemini_rate_limiter = _GeminiRateLimiter(limit)
            _gemini_rate_limiter_limit = limit
        limiter = _gemini_rate_limiter
    waited = limiter.acquire()
    if waited > 0:
        logging.info(
            "%sGemini 送信レート制限（%s 件/分）のため %.1f 秒待機しました。",
            prefix,
            limit,
            waited,
        )
    return waited
def _gemini_batch_max_items() -> int:
    """1 リクエストに載せる明細行の上限（``GEMINI_BATCH_MAX_ITEMS``、既定 120）。

    1 回の出力トークン量が応答時間を決めるため、行数で頭打ちにする。
    """
    raw = (os.environ.get("GEMINI_BATCH_MAX_ITEMS") or "").strip()
    try:
        v = int(float(raw)) if raw else 120
    except (TypeError, ValueError):
        v = 120
    return max(1, v)
def _gemini_max_parallel_requests() -> int:
    """バッチ並列数（``GEMINI_MAX_PARALLEL_REQUESTS``、既定 4）。無料枠 RPM を超えない範囲に丸める。"""
    raw = (os.environ.get("GEMINI_MAX_PARALLEL_REQUESTS") or "").strip()
    try:
        v = int(float(raw)) if raw else 4
    except (TypeError, ValueError):
        v = 4
    return max(1, min(v, GEMINI_FREE_TIER_RPM_LIMIT))
def _gemini_batch_slices(total: int, batch_size: int) -> list[tuple[int, int]]:
    """``[0, total)`` を ``batch_size`` ごとの半開区間に分割する。"""
    n = max(0, int(total))
    if n <= 0:
        return []
    size = int(batch_size)
    if size <= 0:
        return [(0, n)]
    return [(s, min(s + size, n)) for s in range(0, n, size)]
def _gemini_invoke_generate_content(client: genai.Client, model_id: str, contents, config):
    """1 回分の ``generate_content``。思考設定を拒むモデルのときだけ設定を外して即再送する。

    一度拒まれたモデルはプロセス内で覚えておき、以降は最初から思考設定なしで送る
    （バッチごとに無駄な 400 を踏まないため）。

    戻り値: ``(応答, 実際に通った設定)``。呼び出し側は次の試行から後者を使う。
    """
    if (
        config is not None
        and getattr(config, "thinking_config", None) is not None
        and _gemini_model_rejects_thinking_config(model_id)
    ):
        config = _gemini_config_without_thinking(config)
    try:
        if config is None:
            return client.models.generate_content(model=model_id, contents=contents), None
        return (
            client.models.generate_content(model=model_id, contents=contents, config=config),
            config,
        )
    except Exception as e:
        if config is None or getattr(config, "thinking_config", None) is None:
            raise
        if not _gemini_is_thinking_config_unsupported_error(_gemini_err_text_for_exc(e)):
            raise
        logging.warning(
            "Gemini モデル %s は thinkingBudget 指定を受け付けないため、思考設定を外して再送します。",
            model_id,
        )
        _gemini_remember_thinking_config_rejection(model_id)
        fallback = _gemini_config_without_thinking(config)
        if fallback is None:
            return client.models.generate_content(model=model_id, contents=contents), None
        return (
            client.models.generate_content(model=model_id, contents=contents, config=fallback),
            fallback,
        )
_GEMINI_USAGE_RECORD_LOCK = threading.Lock()
def _gemini_generate_json_map_in_batches(
    client: genai.Client,
    *,
    items,
    build_prompt,
    log_label: str = "",
    batch_size: int | None = None,
    max_workers: int | None = None,
    response_schema=None,
    max_output_tokens: int | None = None,
    parse_map=None,
) -> tuple[dict, list[str], int]:
    """明細を分割して並列に ``generate_content`` し、JSON オブジェクトをマージする。

    1 リクエストの出力トークンを抑えることが目的で、並列化はその副次的な時間短縮。
    1 バッチが失敗しても他バッチの結果は捨てない。

    戻り値: ``(マージ済み dict, 使用モデル ID のリスト, 失敗バッチ数)``
    """
    rows = list(items or [])
    if not rows:
        return {}, [], 0
    size = _gemini_batch_max_items() if batch_size is None else int(batch_size)
    slices = _gemini_batch_slices(len(rows), size)
    workers = _gemini_max_parallel_requests() if max_workers is None else int(max_workers)
    workers = max(1, min(workers, len(slices)))
    to_map = parse_map if parse_map is not None else (lambda payload: payload)
    prefix = f"{log_label}: " if log_label else ""
    logging.info(
        "%sGemini へ %s 件を %s バッチ（並列 %s）で送信します。",
        prefix,
        len(rows),
        len(slices),
        workers,
    )

    def _run_one(index: int, lo: int, hi: int) -> tuple[int, dict, str | None]:
        chunk = rows[lo:hi]
        label = f"{log_label} {index + 1}/{len(slices)}" if log_label else f"{index + 1}/{len(slices)}"
        config = _gemini_generate_content_config(
            response_schema=response_schema, max_output_tokens=max_output_tokens
        )
        res, model_id = _gemini_generate_content_with_retry(
            client, contents=build_prompt(chunk), log_label=label, config=config
        )
        with _GEMINI_USAGE_RECORD_LOCK:
            record_gemini_response_usage(res, model_id)
        payload = _gemini_parse_json_object(_gemini_result_text(res))
        if payload is None:
            logging.warning("%sバッチ %s の応答を JSON として読めませんでした。", prefix, index + 1)
            return index, {}, model_id
        parsed = to_map(payload)
        return index, (parsed if isinstance(parsed, dict) else {}), model_id

    from concurrent.futures import ThreadPoolExecutor

    merged: dict = {}
    model_ids: list[str | None] = [None] * len(slices)
    failed = 0
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="gemini-batch") as pool:
        futures = [pool.submit(_run_one, i, lo, hi) for i, (lo, hi) in enumerate(slices)]
        for fut in futures:
            try:
                index, parsed, model_id = fut.result()
            except GeminiApiSkippedError:
                raise
            except Exception as ex:
                failed += 1
                logging.warning("%sバッチ呼び出しに失敗しました（続行）: %s", prefix, ex)
                continue
            model_ids[index] = model_id
            merged.update(parsed)
    used = [m for m in model_ids if m]
    logging.info(
        "%sGemini バッチ完了: 取得 %s 件・失敗 %s バッチ。",
        prefix,
        len(merged),
        failed,
    )
    return merged, used, failed
def _gemini_generate_content_with_retry(
    client: genai.Client,
    *,
    contents,
    model: str | None = None,
    max_attempts: int | None = None,
    log_label: str = "",
    config=None,
):
    """generate_content を再試行する（Gemini generateContent 共通）。

    - モデル列: 環境変数 **GEMINI_MODEL**（単一固定）→ **GEMINI_MODEL_TRY_ORDER**（カンマ区切り）→
      **GEMINI_MODEL_IDS_BY_QUALITY**（コード既定: gemini-3.5-flash 最優先＋Flash-Lite 系を新→旧）。引数 ``model`` を渡したときはその1件のみ。
    - 同一モデルあたり最大 _GEMINI_RETRY_MAX_ATTEMPTS 回（既定 3、GEMINI_RETRY_MAX_ATTEMPTS で変更）。
      そのモデルで試行を使い切ったら、列の次のモデルへ進む（試すモデルがなくなるまで）。
    - モデル未提供（404 等）は直ちに次モデルへ進む。
    - 各試行の直前: 0〜_GEMINI_PRE_REQUEST_JITTER_MAX の乱数待機（同時リクエストのばらつき）
    - 一時エラー待機: (1) 429 等で本文に retry 秒数 (2) 指数バックオフ＋ジッター
    - HTTP タイムアウト（既定 60 秒・GEMINI_REQUEST_TIMEOUT_SEC）: 同一モデルに残試行があれば短待機で再試行。
      試行を使い切ったらモデル列の次点へ進む（_gemini_client の HttpOptions と併用）。
    - 生成設定: 引数 ``config`` 未指定なら ``_gemini_generate_content_config()``（思考は既定で無効）。

    戻り値: (応答オブジェクト, 実際に成功したモデル ID)
    """
    effective_config = config if config is not None else _gemini_generate_content_config()
    chain = _gemini_effective_model_chain(model)
    n = max_attempts if max_attempts is not None else _GEMINI_RETRY_MAX_ATTEMPTS
    if n < 1:
        n = 1
    base = max(0.1, float(_GEMINI_RETRY_BACKOFF_BASE))
    prefix = f"{log_label}: " if log_label else ""
    if _stage2_truthy_env("PM_AI_SKIP_GEMINI_API"):
        logging.warning(
            "%sGemini API 呼び出しをスキップしました（PM_AI_SKIP_GEMINI_API=1）。",
            prefix,
        )
        raise GeminiApiSkippedError("PM_AI_SKIP_GEMINI_API")
    hb_interval = _gemini_progress_log_interval_sec()
    last_raise: BaseException | None = None
    for mi, mid in enumerate(chain):
        cur_config = effective_config
        for attempt in range(n):
            _gemini_pre_request_jitter_sleep()
            _gemini_acquire_request_slot(prefix)
            try:
                logging.info("%sGemini API を呼び出し中（モデル: %s）", prefix, mid)
                _gemini_flush_log_handlers()
                stop_hb = threading.Event()
                hb_thread: threading.Thread | None = None
                if hb_interval > 0:
                    hb_thread = threading.Thread(
                        target=_gemini_heartbeat_loop,
                        args=(stop_hb, prefix, mid, hb_interval),
                        name="gemini-progress-hb",
                        daemon=True,
                    )
                    hb_thread.start()
                t_req = time_module.monotonic()
                try:
                    res, cur_config = _gemini_invoke_generate_content(
                        client, mid, contents, cur_config
                    )
                finally:
                    stop_hb.set()
                    if hb_thread is not None:
                        hb_thread.join(timeout=2.0)
                elapsed_req = time_module.monotonic() - t_req
                logging.info(
                    "%sGemini API 応答を受信しました（約%.1f秒、モデル: %s）",
                    prefix,
                    elapsed_req,
                    mid,
                )
                _gemini_flush_log_handlers()
                return res, mid
            except Exception as e:
                err_text = _gemini_err_text_for_exc(e)
                if _gemini_is_model_endpoint_unavailable_error(err_text) and mi < len(chain) - 1:
                    logging.warning(
                        "%sGemini モデル %s が利用できません: %s — 次モデルへ切り替えます。",
                        prefix,
                        mid,
                        err_text[:800],
                    )
                    last_raise = e
                    break
                if _gemini_is_timeout_error(e, err_text):
                    last_raise = e
                    if attempt < n - 1:
                        wait_sec = min(2.0 + random.uniform(0.0, 1.0), 5.0)
                        logging.warning(
                            "%sGemini API タイムアウト（モデル %s 試行 %s/%s）: %s — %.1f 秒待機して再試行します。",
                            prefix,
                            mid,
                            attempt + 1,
                            n,
                            err_text[:800],
                            wait_sec,
                        )
                        time_module.sleep(wait_sec)
                        continue
                    if mi < len(chain) - 1:
                        logging.warning(
                            "%sGemini API タイムアウト（モデル %s）: %s — 次モデルへ切り替えます。",
                            prefix,
                            mid,
                            err_text[:800],
                        )
                        break
                    raise
                if _gemini_is_transient_api_error(err_text) and attempt < n - 1:
                    wait_sec = None
                    if _gemini_is_quota_style_error(err_text):
                        rs = extract_retry_seconds(err_text)
                        if rs is not None:
                            wait_sec = min(max(rs, 1.0), 120.0) + random.uniform(0.0, 1.5)
                    if wait_sec is None:
                        pow_part = base * (2**attempt)
                        jitter = random.uniform(0.0, min(4.0, base * 2.0))
                        wait_sec = min(pow_part + jitter, 90.0)
                    logging.warning(
                        "%sGemini API 一時エラー（モデル %s 試行 %s/%s）: %s — %.1f 秒待機して再試行します。",
                        prefix,
                        mid,
                        attempt + 1,
                        n,
                        err_text[:800],
                        wait_sec,
                    )
                    time_module.sleep(wait_sec)
                    continue
                if attempt < n - 1:
                    wait_sec = min(2.0 + random.uniform(0.0, 1.0), 5.0)
                    logging.warning(
                        "%sGemini API エラー（モデル %s 試行 %s/%s）: %s — %.1f 秒待機して再試行します。",
                        prefix,
                        mid,
                        attempt + 1,
                        n,
                        err_text[:800],
                        wait_sec,
                    )
                    time_module.sleep(wait_sec)
                    continue
                if mi < len(chain) - 1:
                    logging.warning(
                        "%sGemini モデル %s が %s 回とも失敗したため次モデルへ切り替えます: %s",
                        prefix,
                        mid,
                        n,
                        err_text[:800],
                    )
                    last_raise = e
                    break
                raise
    if last_raise is not None:
        raise last_raise
    raise RuntimeError("Gemini: モデル列が空です。")
def _gemini_result_text(res):
    try:
        return (res.text or "").strip()
    except Exception:
        return ""
_gemini_usage_session: dict[str, dict[str, int]] = {}
def _gemini_cumulative_json_path() -> str:
    path = os.path.join(api_payment_dir, GEMINI_USAGE_CUMULATIVE_JSON_FILE)
    legacy = os.path.join(log_dir, GEMINI_USAGE_CUMULATIVE_JSON_FILE)
    if os.path.isfile(legacy) and not os.path.isfile(path):
        try:
            shutil.move(legacy, path)
        except OSError:
            pass
    return path
def _load_gemini_cumulative_payload() -> dict:
    """API_Payment 内の累計 JSON を読む。無い・壊れでいれみ初期形を返す。"""
    path = _gemini_cumulative_json_path()
    default: dict = {
        "version": 1,
        "updated_at": None,
        "calls_total": 0,
        "prompt_total": 0,
        "candidates_total": 0,
        "thoughts_total": 0,
        "total_tokens_reported": 0,
        "estimated_cost_usd_total": 0.0,
        "by_model": {},
    }
    if not os.path.isfile(path):
        _gemini_buckets_ensure_structure(default)
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict) or int(data.get("version") or 0) != 1:
            _gemini_buckets_ensure_structure(default)
            return default
        data.setdefault("by_model", {})
        _gemini_buckets_ensure_structure(data)
        for k in (
            "calls_total",
            "prompt_total",
            "candidates_total",
            "thoughts_total",
            "total_tokens_reported",
        ):
            data[k] = int(data.get(k) or 0)
        data["estimated_cost_usd_total"] = float(data.get("estimated_cost_usd_total") or 0.0)
        return data
    except Exception:
        _gemini_buckets_ensure_structure(default)
        return default
def _gemini_buckets_ensure_structure(data: dict) -> None:
    """累計 JSON に期間別ポケット用の辞書を用意れる（既存 v1 ファイルもマージ）。"""
    b = data.setdefault("buckets", {})
    if not isinstance(b, dict):
        b = {}
        data["buckets"] = b
    for sub in ("by_year", "by_month", "by_week", "by_day", "by_hour"):
        x = b.setdefault(sub, {})
        if not isinstance(x, dict):
            b[sub] = {}
    b.setdefault(
        "timezone_note",
        "period_key は PC ローカル時刻（datetime.now）で付与。他 PC との集計は混ずないでしてさい。",
    )
def _gemini_time_bucket_keys(dt: datetime) -> tuple[str, str, str, str, str]:
    """年・月・ISO週・日・時 のキー（文字列ソートで時系列比較しやれい形）。"""
    iy, iw, _ = dt.isocalendar()
    y = dt.strftime("%Y")
    ym = dt.strftime("%Y-%m")
    wk = f"{iy}-W{iw:02d}"
    d = dt.strftime("%Y-%m-%d")
    h = dt.strftime("%Y-%m-%dT%H")
    return y, ym, wk, d, h
def _gemini_bucket_add_one_call(
    buckets_root: dict,
    pt: int,
    ct: int,
    th: int,
    tt: int,
    inc_usd: float | None,
    *,
    when: datetime | None = None,
) -> None:
    """1 回の API 呼出しを年・月・週・日・時の坄ポケットに加算れる。"""
    dt = when or datetime.now()
    y, ym, wk, d, h = _gemini_time_bucket_keys(dt)
    pairs = (
        ("by_year", y),
        ("by_month", ym),
        ("by_week", wk),
        ("by_day", d),
        ("by_hour", h),
    )
    for sub, pk in pairs:
        subd = buckets_root.setdefault(sub, {})
        ent = subd.setdefault(
            pk,
            {
                "calls": 0,
                "prompt": 0,
                "candidates": 0,
                "thoughts": 0,
                "total_tokens": 0,
                "estimated_cost_usd": 0.0,
            },
        )
        ent["calls"] = int(ent.get("calls") or 0) + 1
        ent["prompt"] = int(ent.get("prompt") or 0) + pt
        ent["candidates"] = int(ent.get("candidates") or 0) + ct
        ent["thoughts"] = int(ent.get("thoughts") or 0) + th
        ent["total_tokens"] = int(ent.get("total_tokens") or 0) + tt
        if inc_usd is not None:
            ent["estimated_cost_usd"] = float(
                ent.get("estimated_cost_usd") or 0.0
            ) + float(inc_usd)
def _gemini_estimate_cost_usd(
    model_id: str, prompt_tok: int, cand_tok: int, thoughts_tok: int
) -> float | None:
    m = str(model_id).strip().lower()
    rin, rout = None, None
    if "flash" in m:
        rin, rout = _GEMINI_FLASH_IN_PER_M, _GEMINI_FLASH_OUT_PER_M
    elif "pro" in m:
        # 目安（未使用モデル坑けフォールバック）
        rin, rout = 1.25, 5.0
    if rin is None:
        return None
    out_equiv = cand_tok + thoughts_tok
    return (prompt_tok / 1_000_000.0) * rin + (out_equiv / 1_000_000.0) * rout
def _gemini_daily_trend_series(
    cum: dict, *, max_days: int | None = None
) -> tuple[list[str], list[float], str] | None:
    """累計 JSON の by_day から」日付キー（坤→新）・値・系列名。無ければ None。"""
    lim = 14 if max_days is None else max_days
    b = cum.get("buckets")
    if not isinstance(b, dict):
        return None
    subd = b.get("by_day")
    if not isinstance(subd, dict) or not subd:
        return None
    keys = sorted(subd.keys())
    keys = keys[-max(1, lim) :]
    usds: list[float] = []
    calls: list[int] = []
    for pk in keys:
        ent = subd.get(pk)
        if isinstance(ent, dict):
            usds.append(float(ent.get("estimated_cost_usd") or 0.0))
            calls.append(int(ent.get("calls") or 0))
        else:
            usds.append(0.0)
            calls.append(0)
    use_calls = sum(usds) <= 0.0 and sum(calls) > 0
    series = [float(c) for c in calls] if use_calls else usds
    label = "呼出し回数" if use_calls else "推定USD"
    return (keys, series, label)
def _gemini_usage_trend_caption_lines(cum: dict) -> list[str]:
    """累計 JSON の日次推移をテキストサマリ用に要約する。"""
    ser = _gemini_daily_trend_series(cum)
    if ser is None:
        return []
    keys, _, label = ser
    b = cum.get("buckets")
    lines = [
        f"【日次推移】{label}（{keys[0]} ～ {keys[-1]}）",
    ]
    if isinstance(b, dict):
        note = b.get("timezone_note")
        if note:
            lines.append(f"  （{note}）")
    return lines
def _gemini_kv_table_lines(title: str, rows: list[tuple[str, str]]) -> list[str]:
    """累計・当実行坑けの 2 列テキスト表（履歴行は含まない）。"""
    out = [title]
    if not rows:
        return out
    lw = min(22, max(len(a) for a, _ in rows))
    sep = "  " + ("─" * (lw + 2 + 28))
    out.append(sep)
    for a, b in rows:
        out.append(f"  {a:<{lw}}  {b}")
    return out

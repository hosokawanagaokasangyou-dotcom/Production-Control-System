# -*- coding: utf-8 -*-
# planning_core.core.state — body only (loaded via _core exec chain)
_STAGE2_GLOBAL_COMMENT_CACHE: dict | None = None
_STAGE2_MACHINE_CALENDAR_CACHE: dict | None = None
_MH_HANDOFF_TIMELINE_CACHE_KEY: tuple[int, date] | None = None
_MH_HANDOFF_TIMELINE_CACHE_STATE: dict | None = None
_MH_HANDOFF_TIMELINE_CACHE_STATS: dict[str, int] = {"hit": 0, "miss": 0, "incremental": 0}
_STAGE2_DISPATCH_EVENTS_BY_DATE: dict | None = None
_INTERACTIVE_TRIAL_OP_SHORTAGE: list[dict] = []
_INTERACTIVE_TRIAL_AS_SHORTAGE: list[dict] = []
_LAST_INTERACTIVE_TRIAL_METERS_DONE_SNAPSHOT: dict[tuple[str, str, str, date], float] = {}
_LAST_INTERACTIVE_TRIAL_PLAN_TARGETS_SNAPSHOT: dict[tuple[str, str, str, date], float] = {}
_LAST_INTERACTIVE_TRIAL_META_MISS_SHORTFALL: list[dict] = []
_LAST_INTERACTIVE_REMAINING_TASKS_AT_CALENDAR_END: list[dict] = []
_LAST_INTERACTIVE_STAGE3_META: dict = {}
_PLAN_IMPL_INTERACTIVE_DISPATCH_TARGETS: dict | None = None
_PLAN_IMPL_INTERACTIVE_TRIAL_METERS_DONE: dict | None = None
def interactive_stage3_last_run_meta_snapshot() -> dict:
    """直近の配台試行メタ（例: mode=single_phase）。試行未実行時は空 dict。"""
    return dict(_LAST_INTERACTIVE_STAGE3_META or {})
def interactive_trial_meta_miss_shortfall_snapshot() -> list[dict]:
    """直近試行で加工開始日時が空の計画行（タイムライン未割付）。"""
    return list(_LAST_INTERACTIVE_TRIAL_META_MISS_SHORTFALL or [])
def interactive_trial_plan_targets_snapshot() -> dict[tuple[str, str, str, date], float]:
    """直近試行の配台日スライド後 plan（当日配台数量）キー。未設定時は空 dict。"""
    return dict(_LAST_INTERACTIVE_TRIAL_PLAN_TARGETS_SNAPSHOT or {})
def interactive_trial_remaining_tasks_at_calendar_end_snapshot() -> list[dict]:
    """直近試行終了時、勤怠最終日までに割り切れなかったタスク行。"""
    return list(_LAST_INTERACTIVE_REMAINING_TASKS_AT_CALENDAR_END or [])
PLAN_DUE_DAY_COMPLETION_TIME = time(16, 0)
_ai_remarks_cache_name = "ai_remarks_cache.json"
_ai_cache_legacy = os.path.join(output_dir, _ai_remarks_cache_name)
_ai_cache_new = os.path.join(json_data_dir, _ai_remarks_cache_name)
if os.path.isfile(_ai_cache_legacy) and not os.path.isfile(_ai_cache_new):
    try:
        shutil.move(_ai_cache_legacy, _ai_cache_new)
    except OSError:
        pass
ai_cache_path = _ai_cache_new
exclude_rules_sheet_debug_log_path = os.path.join(log_dir, "exclude_rules_sheet_debug.txt")
SUMMARY_AI_DISPATCH_XLSX = "サマリ_AI配台.xlsx"
GEMINI_USAGE_SUMMARY_FOR_MAIN_FILE = "gemini_usage_summary_for_main.txt"
GEMINI_USAGE_CUMULATIVE_JSON_FILE = "gemini_usage_cumulative.json"
GEMINI_USAGE_BUCKETS_CSV_FILE = "gemini_usage_buckets_for_chart.csv"
GEMINI_USAGE_CHART_COL_DATE = 17  # Q
GEMINI_USAGE_CHART_COL_VALUE = 18  # R
GEMINI_USAGE_CHART_COL_TOK_DATE = 19  # S（グラフ用に日付を複製）
GEMINI_USAGE_CHART_COL_TOK_VALUE = 20  # T（total_tokens 相当）
GEMINI_USAGE_CHART_HEADER_ROW = 16
GEMINI_USAGE_CHART_ANCHOR_CELL = "T16"
GEMINI_USAGE_CHART_TOKENS_ANCHOR_CELL = "AA16"
GEMINI_USAGE_CHART_MAX_DAYS = 14
GEMINI_USAGE_CHART_CLEAR_ROWS = 36
GEMINI_USAGE_XLW_CHART_NAME = "_GeminiApiDailyTrend"
GEMINI_USAGE_XLW_CHART_TOKENS_NAME = "_GeminiApiDailyTokens"
def interactive_trial_shortages_snapshot() -> dict:
    snap = {
        "op_shortage": list(_INTERACTIVE_TRIAL_OP_SHORTAGE),
        "as_shortage": list(_INTERACTIVE_TRIAL_AS_SHORTAGE),
    }
    if _LAST_INTERACTIVE_TRIAL_METERS_DONE_SNAPSHOT:
        snap = filter_interactive_trial_shortages_by_meters_done(
            snap, _LAST_INTERACTIVE_TRIAL_METERS_DONE_SNAPSHOT
        )
    return {
        "op_shortage": _dedupe_interactive_trial_shortage_records(
            snap.get("op_shortage")
        ),
        "as_shortage": _dedupe_interactive_trial_shortage_records(
            snap.get("as_shortage")
        ),
    }
def interactive_trial_meters_done_snapshot() -> dict[tuple[str, str, str, date], float]:
    """配台試行でタイムラインが記録した暦日別メートル達成（試行終了直後のスナップショット）。"""
    return dict(_LAST_INTERACTIVE_TRIAL_METERS_DONE_SNAPSHOT)

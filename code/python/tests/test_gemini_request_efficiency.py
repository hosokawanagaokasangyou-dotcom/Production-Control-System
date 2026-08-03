# -*- coding: utf-8 -*-
"""Gemini 呼び出しの応答時間対策（思考無効・構造化出力・バッチ並列・モデル列）の契約テスト。

段階2は 1 リクエストで数万トークンを生成させていたため 10 分以上かかっていた。
出力量そのものを削る仕組みが壊れていないことをここで固定する。
"""

from __future__ import annotations

import json
import threading
import time

import pandas as pd
import pytest

from planning_core import _core as pc


# --------------------------------------------------------------------------
# フォールバックモデル列
# --------------------------------------------------------------------------


def test_default_model_chain_excludes_zero_free_tier_generations():
    for mid in pc.GEMINI_MODEL_IDS_BY_QUALITY:
        assert not mid.startswith("gemini-1."), mid
        assert not mid.startswith("gemini-2."), mid
        assert "pro" not in mid, mid


def test_default_model_chain_starts_with_top_priority_model():
    assert pc.GEMINI_MODEL_IDS_BY_QUALITY[0] == pc.GEMINI_MODEL_FLASH


@pytest.mark.parametrize(
    "model_id",
    ["gemini-3.5-flash", "models/gemini-3.1-flash-lite", "gemini-flash-latest"],
)
def test_model_has_free_tier_allocation_accepts_current_generations(model_id):
    assert pc._gemini_model_has_free_tier_allocation(model_id)


@pytest.mark.parametrize(
    "model_id",
    [
        "gemini-2.5-flash-lite",
        "gemini-2.0-flash-lite",
        "models/gemini-2.5-flash",
        "gemini-3.5-pro",
        "",
    ],
)
def test_model_has_free_tier_allocation_rejects_exhausted_generations(model_id):
    assert not pc._gemini_model_has_free_tier_allocation(model_id)


# --------------------------------------------------------------------------
# 生成設定（思考無効・JSON 構造化出力）
# --------------------------------------------------------------------------


def test_thinking_budget_defaults_to_zero(monkeypatch):
    monkeypatch.delenv("GEMINI_THINKING_BUDGET", raising=False)
    assert pc._gemini_thinking_budget() == 0


@pytest.mark.parametrize(("raw", "expected"), [("512", 512), ("-1", -1), ("0", 0)])
def test_thinking_budget_honors_env(monkeypatch, raw, expected):
    monkeypatch.setenv("GEMINI_THINKING_BUDGET", raw)
    assert pc._gemini_thinking_budget() == expected


def test_thinking_budget_falls_back_to_zero_on_garbage(monkeypatch):
    monkeypatch.setenv("GEMINI_THINKING_BUDGET", "あいうえお")
    assert pc._gemini_thinking_budget() == 0


def test_generate_content_config_disables_thinking_by_default(monkeypatch):
    monkeypatch.delenv("GEMINI_THINKING_BUDGET", raising=False)
    cfg = pc._gemini_generate_content_config()
    assert cfg is not None
    assert cfg.thinking_config.thinking_budget == 0


def test_generate_content_config_omits_thinking_when_budget_is_dynamic(monkeypatch):
    monkeypatch.setenv("GEMINI_THINKING_BUDGET", "-1")
    cfg = pc._gemini_generate_content_config()
    assert cfg.thinking_config is None


def test_generate_content_config_without_schema_does_not_force_json(monkeypatch):
    monkeypatch.delenv("GEMINI_THINKING_BUDGET", raising=False)
    cfg = pc._gemini_generate_content_config()
    assert cfg.response_mime_type is None
    assert cfg.response_schema is None


def test_generate_content_config_enables_json_mode_with_schema():
    schema = {"type": "OBJECT", "properties": {"a": {"type": "STRING"}}}
    cfg = pc._gemini_generate_content_config(response_schema=schema)
    assert cfg.response_mime_type == "application/json"
    assert cfg.response_schema == schema


def test_generate_content_config_sets_max_output_tokens():
    cfg = pc._gemini_generate_content_config(max_output_tokens=4096)
    assert cfg.max_output_tokens == 4096


# --------------------------------------------------------------------------
# JSON 応答パース
# --------------------------------------------------------------------------


@pytest.mark.parametrize(
    "text",
    [
        '{"a": 1}',
        '```json\n{"a": 1}\n```',
        'ここに説明\n```\n{"a": 1}\n```\n以上',
    ],
)
def test_parse_json_object_tolerates_fences_and_prose(text):
    assert pc._gemini_parse_json_object(text) == {"a": 1}


@pytest.mark.parametrize("text", ["", None, "JSON はありません", "{壊れた"])
def test_parse_json_object_returns_none_for_unusable_text(text):
    assert pc._gemini_parse_json_object(text) is None


# --------------------------------------------------------------------------
# テスト用フェイククライアント
# --------------------------------------------------------------------------


class _FakeUsage:
    prompt_token_count = 10
    candidates_token_count = 20
    total_token_count = 30
    thoughts_token_count = 0


class _FakeResponse:
    def __init__(self, text: str) -> None:
        self.text = text
        self.usage_metadata = _FakeUsage()


class _FakeModels:
    def __init__(self, handler) -> None:
        self._handler = handler
        self.calls: list[dict] = []
        self._lock = threading.Lock()

    def generate_content(self, *, model, contents, config=None):
        with self._lock:
            self.calls.append({"model": model, "contents": contents, "config": config})
        return self._handler(model, contents, config)


class _FakeClient:
    def __init__(self, handler) -> None:
        self.models = _FakeModels(handler)


@pytest.fixture
def fast_gemini(monkeypatch):
    """ジッター待機・レート制限・課金集計の副作用を止める（フェイク API 相手のため）。"""
    monkeypatch.setattr(pc, "_GEMINI_PRE_REQUEST_JITTER_MAX", 0.0)
    monkeypatch.setattr(pc, "_gemini_acquire_request_slot", lambda prefix="": 0.0)
    monkeypatch.setattr(pc, "record_gemini_response_usage", lambda *a, **k: None)
    monkeypatch.setattr(pc, "_gemini_progress_log_interval_sec", lambda: 0.0)
    monkeypatch.delenv("PM_AI_SKIP_GEMINI_API", raising=False)
    monkeypatch.delenv("GEMINI_THINKING_BUDGET", raising=False)


# --------------------------------------------------------------------------
# 再試行ラッパが生成設定を渡すこと
# --------------------------------------------------------------------------


def test_retry_wrapper_passes_thinking_disabled_config(fast_gemini):
    client = _FakeClient(lambda m, c, cfg: _FakeResponse("{}"))
    _res, model_id = pc._gemini_generate_content_with_retry(
        client, contents="x", model="gemini-3.5-flash"
    )
    assert model_id == "gemini-3.5-flash"
    cfg = client.models.calls[0]["config"]
    assert cfg is not None
    assert cfg.thinking_config.thinking_budget == 0


def test_retry_wrapper_passes_caller_supplied_config(fast_gemini):
    client = _FakeClient(lambda m, c, cfg: _FakeResponse("{}"))
    schema = {"type": "OBJECT", "properties": {"a": {"type": "STRING"}}}
    supplied = pc._gemini_generate_content_config(response_schema=schema)
    pc._gemini_generate_content_with_retry(
        client, contents="x", model="gemini-3.5-flash", config=supplied
    )
    assert client.models.calls[0]["config"] is supplied


def test_retry_wrapper_retries_without_thinking_config_when_unsupported(fast_gemini):
    seen: list[object] = []

    def handler(model, contents, config):
        seen.append(getattr(config, "thinking_config", None))
        if len(seen) == 1:
            raise RuntimeError(
                "400 INVALID_ARGUMENT: Budgeted thinking is not supported for this model"
            )
        return _FakeResponse("{}")

    client = _FakeClient(handler)
    pc._gemini_generate_content_with_retry(
        client, contents="x", model="gemini-3.5-flash"
    )
    assert len(seen) == 2
    assert seen[0] is not None
    assert seen[1] is None


# --------------------------------------------------------------------------
# バッチ分割・並列実行
# --------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("total", "size", "expected"),
    [
        (0, 10, []),
        (5, 10, [(0, 5)]),
        (10, 10, [(0, 10)]),
        (25, 10, [(0, 10), (10, 20), (20, 25)]),
        (3, 1, [(0, 1), (1, 2), (2, 3)]),
    ],
)
def test_batch_slices(total, size, expected):
    assert pc._gemini_batch_slices(total, size) == expected


def test_batch_slices_treats_non_positive_size_as_single_batch():
    assert pc._gemini_batch_slices(7, 0) == [(0, 7)]


def test_generate_json_map_in_batches_merges_every_batch(fast_gemini):
    items = [f"item{i}" for i in range(25)]

    def handler(model, contents, config):
        keys = [ln for ln in contents.splitlines() if ln.startswith("item")]
        return _FakeResponse(json.dumps({k: {"n": 1} for k in keys}))

    client = _FakeClient(handler)
    merged, model_ids, failed = pc._gemini_generate_json_map_in_batches(
        client,
        items=items,
        build_prompt=lambda chunk: "\n".join(chunk),
        batch_size=10,
        max_workers=3,
        log_label="テスト",
    )

    assert failed == 0
    assert len(client.models.calls) == 3
    assert set(merged) == set(items)
    assert model_ids == ["gemini-3.5-flash"] * 3


def test_generate_json_map_in_batches_runs_batches_concurrently(fast_gemini):
    lock = threading.Lock()
    state = {"active": 0, "peak": 0}

    def handler(model, contents, config):
        with lock:
            state["active"] += 1
            state["peak"] = max(state["peak"], state["active"])
        time.sleep(0.2)
        with lock:
            state["active"] -= 1
        return _FakeResponse("{}")

    client = _FakeClient(handler)
    pc._gemini_generate_json_map_in_batches(
        client,
        items=list(range(12)),
        build_prompt=lambda chunk: "x",
        batch_size=3,
        max_workers=4,
    )
    assert state["peak"] >= 2


def test_generate_json_map_in_batches_keeps_successful_batches(fast_gemini):
    def handler(model, contents, config):
        if "item3" in contents:
            raise RuntimeError("400 INVALID_ARGUMENT: bad batch")
        return _FakeResponse(json.dumps({contents: {"ok": True}}))

    client = _FakeClient(handler)
    merged, _model_ids, failed = pc._gemini_generate_json_map_in_batches(
        client,
        items=[f"item{i}" for i in range(4)],
        build_prompt=lambda chunk: chunk[0],
        batch_size=1,
        max_workers=2,
    )

    assert failed == 1
    assert set(merged) == {"item0", "item1", "item2"}


def test_generate_json_map_in_batches_returns_empty_for_no_items(fast_gemini):
    client = _FakeClient(lambda m, c, cfg: _FakeResponse("{}"))
    merged, model_ids, failed = pc._gemini_generate_json_map_in_batches(
        client, items=[], build_prompt=lambda chunk: "x"
    )
    assert merged == {}
    assert model_ids == []
    assert failed == 0
    assert client.models.calls == []


def test_batch_tuning_defaults(monkeypatch):
    monkeypatch.delenv("GEMINI_BATCH_MAX_ITEMS", raising=False)
    monkeypatch.delenv("GEMINI_MAX_PARALLEL_REQUESTS", raising=False)
    assert pc._gemini_batch_max_items() > 0
    workers = pc._gemini_max_parallel_requests()
    assert 1 <= workers <= 15


def test_batch_tuning_env_overrides(monkeypatch):
    monkeypatch.setenv("GEMINI_BATCH_MAX_ITEMS", "40")
    monkeypatch.setenv("GEMINI_MAX_PARALLEL_REQUESTS", "6")
    assert pc._gemini_batch_max_items() == 40
    assert pc._gemini_max_parallel_requests() == 6


def test_batch_tuning_clamps_parallel_requests_to_free_tier_rpm(monkeypatch):
    monkeypatch.setenv("GEMINI_MAX_PARALLEL_REQUESTS", "999")
    assert pc._gemini_max_parallel_requests() <= 15


# --------------------------------------------------------------------------
# 勤怠備考 AI（差分のみ出力）
# --------------------------------------------------------------------------


def test_attendance_ai_schema_version_marks_diff_only_contract():
    assert pc.ATTENDANCE_REMARK_AI_SCHEMA_ID == "v3_sabun_batch"


def test_attendance_ai_response_schema_requires_only_the_key():
    schema = pc.ATTENDANCE_REMARK_AI_RESPONSE_SCHEMA
    item = schema["properties"]["entries"]["items"]
    assert item["required"] == [pc.ATTENDANCE_AI_ENTRY_KEY_FIELD]
    assert pc.ATTENDANCE_AI_ENTRY_KEY_FIELD in item["properties"]


def test_attendance_ai_prompt_carries_lines_and_demands_diff_only():
    line = "2026-04-01_山田 の備考: 午後半休"
    prompt = pc._attendance_remark_ai_prompt([line])
    assert line in prompt
    assert "変更がある" in prompt
    assert "省略" in prompt


def test_attendance_ai_entries_to_map_reads_entries_array():
    payload = {
        "entries": [
            {pc.ATTENDANCE_AI_ENTRY_KEY_FIELD: "2026-04-01_山田", "is_holiday": True},
            {pc.ATTENDANCE_AI_ENTRY_KEY_FIELD: "2026-04-02_鈴木", "作業効率": 0.5},
        ]
    }
    assert pc._attendance_ai_entries_to_map(payload) == {
        "2026-04-01_山田": {"is_holiday": True},
        "2026-04-02_鈴木": {"作業効率": 0.5},
    }


def test_attendance_ai_entries_to_map_accepts_legacy_object_map():
    payload = {"2026-04-01_山田": {"is_holiday": True}}
    assert pc._attendance_ai_entries_to_map(payload) == {
        "2026-04-01_山田": {"is_holiday": True}
    }


@pytest.mark.parametrize("payload", [None, {}, {"entries": []}, {"entries": "x"}, []])
def test_attendance_ai_entries_to_map_returns_empty_for_unusable_payload(payload):
    assert pc._attendance_ai_entries_to_map(payload) == {}


def test_attendance_ai_entries_to_map_drops_entries_without_key():
    payload = {"entries": [{"is_holiday": True}, {pc.ATTENDANCE_AI_ENTRY_KEY_FIELD: ""}]}
    assert pc._attendance_ai_entries_to_map(payload) == {}


# --------------------------------------------------------------------------
# 差分のみ出力にしても「休日シフト」判定が壊れないこと（回帰防止）
# --------------------------------------------------------------------------


def test_empty_shift_is_false_when_row_was_sent_to_ai():
    """AI が変更なしとして何も返さなくても、解析対象だった行を休日にしてはいけない。"""
    row = pd.Series({"出勤時間": pd.NA, "退勤時間": pd.NA})
    key = "2026-04-01_山田"
    assert pc._attendance_is_empty_shift(row, key=key, analyzed_keys={key}) is False


def test_empty_shift_is_true_when_row_was_never_analyzed():
    row = pd.Series({"出勤時間": pd.NA, "退勤時間": pd.NA})
    assert pc._attendance_is_empty_shift(row, key="2026-04-01_山田", analyzed_keys=set()) is True


@pytest.mark.parametrize(
    ("start", "end"),
    [("08:45", "17:30"), ("08:45", pd.NA), (pd.NA, "17:30")],
)
def test_empty_shift_is_false_when_any_shift_time_exists(start, end):
    row = pd.Series({"出勤時間": start, "退勤時間": end})
    assert pc._attendance_is_empty_shift(row, key="k", analyzed_keys=set()) is False


# --------------------------------------------------------------------------
# thinkingBudget を拒むモデル（gemini-3.5-flash-lite は理由なしの 400 を返す）
# --------------------------------------------------------------------------


_GENERIC_INVALID_ARGUMENT = (
    "400 INVALID_ARGUMENT. {'error': {'code': 400, "
    "'message': 'Request contains an invalid argument.', 'status': 'INVALID_ARGUMENT'}}"
)


@pytest.mark.parametrize(
    "err_text",
    [
        _GENERIC_INVALID_ARGUMENT,
        "400 INVALID_ARGUMENT thinking_config is not supported",
        "Thinking is unsupported for this model",
    ],
)
def test_thinking_rejection_covers_generic_invalid_argument(err_text):
    assert pc._gemini_is_thinking_config_unsupported_error(err_text) is True


@pytest.mark.parametrize(
    "err_text",
    [
        "429 RESOURCE_EXHAUSTED quota exceeded",
        "503 UNAVAILABLE model is overloaded",
        "404 NOT_FOUND models/foo is not found",
        "",
    ],
)
def test_thinking_rejection_ignores_unrelated_errors(err_text):
    assert pc._gemini_is_thinking_config_unsupported_error(err_text) is False


class _RejectsThinking:
    """thinking_config 付きの要求だけ 400 を返すフェイク。"""

    def __init__(self) -> None:
        self.models = self
        self.configs: list[object] = []

    def generate_content(self, *, model, contents, config=None):
        self.configs.append(config)
        if config is not None and getattr(config, "thinking_config", None) is not None:
            raise RuntimeError(_GENERIC_INVALID_ARGUMENT)
        return _FakeResponse("{}")


def test_invoke_drops_thinking_config_on_generic_invalid_argument():
    pc._gemini_forget_thinking_config_rejections()
    client = _RejectsThinking()
    config = pc._gemini_generate_content_config(thinking_budget=0)
    _res, used = pc._gemini_invoke_generate_content(
        client, "gemini-3.5-flash-lite", "x", config
    )
    assert getattr(used, "thinking_config", None) is None
    assert len(client.configs) == 2


def test_invoke_remembers_models_that_reject_thinking_config():
    pc._gemini_forget_thinking_config_rejections()
    client = _RejectsThinking()
    config = pc._gemini_generate_content_config(thinking_budget=0)
    pc._gemini_invoke_generate_content(client, "gemini-3.5-flash-lite", "x", config)
    pc._gemini_invoke_generate_content(client, "gemini-3.5-flash-lite", "y", config)
    # 2 回目は最初から思考設定を外して送るので、余計な 400 を踏まない。
    assert len(client.configs) == 3
    assert getattr(client.configs[2], "thinking_config", None) is None


def test_thinking_config_rejection_is_remembered_per_model():
    pc._gemini_forget_thinking_config_rejections()
    client = _RejectsThinking()
    config = pc._gemini_generate_content_config(thinking_budget=0)
    pc._gemini_invoke_generate_content(client, "gemini-3.5-flash-lite", "x", config)
    pc._gemini_invoke_generate_content(client, "gemini-3.5-flash", "y", config)
    # 別モデルでは記憶を流用せず、思考設定つきで一度試す。
    assert getattr(client.configs[2], "thinking_config", None) is not None


# --------------------------------------------------------------------------
# 送信レート制限（無料枠 RPM 超過による 429 の予防）
# --------------------------------------------------------------------------


def test_requests_per_minute_defaults_to_free_tier_limit(monkeypatch):
    monkeypatch.delenv("GEMINI_REQUESTS_PER_MINUTE", raising=False)
    assert pc._gemini_requests_per_minute() == pc.GEMINI_FREE_TIER_RPM_LIMIT


@pytest.mark.parametrize(("raw", "expected"), [("5", 5), ("1", 1), ("0", 1), ("x", 15)])
def test_requests_per_minute_honors_env(monkeypatch, raw, expected):
    monkeypatch.setenv("GEMINI_REQUESTS_PER_MINUTE", raw)
    assert pc._gemini_requests_per_minute() == expected


def test_rate_limiter_lets_a_burst_through_up_to_the_limit():
    limiter = pc._GeminiRateLimiter(limit=3, window_sec=60.0)
    t0 = time.perf_counter()
    for _ in range(3):
        limiter.acquire()
    assert time.perf_counter() - t0 < 0.5


def test_rate_limiter_blocks_once_the_window_is_full():
    limiter = pc._GeminiRateLimiter(limit=2, window_sec=0.4)
    limiter.acquire()
    limiter.acquire()
    t0 = time.perf_counter()
    waited = limiter.acquire()
    elapsed = time.perf_counter() - t0
    assert elapsed >= 0.2
    assert waited > 0


def test_rate_limiter_frees_slots_after_the_window_passes():
    limiter = pc._GeminiRateLimiter(limit=1, window_sec=0.2)
    limiter.acquire()
    time.sleep(0.25)
    t0 = time.perf_counter()
    limiter.acquire()
    assert time.perf_counter() - t0 < 0.1


def test_rate_limiter_never_exceeds_the_limit_under_threads():
    limiter = pc._GeminiRateLimiter(limit=4, window_sec=0.5)
    started = threading.Barrier(8)
    stamps: list[float] = []
    lock = threading.Lock()

    def worker():
        started.wait()
        limiter.acquire()
        with lock:
            stamps.append(time.perf_counter())

    threads = [threading.Thread(target=worker) for _ in range(8)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=10)
    assert len(stamps) == 8
    stamps.sort()
    # 先頭 4 件は即時、残り 4 件はウィンドウが空くまで待たされる。
    assert stamps[4] - stamps[0] >= 0.4


def test_retry_wrapper_takes_a_rate_limit_slot_per_attempt(fast_gemini, monkeypatch):
    taken: list[str] = []
    monkeypatch.setattr(
        pc, "_gemini_acquire_request_slot", lambda prefix="": taken.append(prefix) or 0.0
    )
    client = _FakeClient(lambda m, c, cfg: _FakeResponse("{}"))
    pc._gemini_generate_content_with_retry(
        client, contents="x", model="gemini-3.5-flash", log_label="レート"
    )
    assert taken == ["レート: "]


# --------------------------------------------------------------------------
# タスク特別指定 AI（差分のみ出力）
# --------------------------------------------------------------------------


_TASK_SPECIAL_BLOB = "依頼NO【Y4-2】 工程名「エンボス」 機械名「E1」 備考: 4/5までに終わらせる"


def test_task_special_prompt_carries_blob_and_reference_year():
    prompt = pc._task_special_ai_prompt(_TASK_SPECIAL_BLOB, 2026)
    assert _TASK_SPECIAL_BLOB in prompt
    assert "2026" in prompt


def test_task_special_prompt_demands_diff_only_output():
    prompt = pc._task_special_ai_prompt(_TASK_SPECIAL_BLOB, 2026)
    assert "読み取れる項目だけ" in prompt
    assert "省略" in prompt


def test_task_special_prompt_makes_source_row_labels_optional():
    """process_name / machine_name は照合時に空を許容するため、必須にして出力を膨らませない。"""
    prompt = pc._task_special_ai_prompt(_TASK_SPECIAL_BLOB, 2026)
    assert "必須" not in prompt


def test_task_special_prompt_forbids_prose_and_fences():
    prompt = pc._task_special_ai_prompt(_TASK_SPECIAL_BLOB, 2026)
    assert "コードフェンス" in prompt

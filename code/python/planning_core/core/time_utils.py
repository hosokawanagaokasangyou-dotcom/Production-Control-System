# -*- coding: utf-8 -*-
# planning_core.core.time_utils — body only (loaded via _core exec chain)
import bisect
def _eod_minutes_window_covers_start(
    team_start: datetime, team_end_limit: datetime
) -> bool:
    """ASSIGN_END_OF_DAY_DEFER_MINUTES は正のとき」開始は終業上限のしの分数以内か。"""
    gap = ASSIGN_END_OF_DAY_DEFER_MINUTES
    if gap <= 0:
        return False
    if team_start >= team_end_limit:
        return False
    return (team_end_limit - team_start) <= timedelta(minutes=gap)
def _eod_same_request_continuation_exempt(
    machine_occ_key: str, task: dict, machine_handoff: dict | None
) -> bool:
    """
    同一設備占有キーで直前に載せた加工が同一依頼NO（task_id）のとき True。
    終業直前デファーは「新規開始」に寄せるため、この場合は小残・収容閾値の EOD 抑止を外す。
    """
    if not machine_handoff:
        return False
    occ = str(machine_occ_key or "").strip()
    if not occ:
        return False
    prev = (machine_handoff.get("last_tid") or {}).get(occ)
    cur = str(task.get("task_id") or "").strip()
    if not prev or not cur:
        return False
    return str(prev).strip() == cur
def _eod_reject_capacity_units_below_threshold(
    units_fit_until_close: int,
    team_start: datetime,
    team_end_limit: datetime,
    *,
    eod_same_request_continuation_exempt: bool = False,
    remaining_units_ceil: int | None = None,
) -> bool:
    """
    終業まであと ASSIGN_END_OF_DAY_DEFER_MINUTES 分以内のウィンドウ内で、
    必要収容ロール数ロール分以上は回せない（収容ロール数が閾値未満）とき True（新規加工を始めない＝候補却下）。
    必要収容ロール数 = min(ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS, remaining_units_ceil)
    （remaining_units_ceil が正のとき。未指定時は従来どおり ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS のみ。）
    eod_same_request_continuation_exempt が True のときは常に False（同一依頼の連続ロール）。
    """
    if eod_same_request_continuation_exempt:
        return False
    th = ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS
    if th <= 0:
        return False
    if not _eod_minutes_window_covers_start(team_start, team_end_limit):
        return False
    eff_th = int(th)
    if remaining_units_ceil is not None and int(remaining_units_ceil) > 0:
        eff_th = min(eff_th, int(remaining_units_ceil))
    return int(units_fit_until_close) < int(eff_th)
def merge_time_intervals(intervals):
    """時刻区間のリストをソートし、重なる区間を統合して返す。"""
    if not intervals:
        return []
    intervals.sort(key=lambda x: x[0])
    merged = [intervals[0]]
    for current_start, current_end in intervals[1:]:
        last_start, last_end = merged[-1]
        if current_start <= last_end:
            merged[-1] = (last_start, max(last_end, current_end))
        else:
            merged.append((current_start, current_end))
    return merged
def _contiguous_work_minutes_until_next_break_or_limit(
    start_dt: datetime,
    breaks_dt: list,
    end_limit_dt: datetime,
) -> int:
    """
    start_dt から次の休憩開始（または終業上限）までの」連続して実僝に使うる分数。
    開始は休憩帯内なら 0（呼び出し元で坴下）。breaks_dt は merge 済み想定。
    """
    if start_dt >= end_limit_dt:
        return 0
    for bs, be in breaks_dt:
        if bs <= start_dt < be:
            return 0
    next_stop = end_limit_dt
    for bs, be in breaks_dt:
        if be <= start_dt:
            continue
        if start_dt < bs:
            next_stop = min(next_stop, bs)
    return max(0, int((next_stop - start_dt).total_seconds() / 60))
def _break_end_to_skip_if_contiguous_under(
    start_dt: datetime,
    breaks_dt: list,
    end_limit_dt: datetime,
    min_contiguous_mins: int,
) -> datetime | None:
    """
    休憩帯外でも」次の休憩開始までの連続実僝は min_contiguous_mins 未満なら」
    しの休憩区間の終了時刻を返す（午後休憩直後に 1 ロール分は坎まらない開始の値進ゝる）。
    終業までしか実僝は続かない場合は None。
    """
    if min_contiguous_mins <= 0:
        return None
    if start_dt >= end_limit_dt:
        return None
    c = _contiguous_work_minutes_until_next_break_or_limit(
        start_dt, breaks_dt, end_limit_dt
    )
    if c >= min_contiguous_mins:
        return None
    next_stop = end_limit_dt
    for bs, be in breaks_dt:
        if be <= start_dt:
            continue
        if start_dt < bs:
            next_stop = min(next_stop, bs)
    if next_stop >= end_limit_dt:
        return None
    for bs, be in breaks_dt:
        if bs == next_stop:
            return be
    return None
def _defer_team_start_past_prebreak_and_end_of_day(
    task: dict,
    team: tuple,
    team_start: datetime,
    team_end_limit: datetime,
    team_breaks: list,
    refloor_fn,
    min_contiguous_work_mins: int | None = None,
    *,
    eod_same_request_continuation_exempt: bool = False,
) -> datetime | None:
    """
    - ASSIGN_END_OF_DAY_DEFER_MINUTES > 0 かつ (team_end_limit - 試行開始) がその分数以下で、
      remaining_units 切り上げが ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS 以下のとき、当日開始不可（None）。
      eod_same_request_continuation_exempt が True のときはこの終業直前・小残分岐をスキップ（同一依頼の連続ロール）。
    - 試行開始が休憩帯内のときは **休憩終了時刻へ繰り下げ**し、`refloor_fn` で設備下限・avail を再適用する。
      繰り下げのあと終業超過・EOD デファーに該当すれば None。
    - min_contiguous_work_mins が正のとき、帯外でも **次の休憩までの連続実働**がそれ未満なら
      当該休憩の終了へ繰り下げ（上と同様に refloor しループ）。
    """
    _tid = str(task.get("task_id", "") or "").strip()
    _team_txt = ", ".join(str(x) for x in team) if team else "—"

    def _trace_block(msg: str, *a) -> None:
        if not _trace_schedule_task_enabled(_tid):
            return
        _log_dispatch_trace_schedule(
            _tid,
            "[配台トレース task=%s] ブロック判定: " + msg,
            _tid,
            *a,
        )

    ts = refloor_fn(team_start)
    for _ in range(64):
        if ts >= team_end_limit:
            _trace_block(
                "開始試行(終業超靎) machine=%s team=%s rem=%.4f trial_start=%s end_limit=%s",
                task.get("machine"),
                _team_txt,
                float(task.get("remaining_units") or 0),
                ts,
                team_end_limit,
            )
            return None

        break_end = None
        for bs, be in team_breaks:
            if bs <= ts < be:
                break_end = be
                break
        if break_end is not None:
            _trace_block(
                "休憩帯内のため、終了へ繰り下き machine=%s team=%s rem=%.4f break_end=%s trial_was=%s",
                task.get("machine"),
                _team_txt,
                float(task.get("remaining_units") or 0),
                break_end,
                ts,
            )
            ts = refloor_fn(break_end)
            continue

        if min_contiguous_work_mins is not None and min_contiguous_work_mins > 0:
            slip_end = _break_end_to_skip_if_contiguous_under(
                ts, team_breaks, team_end_limit, min_contiguous_work_mins
            )
            if slip_end is not None:
                _trace_block(
                    "休憩直後で連続実僝丝足のため、休憩終了へ繰り下き machine=%s team=%s rem=%.4f need_contig_min=%s trial_was=%s break_end=%s",
                    task.get("machine"),
                    _team_txt,
                    float(task.get("remaining_units") or 0),
                    min_contiguous_work_mins,
                    ts,
                    slip_end,
                )
                ts = refloor_fn(slip_end)
                continue

        gap_end = ASSIGN_END_OF_DAY_DEFER_MINUTES
        rem_ceil = math.ceil(float(task.get("remaining_units") or 0))
        if (
            not eod_same_request_continuation_exempt
            and not _stage3_qty_strict_active()
            and gap_end > 0
            and (team_end_limit - ts) <= timedelta(minutes=gap_end)
            and rem_ceil <= ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS
        ):
            _trace_block(
                "開始試行(終業直後・尝残ロール) machine=%s team=%s rem_ceil=%s max_rem=%s trial_start=%s end_limit=%s gap_end_min=%s",
                task.get("machine"),
                _team_txt,
                rem_ceil,
                ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS,
                ts,
                team_end_limit,
                gap_end,
            )
            return None

        return ts

    _trace_block(
        "開始試行(休憩繰り下き打切り) machine=%s team=%s rem=%.4f trial_start=%s",
        task.get("machine"),
        _team_txt,
        float(task.get("remaining_units") or 0),
        ts,
    )
    return None
def _expand_timeline_events_for_equipment_grid(timeline_events: list) -> list:
    """
    設備毎の時間割・メンバー日程・稼働率用インデックス坑け。
    1 本のイベントは日をまたし場合」e["date"] の値当日に載せると翌朝セグメントは欠けるため、
    start_dt〜end_dt を坄就業日 DEFAULT_START_TIME〜DEFAULT_END_TIME にクリップした複製へ展開れる。
    """
    expanded: list = []
    for e in timeline_events:
        sd = e.get("start_dt")
        ed = e.get("end_dt")
        if not isinstance(sd, datetime) or not isinstance(ed, datetime):
            expanded.append(e)
            continue
        if ed <= sd:
            expanded.append(e)
            continue
        segments: list = []
        cal = sd.date()
        last_d = ed.date()
        while cal <= last_d:
            w0 = datetime.combine(cal, DEFAULT_START_TIME)
            w1 = datetime.combine(cal, DEFAULT_END_TIME)
            a = max(sd, w0)
            b = min(ed, w1)
            if a < b:
                ne = dict(e)
                ne["date"] = cal
                ne["start_dt"] = a
                ne["end_dt"] = b
                segments.append(ne)
            cal += timedelta(days=1)
        if segments:
            expanded.extend(segments)
        else:
            expanded.append(e)
    return expanded
def get_actual_work_minutes(start_dt, end_dt, breaks_dt):
    """
    start_dt から end_dt までの「休憩を除いた実僝分数」。
    breaks_dt … (区間開始, 区間終了) の列（datetime または time。呼び出し元の勤怠イベントと整合）。
    """
    current = start_dt
    actual_mins = 0
    while current < end_dt:
        next_event = end_dt
        in_break = False
        b_end_time = None
        for b_s, b_e in breaks_dt:
            if b_s <= current < b_e:
                in_break = True
                b_end_time = b_e
                break
            elif current < b_s < next_event:
                next_event = b_s
        
        if in_break:
            current = b_end_time
        else:
            actual_mins += int((next_event - current).total_seconds() / 60)
            current = next_event
    return actual_mins
def calculate_end_time(start_dt, duration_minutes, breaks_dt, end_limit_dt):
    """
    start_dt から実僝 duration_minutes 分進ゝた終了 datetime を求ゝる（休憩はスキップ）。
    end_limit_dt を超ごないよご打ち切り。戻り値: (終了時刻, 実際に進ゝた実僝分, 残り未消化分)
    """
    current = start_dt
    remaining_work = duration_minutes
    actual_work_time = 0 

    while current < end_limit_dt and remaining_work > 0:
        next_event = end_limit_dt
        in_break = False
        break_end = None
        for b_start, b_end in breaks_dt:
            if b_start <= current < b_end:
                in_break = True
                break_end = b_end
                break
            elif current < b_start < next_event:
                next_event = b_start
                
        if in_break:
            current = break_end
            continue
            
        block_mins = int((next_event - current).total_seconds() / 60)
        if remaining_work <= block_mins:
            actual_work_time += remaining_work
            current += timedelta(minutes=remaining_work)
            remaining_work = 0
        else:
            actual_work_time += block_mins
            remaining_work -= block_mins
            current = next_event

    end_dt = min(current, end_limit_dt)
    return end_dt, actual_work_time, remaining_work
def _dt_close_minutes(a: datetime, b: datetime, tol_sec: int = 59) -> bool:
    return abs((a - b).total_seconds()) <= tol_sec
def _find_latest_prep_start_matching_end(
    end_at: datetime,
    dur_mins: int,
    breaks_merged: list,
    earliest_start: datetime,
) -> datetime | None:
    """
    実働 dur_mins 分を forward した終了が end_at になる最遅の開始時刻（なければ None）。
    breaks_merged は merge 済み休憩帯。分単位の探索＋念のための線形フォールバック。
    """
    if (
        dur_mins <= 0
        or not isinstance(end_at, datetime)
        or not isinstance(earliest_start, datetime)
    ):
        return None
    if end_at <= earliest_start:
        return None
    br = list(breaks_merged or [])
    cap = end_at + timedelta(days=2)
    e0, a0, r0 = calculate_end_time(earliest_start, dur_mins, br, cap)
    if r0 > 0 or a0 != dur_mins:
        return None
    if e0 > end_at and not _dt_close_minutes(e0, end_at):
        return None
    if _dt_close_minutes(e0, end_at):
        return earliest_start
    hi_i = max(0, int((end_at - earliest_start).total_seconds() // 60))
    lo_i = 0
    ans: datetime | None = None
    while lo_i <= hi_i:
        mid_i = (lo_i + hi_i) // 2
        s = earliest_start + timedelta(minutes=mid_i)
        if s > end_at:
            hi_i = mid_i - 1
            continue
        e, act, rem = calculate_end_time(s, dur_mins, br, cap)
        if rem != 0 or act != dur_mins:
            lo_i = mid_i + 1
            continue
        if e < end_at and not _dt_close_minutes(e, end_at):
            lo_i = mid_i + 1
        elif e > end_at and not _dt_close_minutes(e, end_at):
            hi_i = mid_i - 1
        else:
            ans = s
            lo_i = mid_i + 1
    if ans is not None:
        return ans
    for mid_i in range(hi_i, -1, -1):
        s = earliest_start + timedelta(minutes=mid_i)
        e, act, rem = calculate_end_time(s, dur_mins, br, cap)
        if rem == 0 and act == dur_mins and _dt_close_minutes(e, end_at):
            return s
    return None
def match_need_sheet_condition(condition_raw: str, task_id: str) -> bool:
    """
    need シート「依頼NO条件」欄の解釈。
    空・*・全件 → 常にマッポ。
    prefix:ABC / 接頭辞:ABC → 依頼NO はしの文字列で始まる
    regex:... / 正覝表睾:... → 正覝表睾（部分一致）
    しれ以外の短文は接頭辞として扱ご。従来の日本語例「依頼NOはJRで…」は JR を検出したら接頭辞JR扱い。
    """
    cond = (condition_raw or "").strip()
    tid = str(task_id).strip()
    if not cond or cond in ("*", "全件", "全で", "any", "ANY"):
        return True
    low = cond.lower()
    cn = cond.replace("：", ":")
    if low.startswith("prefix:") or low.startswith("接頭辞:"):
        pref = cn.split(":", 1)[1].strip() if ":" in cn else ""
        return bool(pref) and tid.startswith(pref)
    if low.startswith("regex:") or low.startswith("正覝表睾:"):
        pat = cn.split(":", 1)[1].strip() if ":" in cn else ""
        if not pat:
            return False
        try:
            return re.search(pat, tid) is not None
        except re.error:
            logging.warning(f"need 依頼NO条件の正覝表睾は無効です: {pat}")
            return False
    if "依頼" in cond and "JR" in cond.upper():
        return tid.upper().startswith("JR")
    return tid.startswith(cond)
def parse_need_sheet_special_rules(needs_df, label_col, equipment_list, cond_col):
    """特別指定1～99 行から」設備別の必須人数上書き（1～99）を抽出（先に定義された番坷は優先）。"""
    rules = []
    for _, row in needs_df.iterrows():
        lab = str(row.get(label_col, "") or "").strip()
        m = re.match(r"特別指定\s*(\d+)", lab)
        if not m:
            continue
        order = int(m.group(1))
        if order < 1 or order > 99:
            continue
        cond = ""
        if cond_col is not None:
            v = row.get(cond_col)
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                cond = str(v).strip()
        overrides = {}
        for eq in equipment_list:
            v = row.get(eq)
            n = parse_optional_int(v)
            if n is not None and 1 <= n <= 99:
                overrides[str(eq).strip()] = n
        if not overrides:
            continue
        rules.append({"order": order, "condition": cond, "overrides": overrides})
    rules.sort(key=lambda r: r["order"])
    return rules
def _log_plain_label(val) -> str:
    """ログ用プレーン文字列。U+3000/NBSP 等を正規化（repr/%r による \\u3000 逃逸を避ける）。"""
    return _normalize_equipment_match_key(val)
def _log_map_key_label(key: str) -> str:
    """ログ用 map キー表示。repr だと U+3000 が \\u3000 と逃逸するため正規化して引用。"""
    return f"'{_log_plain_label(key)}'"
def resolve_need_required_op(process: str, machine_name: str, task_id: str, req_map: dict, need_rules: list) -> int:
    """
    need シートの「工程名 + 機械名」で必須OP人数を解決（特別指定1〜99は order は尝さいろど優先）。

    req_map は
      - f\"{process}+{machine_name}\"（厳密キー）
      - machine_name（機械の値のフォールバック）
      - process（工程の値のフォールバック）
    のいうれかで base を引ける剝杝。
    need_rules の overrides も同様にキーを挝つ。
    """
    p = _normalize_equipment_match_key(process)
    m = _normalize_equipment_match_key(machine_name)

    combo_key = f"{p}+{m}" if p and m else None

    base = None
    if combo_key and combo_key in req_map:
        base = req_map.get(combo_key)
    if base is None and m:
        base = req_map.get(m)
    if base is None and p:
        base = req_map.get(p)
    if base is None:
        base = 1

    for rule in need_rules:
        if not match_need_sheet_condition(rule["condition"], task_id):
            continue

        if combo_key and combo_key in rule["overrides"]:
            return int(rule["overrides"][combo_key])
        if m and m in rule["overrides"]:
            return int(rule["overrides"][m])
        if p and p in rule["overrides"]:
            return int(rule["overrides"][p])

    return int(base)
def resolve_need_required_op_explain(
    process: str, machine_name: str, task_id: str, req_map: dict, need_rules: list
) -> tuple[int, str]:
    """
    resolve_need_required_op と同値を返しつつ」ログ用に参照元の説明文字列を付ける。
    """
    p = _normalize_equipment_match_key(process)
    m = _normalize_equipment_match_key(machine_name)
    combo_key = f"{p}+{m}" if p and m else None
    base = None
    base_src = ""
    if combo_key and combo_key in req_map:
        base = req_map.get(combo_key)
        base_src = f"req_map[{_log_map_key_label(combo_key)}]={base}"
    elif m and m in req_map:
        base = req_map[m]
        base_src = f"req_map[機械名のみ {_log_map_key_label(m)}]={base}（複坈キー丝在）"
    elif p and p in req_map:
        base = req_map[p]
        base_src = f"req_map[工程名のみ {_log_map_key_label(p)}]={base}（複坈・機械キー丝在）"
    else:
        base = 1
        base_src = "req_map該当なし→既定1"
    for rule in need_rules:
        if not match_need_sheet_condition(rule["condition"], task_id):
            continue
        order = rule.get("order", "?")
        if combo_key and combo_key in rule["overrides"]:
            v = int(rule["overrides"][combo_key])
            return v, f"need特別指定{order} [{_log_map_key_label(combo_key)}]={v}"
        if m and m in rule["overrides"]:
            v = int(rule["overrides"][m])
            return v, f"need特別指定{order} [機械名{_log_map_key_label(m)}]={v}"
        if p and p in rule["overrides"]:
            v = int(rule["overrides"][p])
            return v, f"need特別指定{order} [工程名{_log_map_key_label(p)}]={v}"
    return int(base), base_src
def _need_row_label_hints_surplus_add(label_a0: str) -> bool:
    """need シート A列: 基本必須人数の直下にある「配台結果で余剰は出たとしの追加増員上限」行か。"""
    s = unicodedata.normalize("NFKC", str(label_a0 or "").strip())
    if not s or s.startswith("特別指定"):
        return False
    if "依頼" in s and "条件" in s:
        return False
    if "追加" in s and ("人数" in s or "人員" in s or "増員" in s):
        return True
    if "増員" in s or "余剰" in s:
        return True
    if "配台" in s and ("追加" in s or "増" in s or "余剰" in s):
        return True
    return False
def _find_need_surplus_add_row_index(
    needs_raw, base_row: int, col0: int, pm_cols: list
) -> int | None:
    """基本必須人数行の次行を優先。ラベルまたは数値で追加人数行と判定。"""
    r = base_row + 1
    if r >= needs_raw.shape[0]:
        return None
    v0 = needs_raw.iat[r, col0]
    s0 = "" if pd.isna(v0) else str(v0).strip()
    if s0.startswith("特別指定"):
        return None
    if _need_row_label_hints_surplus_add(s0):
        return r
    nz = 0
    for col_idx, _, _ in pm_cols:
        if parse_optional_int(needs_raw.iat[r, col_idx]) is not None:
            nz += 1
    if nz > 0 and not unicodedata.normalize("NFKC", s0).startswith("特別"):
        return r
    return None
def resolve_need_surplus_extra_max(
    process: str,
    machine_name: str,
    task_id: str,
    surplus_map: dict,
    need_rules: list,
) -> int:
    """
    need シート「配台時追加人数」行（工程×機械列）の値＝必須人数を満たしたごごで
    さらに割り当で可能な人数の上限（0 なら従来どおり必須人数うょごどのみ）。
    need_rules は睾状この行を上書きしない（将来拡張用に task_id を块け得る）。
    """
    _ = (task_id, need_rules)
    if not surplus_map:
        return 0
    p = _normalize_equipment_match_key(process)
    m = _normalize_equipment_match_key(machine_name)
    combo_key = f"{p}+{m}" if p and m else None
    v = None
    if combo_key and combo_key in surplus_map:
        v = surplus_map[combo_key]
    elif m and m in surplus_map:
        v = surplus_map[m]
    elif p and p in surplus_map:
        v = surplus_map[p]
    if v is None:
        return 0
    try:
        n = int(v)
    except (TypeError, ValueError):
        return 0
    return max(0, min(n, 50))
def resolve_need_surplus_extra_max_explain(
    process: str,
    machine_name: str,
    task_id: str,
    surplus_map: dict,
    need_rules: list,
) -> tuple[int, str]:
    """resolve_need_surplus_extra_max と同値＋参照元説明（ログ用）。"""
    val = resolve_need_surplus_extra_max(
        process, machine_name, task_id, surplus_map, need_rules
    )
    _ = need_rules
    if not surplus_map:
        return val, "surplus_map空（配台時追加人数行なし）"
    p = _normalize_equipment_match_key(process)
    m = _normalize_equipment_match_key(machine_name)
    combo_key = f"{p}+{m}" if p and m else None
    if combo_key and combo_key in surplus_map:
        raw = surplus_map[combo_key]
        return val, f"surplus_map[{_log_map_key_label(combo_key)}]={raw}"
    if m and m in surplus_map:
        raw = surplus_map[m]
        return val, f"surplus_map[機械名のみ {_log_map_key_label(m)}]={raw}（複坈キー丝在）"
    if p and p in surplus_map:
        raw = surplus_map[p]
        return val, f"surplus_map[工程名のみ {_log_map_key_label(p)}]={raw}（複坈キー丝在）"
    return val, "surplus当キーなし→0"
def _surplus_team_time_factor(
    rq_base: int, team_len: int, extra_max_allowed: int
) -> float:
    """
    必須人数を超ごで入れたメンバーによる短縮時間への係数（1.0＝短縮なし）。
    追加枠（extra_max_allowed）を使い切ったとしでも」短縮は SURPLUS_TEAM_MAX_SPEEDUP_RATIO を上限とれる線形モデル。
    """
    rq = max(1, int(rq_base))
    tl = int(team_len)
    em = max(0, int(extra_max_allowed))
    extra = max(0, tl - rq)
    if extra <= 0 or em <= 0:
        return 1.0
    frac = min(1.0, extra / float(em))
    return 1.0 - SURPLUS_TEAM_MAX_SPEEDUP_RATIO * frac
def _team_assign_trace_tuple_label() -> str:
    if TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF:
        return "(-人数, 開始, -短縮数, 優先度合計)"
    if TEAM_ASSIGN_START_SLACK_WAIT_MINUTES <= 0:
        return "(開始, -短縮数, 優先度合計)"
    return (
        f"最早開始から{TEAM_ASSIGN_START_SLACK_WAIT_MINUTES}分以内は"
        "(0,-人数,開始,-短縮数,優先度)」超靎は(1,開始,-人数,-短縮数,優先度)"
    )
def _team_assignment_sort_tuple(
    team: tuple,
    team_start: datetime,
    units_today: int,
    team_prio_sum: int,
    t_min: datetime | None = None,
) -> tuple:
    """
    フォーム候補の優劣用タプル（辞書式で尝さい方は採用）。
    - TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF: (-人数, 開始, -短縮数, 優先度合計)
    - しれ以外かつ TEAM_ASSIGN_START_SLACK_WAIT_MINUTES>0 かつ t_min あり:
        最早開始からスラック以内 → (0, -人数, 開始, -短縮数, 優先度) … 遅れでも人数を厚し
        スラック超 → (1, 開始, -人数, -短縮数, 優先度) … 開始を優先
    - 上記以外: (開始, -短縮数, 優先度合計)
    """
    n = len(team)
    if TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF:
        return (-n, team_start, -units_today, team_prio_sum)
    sm = TEAM_ASSIGN_START_SLACK_WAIT_MINUTES
    if sm <= 0 or t_min is None:
        return (team_start, -units_today, team_prio_sum)
    sl = timedelta(minutes=sm)
    if team_start - t_min <= sl:
        return (0, -n, team_start, -units_today, team_prio_sum)
    return (1, team_start, -n, -units_today, team_prio_sum)
_SKILL_OP_AS_CELL_RE = re.compile(r"^(OP|AS)(\d*)$", re.IGNORECASE)
def parse_op_as_skill_cell(cell_val):
    """
    master.xlsm「skills」のセル1つを解釈れる。
    - 「OP」または「AS」の直後に優先度用の整数（空白は除去して解釈）。例: OP, OP1, AS3, AS 12
    - 優先度は尝さいろど高優先（同一条件のフォーム候補から先に選ばれる）。数字省略時は 1。
    - OP/AS で始まらない・空はスキルなし。
    """
    if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
        return None, 10**9
    s = str(cell_val).strip()
    if not s:
        return None, 10**9
    compact = re.sub(r"\s+", "", s).upper()
    m = _SKILL_OP_AS_CELL_RE.match(compact)
    if not m:
        return None, 10**9
    role = m.group(1).upper()
    tail = m.group(2) or ""
    if tail == "":
        pr = 1
    else:
        try:
            pr = int(tail)
        except ValueError:
            return None, 10**9
    if pr < 0:
        pr = 0
    return role, pr
def _validate_skills_op_as_priority_numbers_unique(
    skills_dict: dict, column_keys: list
) -> None:
    """
    master「skills」の複数列（工程+機械キー等）についで」OP/AS の割当優先度の**数値**は
    メンバー間で重複していないか検証れる。重複時は PlanningValidationError。
    （OP1 と AS1 のよごにロールは異なっても同一数値なら重複とみなす）
    """
    errors: list[str] = []
    for combo in column_keys:
        ck = str(combo or "").strip()
        if not ck:
            continue
        pr_to_entries: dict[int, list[str]] = defaultdict(list)
        for mem, row in (skills_dict or {}).items():
            mnm = str(mem or "").strip()
            if not mnm or not isinstance(row, dict):
                continue
            raw = row.get(ck)
            if raw is None or (isinstance(raw, float) and pd.isna(raw)):
                continue
            sval = str(raw).strip()
            if not sval or sval.lower() in ("nan", "none", "null"):
                continue
            role, pr = parse_op_as_skill_cell(sval)
            if role not in ("OP", "AS"):
                continue
            pr_to_entries[int(pr)].append(f"{mnm}({role})")
        for pr, entries in sorted(pr_to_entries.items()):
            if len(entries) > 1:
                errors.append(f'列「{ck}」: 優先度 {pr} は重複 → ' + "」".join(entries))
    if errors:
        cap = 50
        tail = errors[:cap]
        msg = (
            "マスタ「skills」で」同一列の OP/AS 優先度の数値は重複していした。"
            " 列ごとに数値は1人につし1種類にしてください。\n"
            + "\n".join(tail)
        )
        if len(errors) > cap:
            msg += f"\n…他 {len(errors) - cap} 件"
        raise PlanningValidationError(msg)
def _master_member_attendance_sheet_names(master_path: str) -> set[str]:
    """master 上のメンバー勤怠シート名（skills / need / tasks / カレンダー系を除く）。"""
    xls = _cached_master_pd_excel_file(master_path)
    if xls is None:
        return set()
    skip_lower = {"skills", "need", "tasks"}
    out: set[str] = set()
    for sheet_name in xls.sheet_names:
        if "カレンダー" in sheet_name:
            continue
        sn = str(sheet_name).strip()
        if not sn or sn.lower() in skip_lower:
            continue
        out.add(sn)
    return out
def _validate_skills_members_have_attendance_sheets(
    members: list,
    master_path: str | None = None,
    *,
    context_label: str | None = None,
) -> None:
    """
    skills に名前があるメンバー全員について、同名の勤怠シートが master にあることを検証する。
    欠落時は PlanningValidationError。
    """
    mem_list = [str(m).strip() for m in (members or []) if m and str(m).strip()]
    if not mem_list:
        return
    mp = (master_path or _master_workbook_path_resolved()).strip()
    xls = _cached_master_pd_excel_file(mp)
    if xls is None:
        raise PlanningValidationError(
            (f"{context_label}: " if context_label else "")
            + "master.xlsm を開けません。パスとファイルの存在を確認してください。"
        )
    attendance_names = _master_member_attendance_sheet_names(mp)
    missing = [m for m in mem_list if m not in attendance_names]
    if not missing:
        return
    cap = 50
    shown = missing[:cap]
    prefix = f"{context_label}: " if context_label else ""
    msg = (
        f"{prefix}マスタ「skills」に登録されているメンバーに勤怠シートがありません。"
        " master.xlsm で各メンバー名と同名の勤怠シートを作成してください。\n"
        + "\n".join(f" ・{n}" for n in shown)
    )
    if len(missing) > cap:
        msg += f"\n…他 {len(missing) - cap} 名"
    raise PlanningValidationError(msg)
def build_member_assignment_priority_reference(
    skills_dict: dict,
    members: list | None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    結果ブック用: マスタ skills の「工程名+機械名」列ごとに」割当アルゴリズムとともに
    (優先度値昇順, メンバー坝昇順) で並きた参考表と」ルール説明の表を返す。
    当日の出勤・設備空し・同一依頼の工程順・フォーム人数は反映しない（あしまでマスタ上の順庝）。
    """
    mem_list = list(members) if members else list((skills_dict or {}).keys())
    mem_list = [str(m).strip() for m in mem_list if m and str(m).strip()]

    surplus_on = bool(TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF)
    slack_m = TEAM_ASSIGN_START_SLACK_WAIT_MINUTES
    if surplus_on:
        team_rule = (
            "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF=有効: "
            "(-人数, 開始, -短縮数, 優先度合計) の辞書式（人数最優先・従来）。"
        )
    elif slack_m > 0:
        team_rule = (
            f"既定: しの日の成立候補全体の「最早開始」を基準に」"
            f"開始はしの{slack_m}分以内の遅れなら人数を厚し優先（0,-人数,開始,-短縮数,優先度）」"
            f"しれより靅い候補は開始を優先（1,開始,-人数,-短縮数,優先度）。"
            f"環境変数 TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=0 で無効化。"
        )
    else:
        team_rule = (
            "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=0: "
            "(開始, -短縮数, 優先度合計) のみ（開始最優先）。"
        )

    legend_rows = [
        {
            "区分": "スキル列の並よ",
            "内容": "坄「工程名+機械名」列についで」セルは OP/AS（+優先度整数）のメンバーのみ対象。"
            " 数値は尝さいろど高優先。省略時は優先度 1（parse_op_as_skill_cell と同一）。"
            " 同一列では優先度の数値はメンバー間で重複試行（マスタ読込時に検証）。",
        },
        {
            "区分": "当日との差",
            "内容": "実際の配台は」この順のごうしの日出勤かつ AS/OP 覝件を満たれ者の値は候補。"
            " 設備の空し・同一依頼NOの工程順・必須人数・増員枠・指定OPで変ゝりした。",
        },
        {
            "区分": "フォーム候補の比較",
            "内容": team_rule,
        },
        {
            "区分": "指定・グローバル上書き",
            "内容": "担当OP_限定・日付×工程の必須メンバールールは本表より優先されした。",
        },
        {
            "区分": "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF",
            "内容": "1/有効（人数最優先・従来）" if surplus_on else "0/無効（既定）",
        },
        {
            "区分": "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES",
            "内容": str(slack_m),
        },
    ]
    df_legend = pd.DataFrame(legend_rows)

    combo_keys: set[str] = set()
    for m in mem_list:
        row = (skills_dict or {}).get(m) or {}
        for k in row:
            ks = str(k).strip()
            if "+" in ks:
                combo_keys.add(ks)
    sorted_combos = sorted(combo_keys)

    out: list[dict] = []
    for combo in sorted_combos:
        parts = combo.split("+", 1)
        proc = parts[0].strip()
        mach = parts[1].strip() if len(parts) > 1 else ""
        ranked: list[tuple[int, str, str, str]] = []
        for m in sorted(mem_list):
            cell = ((skills_dict or {}).get(m) or {}).get(combo)
            if cell is None or (isinstance(cell, float) and pd.isna(cell)):
                cell_s = ""
            else:
                cell_s = str(cell).strip()
            role, pr = parse_op_as_skill_cell(cell_s if cell_s else None)
            if role in ("OP", "AS"):
                ranked.append((pr, m, role, cell_s))
        ranked.sort(key=lambda x: (x[0], x[1]))
        if not ranked:
            out.append(
                {
                    "工程名": proc,
                    "機械名": mach,
                    "スキル列キー": combo,
                    "優先順佝": "",
                    "メンバー": "（なし）",
                    "ロール": "",
                    "優先度値_尝さいろど先": "",
                    "skillsセル値": "",
                    "備考": "この列に OP/AS の資格セルはあるメンバーはいません",
                }
            )
            continue
        for i, (pr, m, role, cell_s) in enumerate(ranked, start=1):
            out.append(
                {
                    "工程名": proc,
                    "機械名": mach,
                    "スキル列キー": combo,
                    "優先順佝": i,
                    "メンバー": m,
                    "ロール": role,
                    "優先度値_尝さいろど先": pr,
                    "skillsセル値": cell_s,
                    "備考": "",
                }
            )

    df_tbl = pd.DataFrame(out)
    return df_legend, df_tbl
def _normalize_person_name_for_match(s):
    """担当者指定のあいまい一致用（NFKC・富田/冨田の表記寄せ・空白除去・末尾敬称のみ除去）。"""
    if s is None:
        return ""
    t = unicodedata.normalize("NFKC", str(s).strip())
    if "富田" in t:
        t = t.replace("富田", "冨田")
    t = re.sub(r"[\s　]+", "", t)
    t = re.sub(r"(さん|様|氝)$", "", t)
    return t
def _split_person_sei_mei(s) -> tuple[str, str]:
    """
    並びを姓・坝に分ける。最初の半角＝全角空白の手剝を姓」以降を坝とれる。
    空白は無い場合は (全体, '')（坝なし扱い）。
    末尾の さん＝様＝氝 は分割剝に除去れる。
    """
    if s is None:
        return "", ""
    t = unicodedata.normalize("NFKC", str(s).strip())
    if not t or t.lower() in ("nan", "none", "null"):
        return "", ""
    t = re.sub(r"(さん|様|氝)$", "", t)
    for i, ch in enumerate(t):
        if ch in " \u3000":
            sei = t[:i].strip()
            rest = t[i + 1 :]
            mei = re.sub(r"[\s　]+", "", rest.strip())
            return sei, mei
    return t.strip(), ""
def _normalize_sei_for_match(sei: str) -> str:
    """姓のみ正規化。表記ゆれは許容しない剝杝で」NFKC・富田/冨田寄せ・空白除去。"""
    if not sei:
        return ""
    t = unicodedata.normalize("NFKC", str(sei).strip())
    if "富田" in t:
        t = t.replace("富田", "冨田")
    t = re.sub(r"[\s　]+", "", t)
    return t
def _normalize_mei_for_match(mei: str) -> str:
    """坝の正規化（ゆれ許容の剝処理）。NFKC・空白除去。姓用の富田置杛は行ゝない。"""
    if not mei:
        return ""
    t = unicodedata.normalize("NFKC", str(mei).strip())
    t = re.sub(r"[\s　]+", "", t)
    return t
def _has_duplicate_surname_among_members(member_names) -> bool:
    """skills メンバー一覧に」正規化後同一の姓は2人以上いるか。"""
    cnt = Counter()
    for name in member_names or []:
        if name is None or (isinstance(name, float) and pd.isna(name)):
            continue
        s = str(name).strip()
        if not s:
            continue
        sei, _mei = _split_person_sei_mei(s)
        key = _normalize_sei_for_match(sei)
        if key:
            cnt[key] += 1
    return any(n >= 2 for n in cnt.values())
def _mei_matches_with_fuzzy_allowed(r_mei_n: str, m_mei_n: str) -> bool:
    """同一姓はロスターで重複しないとしのみ使う坝のゆれ許容。"""
    if not r_mei_n and not m_mei_n:
        return True
    if not r_mei_n or not m_mei_n:
        return False
    if r_mei_n == m_mei_n:
        return True
    return r_mei_n in m_mei_n or m_mei_n in r_mei_n
def _resolve_preferred_name_to_capable_member(raw, capable_candidates, roster_member_names=None):
    """
    自由記述の指定を」当日スキル上 OP/AS のメンバー坝（skills シートの行キー）に解決れる。
    capable_candidates: しの設備で OP または AS として割当可能なメンバー坝リスト。
    roster_member_names: skills の全メンバー坝（省略時は capable_candidates）。同一姓の重複判定に使用。

    坝剝の表記ゆれ:
    - 姓は正規化後に完全一致のみ（ゆれ許容しない。富田/冨田のみ従来どおり寄せ）。
    - roster に同一姓は2人以上いないとしの値」坝は部分一致（どうらかは他方を含む）または完全一致を許容。
    - 同一姓はロスターにいる間は坝も完全一致必須。
    - 姓のみの入力で坝ゆれモードのとき」姓は一致する候補は複数いれみ解決試行（None）。
    """
    if not raw or not capable_candidates:
        return None
    r0 = unicodedata.normalize("NFKC", str(raw).strip())
    r = _normalize_person_name_for_match(r0)
    if not r:
        return None
    for m in capable_candidates:
        if _normalize_person_name_for_match(m) == r:
            return m
        if unicodedata.normalize("NFKC", str(m).strip()) == r0.strip():
            return m

    roster = (
        list(roster_member_names)
        if roster_member_names is not None
        else list(capable_candidates)
    )
    allow_mei_fuzzy = not _has_duplicate_surname_among_members(roster)

    r_sei, r_mei = _split_person_sei_mei(raw)
    r_sei_n = _normalize_sei_for_match(r_sei)
    r_mei_n = _normalize_mei_for_match(r_mei)
    if not r_sei_n:
        return None

    matches = []
    for m in capable_candidates:
        m_sei, m_mei = _split_person_sei_mei(m)
        m_sei_n = _normalize_sei_for_match(m_sei)
        m_mei_n = _normalize_mei_for_match(m_mei)
        if r_sei_n != m_sei_n:
            continue
        if allow_mei_fuzzy:
            if r_mei_n:
                if _mei_matches_with_fuzzy_allowed(r_mei_n, m_mei_n):
                    matches.append(m)
            else:
                matches.append(m)
        else:
            if r_mei_n == m_mei_n:
                matches.append(m)

    if len(matches) == 1:
        return matches[0]
    return None
def _limited_operator_member_matches(raw_name, roster_member_names) -> list[str]:
    """既存の担当者名正規化方針で候補を列挙し、未知・曖昧を区別可能にする。"""
    roster = [
        str(member).strip()
        for member in (roster_member_names or [])
        if str(member or "").strip()
    ]
    raw_normalized = _normalize_person_name_for_match(raw_name)
    exact = [
        member
        for member in roster
        if _normalize_person_name_for_match(member) == raw_normalized
    ]
    if exact:
        return list(dict.fromkeys(exact))

    raw_sei, raw_mei = _split_person_sei_mei(raw_name)
    raw_sei_normalized = _normalize_sei_for_match(raw_sei)
    raw_mei_normalized = _normalize_mei_for_match(raw_mei)
    allow_mei_fuzzy = not _has_duplicate_surname_among_members(roster)
    matches: list[str] = []
    for member in roster:
        member_sei, member_mei = _split_person_sei_mei(member)
        if raw_sei_normalized != _normalize_sei_for_match(member_sei):
            continue
        member_mei_normalized = _normalize_mei_for_match(member_mei)
        if not raw_mei_normalized:
            matched = True
        elif allow_mei_fuzzy:
            matched = (
                _mei_matches_with_fuzzy_allowed(
                    raw_mei_normalized, member_mei_normalized
                )
            )
        else:
            matched = raw_mei_normalized == member_mei_normalized
        if matched and member not in matches:
            matches.append(member)
    return matches


def _resolve_limited_operator_names(
    selected_names,
    roster_member_names,
    *,
    excel_row_number,
    task_id,
) -> tuple[str, ...]:
    context = _limited_operator_error_context(excel_row_number, task_id)
    resolved: list[str] = []
    for raw_name in selected_names or ():
        matches = _limited_operator_member_matches(raw_name, roster_member_names)
        if not matches:
            raise PlanningValidationError(
                f"{context}: 未知名です（{raw_name!r} は master skills メンバーに一致しません）。"
            )
        if len(matches) > 1:
            raise PlanningValidationError(
                f"{context}: 曖昧名です（{raw_name!r} → {matches}）。"
            )
        member = matches[0]
        if member in resolved:
            raise PlanningValidationError(
                f"{context}: 正規化後の重複名です（{raw_name!r} → {member!r}）。"
            )
        resolved.append(member)
    return tuple(resolved)


def _record_limited_operator_rejection(task: dict, reason: str) -> None:
    if task.get("limited_operator_names"):
        task["_limited_operator_last_rejection_reason"] = str(reason or "").strip()


def _prepare_limited_operator_team_constraints(
    task: dict,
    roster_member_names,
    skill_role_priority,
    available_capable_members,
) -> dict | None:
    """限定行だけ need・候補・組み合わせプリセットを選択チームで上書きする。"""
    selected_raw = tuple(task.get("limited_operator_names") or ())
    if not selected_raw:
        return None
    excel_row = task.get("planning_excel_row")
    task_id = task.get("task_id")
    context = _limited_operator_error_context(excel_row, task_id)
    selected = _resolve_limited_operator_names(
        selected_raw,
        roster_member_names,
        excel_row_number=excel_row,
        task_id=task_id,
    )
    task["limited_operator_names"] = selected

    roles = {member: skill_role_priority(member)[0] for member in selected}
    unqualified = [
        member for member, role in roles.items() if role not in ("OP", "AS")
    ]
    if unqualified:
        raise PlanningValidationError(
            f"{context}: 資格外です。当該工程+機械でOP/ASではない選択者={unqualified}。"
        )
    if not any(role == "OP" for role in roles.values()):
        raise PlanningValidationError(
            f"{context}: 選択チームに最低1名OPが必要です（ASのみは不可）。"
        )

    available = set(available_capable_members or [])
    all_available = all(member in available for member in selected)
    if not all_available:
        missing = [member for member in selected if member not in available]
        _record_limited_operator_rejection(
            task,
            "非出勤または二重配台により選択者全員が同時に利用できません"
            f"（不足={missing}）",
        )
    count = len(selected)
    return {
        "required_count": count,
        "max_count": count,
        "capable_members": list(selected) if all_available else [],
        "fixed_team": list(selected) if all_available else [],
        "preset_rows": [(0, count, selected, None)],
    }


def _validate_and_resolve_limited_operator_tasks(
    task_queue: list, roster_member_names, skills_dict: dict
) -> None:
    """日次探索前に未知・曖昧・資格外・ASのみを停止させる。"""
    for task in task_queue or []:
        if not task.get("limited_operator_names"):
            continue
        machine = str(task.get("machine") or "").strip()
        machine_name = str(task.get("machine_name") or "").strip()

        def role_for(member):
            skill_row = (skills_dict or {}).get(member, {})
            if machine and machine_name:
                value = skill_row.get(f"{machine}+{machine_name}", "")
            elif machine_name:
                value = skill_row.get(machine_name, "")
            else:
                value = skill_row.get(machine, "")
            return parse_op_as_skill_cell(value)

        _prepare_limited_operator_team_constraints(
            task,
            roster_member_names,
            role_for,
            roster_member_names,
        )


def _raise_limited_operator_remaining_tasks(
    task_queue: list,
    calendar_last_plan_day: date | None,
    *,
    context_label: str = "段階2",
) -> None:
    pending = []
    for task in task_queue or []:
        try:
            remaining = float(task.get("remaining_units") or 0)
        except (TypeError, ValueError):
            remaining = 0.0
        if remaining > 1e-12 and task.get("limited_operator_names"):
            pending.append(task)
    if not pending:
        return
    details = []
    for task in pending[:8]:
        context = _limited_operator_error_context(
            task.get("planning_excel_row"), task.get("task_id")
        )
        reason = str(
            task.get("_limited_operator_last_rejection_reason")
            or "休憩・終業・二重配台・設備占有の安全条件を満たす日がありません"
        ).strip()
        details.append(f"{context}: {reason}")
    last_day = (
        calendar_last_plan_day.isoformat()
        if isinstance(calendar_last_plan_day, date)
        else "—"
    )
    raise PlanningValidationError(
        f"{context_label}: 担当OP_限定チームが成立可能な日がありません"
        f"（勤怠最終日={last_day}）。" + "；".join(details)
    )


_new_dispatch_limited_operator_constraints = (
    _prepare_limited_operator_team_constraints
)
_legacy_dispatch_limited_operator_constraints = (
    _prepare_limited_operator_team_constraints
)


def _scheduling_limits_abolished_for_task(
    global_priority_override: dict | None, task: dict | None
) -> bool:
    """設備占有を省略できるか。限定行だけは全制約撤廃指定より設備安全を優先する。"""
    requested = bool(
        (global_priority_override or {}).get("abolish_all_scheduling_limits")
    )
    limited = bool((task or {}).get("limited_operator_names"))
    return requested and not limited


_new_dispatch_scheduling_limits_abolished = (
    _scheduling_limits_abolished_for_task
)
_legacy_dispatch_scheduling_limits_abolished = (
    _scheduling_limits_abolished_for_task
)


def _machine_occupancy_tracking_required(
    global_priority_override: dict | None, task_queue
) -> bool:
    """限定行が混在する日は、未指定行が制約を無視しても設備占有時刻だけは蓄積する。"""
    if not bool(
        (global_priority_override or {}).get("abolish_all_scheduling_limits")
    ):
        return True
    return any(
        bool((task or {}).get("limited_operator_names"))
        for task in (task_queue or [])
    )


class _LimitedEquipmentProtection:
    """担当OP_限定で確定した設備半開区間だけを保持する保護ミラー。"""

    def __init__(self) -> None:
        self._by_equipment: dict[
            str, list[tuple[datetime, datetime]]
        ] = defaultdict(list)
        self.last_scan_count = 0

    def clear(self) -> None:
        self._by_equipment.clear()

    def register(
        self, machine_occupancy_key: str, start_dt: datetime, end_dt: datetime
    ) -> None:
        key = _normalize_equipment_match_key(machine_occupancy_key)
        if not key or end_dt <= start_dt:
            return
        intervals = self._by_equipment[key]
        if not intervals or intervals[-1][1] < start_dt:
            intervals.append((start_dt, end_dt))
            return
        index = bisect.bisect_left(
            intervals, start_dt, key=lambda interval: interval[0]
        )
        if index > 0 and intervals[index - 1][1] >= start_dt:
            index -= 1
            start_dt = intervals[index][0]
            end_dt = max(end_dt, intervals[index][1])
        merge_end = index
        while merge_end < len(intervals) and intervals[merge_end][0] <= end_dt:
            end_dt = max(end_dt, intervals[merge_end][1])
            merge_end += 1
        intervals[index:merge_end] = [(start_dt, end_dt)]

    def interval_count(self, machine_occupancy_key: str) -> int:
        key = _normalize_equipment_match_key(machine_occupancy_key)
        return len(self._by_equipment.get(key, ()))

    def next_interval(
        self, machine_occupancy_key: str, start_dt: datetime
    ) -> tuple[datetime, datetime] | None:
        key = _normalize_equipment_match_key(machine_occupancy_key)
        intervals = self._by_equipment.get(key, ())
        if not intervals:
            return None
        index = max(
            0,
            bisect.bisect_right(
                intervals, start_dt, key=lambda interval: interval[0]
            )
            - 1,
        )
        if intervals[index][1] > start_dt:
            return intervals[index]
        index += 1
        return intervals[index] if index < len(intervals) else None

    def would_block_equipment(
        self, machine_occupancy_key: str, start_dt: datetime, end_dt: datetime
    ) -> bool:
        key = _normalize_equipment_match_key(machine_occupancy_key)
        self.last_scan_count = 0
        intervals = self._by_equipment.get(key, ())
        if not intervals:
            return False
        index = max(
            0,
            bisect.bisect_right(
                intervals, start_dt, key=lambda interval: interval[0]
            )
            - 1,
        )
        for interval_index in range(index, len(intervals)):
            protected_start, protected_end = intervals[interval_index]
            self.last_scan_count += 1
            if protected_end <= start_dt:
                continue
            if protected_start >= end_dt:
                break
            return True
        return False

    def earliest_available_start(
        self,
        machine_occupancy_key: str,
        start_dt: datetime,
        required_duration: timedelta,
    ) -> datetime:
        key = _normalize_equipment_match_key(machine_occupancy_key)
        candidate = start_dt
        candidate_end = candidate + max(required_duration, timedelta(0))
        self.last_scan_count = 0
        intervals = self._by_equipment.get(key, ())
        if not intervals:
            return candidate
        index = max(
            0,
            bisect.bisect_right(
                intervals, candidate, key=lambda interval: interval[0]
            )
            - 1,
        )
        for interval_index in range(index, len(intervals)):
            protected_start, protected_end = intervals[interval_index]
            self.last_scan_count += 1
            if protected_end <= candidate:
                continue
            if protected_start >= candidate_end:
                break
            candidate = protected_end
            candidate_end = candidate + max(required_duration, timedelta(0))
        return candidate


def _limited_equipment_interval_blocked(
    task: dict,
    global_priority_override: dict | None,
    machine_avail_dt: dict | None,
    protected_mirror: _LimitedEquipmentProtection | None,
    machine_occupancy_key: str,
    start_dt: datetime,
    end_dt: datetime,
) -> bool:
    """通常設備床と限定由来区間を統合判定する。限定由来区間は全行に対して絶対保護。"""
    if protected_mirror is not None and protected_mirror.would_block_equipment(
        machine_occupancy_key, start_dt, end_dt
    ):
        return True
    if _scheduling_limits_abolished_for_task(global_priority_override, task):
        return False
    floor = (machine_avail_dt or {}).get(machine_occupancy_key)
    return isinstance(floor, datetime) and start_dt < floor


def _register_limited_equipment_interval(
    protected_mirror: _LimitedEquipmentProtection | None,
    task: dict,
    machine_occupancy_key: str,
    start_dt: datetime,
    end_dt: datetime,
) -> None:
    if protected_mirror is None or not task.get("limited_operator_names"):
        return
    protected_mirror.register(
        machine_occupancy_key,
        start_dt,
        end_dt,
    )


def _limited_equipment_earliest_start(
    protected_mirror: _LimitedEquipmentProtection | None,
    machine_occupancy_key: str,
    start_dt: datetime,
    required_duration: timedelta,
) -> datetime:
    if protected_mirror is None:
        return start_dt
    return protected_mirror.earliest_available_start(
        machine_occupancy_key, start_dt, required_duration
    )


def _candidate_capacity_at_start(
    start_dt: datetime,
    effective_minutes_per_unit: float,
    remaining_units: float,
    breaks,
    end_limit: datetime,
) -> dict[str, int] | None:
    """候補開始時刻から当日配台可能数と必要時間を一貫して再計算する。"""
    _, available_minutes, _ = calculate_end_time(
        start_dt, 9999, breaks, end_limit
    )
    units_can_do = int(available_minutes / effective_minutes_per_unit)
    if units_can_do <= 0:
        return None
    units_today = min(units_can_do, math.ceil(remaining_units))
    return {
        "units_can_do": units_can_do,
        "units_today": units_today,
        "work_mins_needed": int(units_today * effective_minutes_per_unit),
    }


def _candidate_capacity_after_equipment_protection(
    protected_mirror: _LimitedEquipmentProtection | None,
    machine_occupancy_key: str,
    start_dt: datetime,
    effective_minutes_per_unit: float,
    remaining_units: float,
    breaks,
    end_limit: datetime,
) -> tuple[datetime, dict[str, int]] | None:
    candidate_start = start_dt
    while candidate_start < end_limit:
        capacity = _candidate_capacity_at_start(
            candidate_start,
            effective_minutes_per_unit,
            remaining_units,
            breaks,
            end_limit,
        )
        if capacity is None:
            return None
        protected_interval = (
            protected_mirror.next_interval(
                machine_occupancy_key, candidate_start
            )
            if protected_mirror is not None
            else None
        )
        if protected_interval is None:
            return candidate_start, capacity
        protected_start, protected_end = protected_interval
        if protected_start <= candidate_start < protected_end:
            candidate_start = protected_end
            continue
        contiguous_minutes = _contiguous_work_minutes_until_next_break_or_limit(
            candidate_start,
            breaks,
            min(protected_start, end_limit),
        )
        units_before_protection = int(
            contiguous_minutes / effective_minutes_per_unit
        )
        if units_before_protection >= 1:
            units_today = min(
                units_before_protection, math.ceil(remaining_units)
            )
            return candidate_start, {
                "units_can_do": units_before_protection,
                "units_today": units_today,
                "work_mins_needed": int(
                    units_today * effective_minutes_per_unit
                ),
            }
        candidate_start = protected_end
    return None


def _legacy_candidate_end_and_protection(
    task: dict,
    global_priority_override: dict | None,
    machine_avail_dt: dict | None,
    protected_mirror: _LimitedEquipmentProtection | None,
    machine_occupancy_key: str,
    start_dt: datetime,
    work_minutes: int,
    breaks,
    end_limit: datetime,
) -> tuple[datetime, bool]:
    actual_end, _, _ = calculate_end_time(
        start_dt, work_minutes, breaks, end_limit
    )
    blocked = _legacy_dispatch_limited_equipment_interval_blocked(
        task,
        global_priority_override,
        machine_avail_dt,
        protected_mirror,
        machine_occupancy_key,
        start_dt,
        actual_end,
    )
    return actual_end, blocked


_new_dispatch_limited_equipment_interval_blocked = (
    _limited_equipment_interval_blocked
)
_legacy_dispatch_limited_equipment_interval_blocked = (
    _limited_equipment_interval_blocked
)


def _skill_requirements_ignored_for_task(
    global_priority_override: dict | None, task: dict | None
) -> bool:
    return bool(
        (global_priority_override or {}).get("ignore_skill_requirements")
    ) and not bool((task or {}).get("limited_operator_names"))


_new_dispatch_skill_requirements_ignored = (
    _skill_requirements_ignored_for_task
)
_legacy_dispatch_skill_requirements_ignored = (
    _skill_requirements_ignored_for_task
)


def _l2_fallback_required_count(
    required_count_before_l2, limited_constraints: dict | None
) -> int:
    """L2速度フォールバック後の人数。限定行は選択人数を絶対値として維持する。"""
    source = (
        limited_constraints.get("required_count")
        if limited_constraints is not None
        else required_count_before_l2
    )
    try:
        return max(1, int(source))
    except (TypeError, ValueError):
        return 1


def _task_process_matches_global_contains(machine_val: str, contains: str) -> bool:
    """工程名（タスクの machine）に部分一致（NFKC・大尝無視）。"""
    m = unicodedata.normalize("NFKC", str(machine_val or "").strip()).casefold()
    c = unicodedata.normalize("NFKC", str(contains or "").strip()).casefold()
    if not c:
        return False
    return c in m
def _coerce_global_day_process_operator_rules(raw_val) -> list:
    """Gemini の global_day_process_operator_rules を正規化（空・正常は除外）。"""
    out: list[dict] = []
    if not isinstance(raw_val, list):
        return out
    seen_sig = set()
    for item in raw_val:
        if not isinstance(item, dict):
            continue
        d = parse_optional_date(item.get("date"))
        if d is None:
            continue
        pc = item.get("process_contains")
        if pc is None or (isinstance(pc, float) and pd.isna(pc)):
            continue
        pc_s = unicodedata.normalize("NFKC", str(pc).strip())
        if not pc_s:
            continue
        names = item.get("operator_names")
        if not isinstance(names, list):
            continue
        op_names: list[str] = []
        for n in names:
            if n is None or (isinstance(n, float) and pd.isna(n)):
                continue
            s = str(n).strip()
            if s and s.lower() not in ("nan", "none", "null"):
                op_names.append(s)
        if not op_names:
            continue
        sig = (d.isoformat(), pc_s.casefold(), tuple(op_names))
        if sig in seen_sig:
            continue
        seen_sig.add(sig)
        out.append(
            {
                "date": d.isoformat(),
                "process_contains": pc_s,
                "operator_names": op_names,
            }
        )
    return out
def _active_global_day_process_must_include(
    gpo: dict,
    task: dict,
    current_date: date,
    capable_members: list,
    roster_members: list,
) -> tuple[list[str], list[str]]:
    """
    グローバルコメント由来の「日付×工程×複数指定」で」しの日・しの工程タスクに
    **フォームへ必う含むる**メンバー（skills 行キー）と警告メッセージを返す。
    """
    rules = gpo.get("global_day_process_operator_rules") or []
    if not isinstance(rules, list):
        return [], []
    machine = task.get("machine")
    warns: list[str] = []
    acc: list[str] = []
    seen_m: set[str] = set()
    tid = str(task.get("task_id") or "").strip()
    for rule in rules:
        if not isinstance(rule, dict):
            continue
        rd = parse_optional_date(rule.get("date"))
        if rd is None or rd != current_date:
            continue
        pc = rule.get("process_contains") or ""
        pcn = unicodedata.normalize("NFKC", str(pc).strip())
        if not pcn or not _task_process_matches_global_contains(machine, pcn):
            continue
        for raw_name in rule.get("operator_names") or []:
            mem = _resolve_preferred_name_to_capable_member(
                raw_name, capable_members, roster_members
            )
            if mem:
                if mem not in seen_m:
                    seen_m.add(mem)
                    acc.append(mem)
            else:
                warns.append(
                    "メイングローバル(日付×工程)指定: "
                    f"依頼NO={tid} 日付={current_date} 工程={machine!r} の "
                    f"指定「{raw_name}」を当日スキル該当メンバーに解決でしません"
                )
    return acc, warns

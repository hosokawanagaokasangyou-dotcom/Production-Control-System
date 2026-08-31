# -*- coding: utf-8 -*-
# planning_core.core.dispatch_loop — body only (loaded via _core exec chain)
def _assign_sequential_dispatch_trial_order(task_queue: list) -> None:
    """
    `task_queue` のリスト順に合わせで `dispatch_trial_order` を 1..n へ付け直れ。
    `_reorder_task_queue_b2_ec_inspection_consecutive` の直後（よよよキュー再ソートの直後）に呼よ」
    EC と後続（検査＝巻返し）の連続番坷を保証れる。
    """
    for i, t in enumerate(task_queue, start=1):
        t["dispatch_trial_order"] = i
def _interactive_trial_shortage_meters_done_for_rec(
    rec: dict,
    meters_done: dict[tuple[str, str, str, date], float] | None,
    *,
    eps: float = 1e-3,
) -> float:
    """不足レコードの (依頼NO, 工程, 機械, 日) に対応するタイムライン実績 m。"""
    if not meters_done:
        return 0.0
    tid = _interactive_norm_cell(rec.get("task_id"))
    proc = _interactive_dispatch_target_process_key(rec.get("process"))
    mach = _interactive_norm_cell(rec.get("machine_name"))
    date_s = str(rec.get("date") or "").strip()
    try:
        dd = date.fromisoformat(date_s[:10])
    except ValueError:
        return 0.0
    key = (tid, proc, mach, dd)
    try:
        done = float(meters_done.get(key, 0.0))
    except (TypeError, ValueError):
        done = 0.0
    if done > eps:
        return done
    for kk, vv in meters_done.items():
        if not isinstance(kk, tuple) or len(kk) != 4:
            continue
        if kk[0] != tid or kk[2] != mach or kk[3] != dd:
            continue
        try:
            done = max(done, float(vv or 0.0))
        except (TypeError, ValueError):
            continue
    return done
def _interactive_trial_shortage_meters_done_for_triple(
    rec: dict,
    meters_done: dict[tuple[str, str, str, date], float] | None,
    *,
    eps: float = 1e-3,
) -> float:
    """不足レコードの (依頼NO, 工程, 機械) について、全配台日の実績 m 合計。"""
    if not meters_done:
        return 0.0
    tid = _interactive_norm_cell(rec.get("task_id"))
    proc = _interactive_dispatch_target_process_key(rec.get("process"))
    mach = _interactive_norm_cell(rec.get("machine_name"))
    total = 0.0
    for kk, vv in meters_done.items():
        if not isinstance(kk, tuple) or len(kk) != 4:
            continue
        if kk[0] != tid:
            continue
        if _interactive_dispatch_target_process_key(kk[1]) != proc:
            continue
        if _interactive_norm_cell(kk[2]) != mach:
            continue
        try:
            total += float(vv or 0.0)
        except (TypeError, ValueError):
            continue
    return total
def _interactive_trial_calendar_legacy_active() -> bool:
    """
    True のときのみ、従来のインタラクティブ試行専用の機械カレンダー解釈
    （※占有のみ・未定義スロットの全日ブロック・終業 23:59 延長など）を使う。
    段階2同一パリティでは False。
    """
    return _interactive_dispatch_trial_env_active() and not _interactive_stage2_parity_active()
def _interactive_trial_relax_team_end_limit_to_eod(
    team_end_limit: datetime, current_date: date
) -> datetime:
    """
    デスクトップ配台試行のみ: 出勤簿の終業より遅く機械が空く場合でも同日ロールを試せるよう、
    チーム終業上限を当日 23:59 まで引き上げる（暦日跨ぎ加工は別チェックでエラー）。
    """
    if not (_interactive_trial_calendar_legacy_active() or _stage3_qty_strict_active()):
        return team_end_limit
    cap = datetime.combine(current_date, time(23, 59))
    return cap if team_end_limit < cap else team_end_limit
def _interactive_trial_pair_dates_from_targets(
    targets: dict | None,
) -> dict[tuple[str, str], set[date]]:
    """
    結果_配台表 JSON 由来の targets に現れる (依頼NO, 機械名) ごとの配台日集合。
    キーに含まれない行は暦日制限をかけない（従来どおり段階2に委ねる）。

    段階3（PM_AI_INTERACTIVE_DISPATCH_TRIAL）では既定で本マップを割当フィルタに使わない。
    有効化は環境変数 PM_AI_INTERACTIVE_TRIAL_PAIR_DATES=1。
    """
    out: dict[tuple[str, str], set[date]] = {}
    if not targets:
        return out
    for key in targets.keys():
        if not isinstance(key, tuple):
            continue
        if len(key) == 4:
            tid, _proc, mach, dd = key[0], key[1], key[2], key[3]
        elif len(key) == 3:
            tid, mach, dd = key[0], key[1], key[2]
        else:
            continue
        k = (_interactive_norm_cell(tid), _interactive_norm_cell(mach))
        if k not in out:
            out[k] = set()
        if isinstance(dd, date):
            out[k].add(dd)
    return out
def _changeover_need_cleanup_for_next_assign(
    *,
    machine_handoff: dict,
    machine_occ_key: str,
    current_date: date,
    cur_task_id: str,
    last_eq: str | None,
) -> tuple[bool, int, str, str]:
    """
    依頼NO切替時の後始末要否（同一依頼NOは不要）。
    """
    mach_occ = str(machine_occ_key or "").strip()
    last_eq_s = str(last_eq or "").strip() or str(
        (machine_handoff.get("last_eq") or {}).get(mach_occ, "") or ""
    ).strip()
    if not _needs_request_switch_prep(
        machine_handoff, mach_occ, current_date, cur_task_id
    ):
        return False, 0, "", last_eq_s
    _prev_proc, _prev_mn = _normalize_proc_machine_for_prep_lookup(
        "", "", eq_line=last_eq_s
    )
    cu = _lookup_post_machining_cleanup_minutes(
        _prev_proc, _prev_mn, eq_line=last_eq_s
    )
    last_lead = str(
        (machine_handoff.get("last_lead_op") or {}).get(mach_occ, "") or ""
    ).strip()
    return cu > 0, cu, last_lead, last_eq_s
def _changeover_plan_segments_and_machining_lower_bound(
    *,
    prev_machining_end_dt: datetime,
    machine_day_floor: datetime,
    current_date: date,
    machine_occ_key: str,
    task_id: str,
    eq_line: str,
    machine_name: str,
    machine_proc: str,
    machine_handoff: dict,
    daily_status: dict,
    skills_dict: dict,
    abolish_limits: bool,
) -> tuple[datetime | None, list[dict]]:
    """
    前ロール加工終了 prev_machining_end_dt から、日次始業（当日先頭のみ）のみを
    組み立て、(加工開始最早時刻, タイムライン用セグメント雛形) を返す。
    日次始業の壁時計ブロックは、定常／工場開始と機械占有のうち遅い時刻を下限とし、
    当該機械の OP スキル保持者が勤務帯で一括覆える最早の半開区間 [開始, 開始+N分) に置く。
    担当者名はタイムライン追記時に別途埋める（セグメント生成時点では op は空のまま）。
    """
    if abolish_limits:
        return prev_machining_end_dt, []
    mach_occ = str(machine_occ_key or "").strip()
    machining_today_occ = machine_handoff.get("machining_today_occ") or machine_handoff.get(
        "started_today", set()
    )

    segments: list[dict] = []
    t = prev_machining_end_dt

    se = _daily_startup_segment_start_end(
        prev_machining_end_dt=prev_machining_end_dt,
        current_date=current_date,
        machine_name=machine_name,
        machine_proc=str(machine_proc or "").strip(),
        machine_occ_key=mach_occ,
        machine_handoff=machine_handoff,
        skills_dict=skills_dict,
        daily_status=daily_status,
        daily_startup_by_machine=None,
    )
    if se:
        reg_start_dt, reg_end_dt = se
        segments.append(
            {
                "start_dt": reg_start_dt,
                "end_dt": reg_end_dt,
                "op": "",
                "event_kind": TIMELINE_EVENT_MACHINE_DAILY_STARTUP,
                "machine": eq_line,
                "machine_occupancy_key": mach_occ,
            }
        )
        t = max(t, reg_end_dt)

    return t, segments
def _changeover_timeline_op_sub_for_event(
    *,
    event_kind: str,
    op_from_segment: str,
    sub_from_segment: str,
    machine_occ_key: str,
    machining_lead_op: str,
    machining_sub_str: str,
    machine_handoff: dict,
    daily_status: dict,
) -> tuple[str, str]:
    """タイムライン用の主＝補。日次始業はセグメントに事前設定があればそれを採用。"""
    ek = str(event_kind or "").strip()
    op_s = str(op_from_segment or "").strip()
    if ek in (
        TIMELINE_EVENT_MACHINE_DAILY_STARTUP,
        TIMELINE_EVENT_REQUEST_SWITCH_PREP,
        TIMELINE_EVENT_BREAK_RESUME_PREP,
        TIMELINE_EVENT_POST_MACHINING_CLEANUP,
        TIMELINE_EVENT_REQUEST_INTERVAL_BUFFER,
    ):
        if op_s:
            return op_s, str(sub_from_segment or "").strip()
        return "", ""
    return op_s, ""
def _trial_order_flow_day_start_floor(
    task: dict,
    current_date: date,
    macro_run_date: date,
    macro_now_dt: datetime,
    task_queue: list | None = None,
) -> datetime:
    """当該暦日の加工開始下限。配台可能日時（列）があれば原反投入日+12:45 より優先。"""
    floor = datetime.combine(current_date, DEFAULT_START_TIME)
    # §B-2 検査 / §B-3 巻返しは EC 完了を待って開始でしるため、
    # 原板投入日（=同日12:45以降）の制約をしのまま適用すると後続は丝必須に後ゝへ倒れる。
    # EC完了時刻下限（_roll_pipeline_b2_inspection_ec_completion_floor_dt）で整合を得る。
    # EC 行がキューに無い（完走後に行欠落）場合も後続フラグが付いていれば B2 後続として扱い、
    # 原板同日の 12:45 下限を付けない（EC 行が残っている完走ケースと整合。L10 スリット欠落と同趣旨）。
    _tid_floor = str(task.get("task_id", "") or "").strip()
    is_b2_follower_delayed = bool(
        (task.get("roll_pipeline_inspection") or task.get("roll_pipeline_rewind"))
        and _tid_floor
        and task_queue is not None
    )
    _disp_dt = task.get("dispatchable_datetime")
    _has_disp_col = isinstance(_disp_dt, datetime)
    rid = task.get("raw_input_date")
    if (
        not is_b2_follower_delayed
        and isinstance(rid, date)
        and rid == current_date
        and not _has_disp_col
    ):
        floor = max(
            floor,
            datetime.combine(
                current_date, dispatchable_from_time_for(task.get("stock_location"))
            ),
        )
    sdl = task.get("same_day_raw_start_limit")
    s_req = task.get("start_date_req")
    if (
        (not is_b2_follower_delayed)
        and sdl
        and isinstance(s_req, date)
        and current_date == s_req
        and isinstance(sdl, time)
    ):
        floor = max(floor, datetime.combine(current_date, sdl))
    est = task.get("earliest_start_time")
    if (
        (not is_b2_follower_delayed)
        and isinstance(s_req, date)
        and current_date == s_req
        and est
    ):
        if isinstance(est, time):
            floor = max(floor, datetime.combine(current_date, est))
    skip_macro_now = _interactive_stage3_skip_macro_now_start_floor(task, current_date)
    if current_date == macro_run_date and floor < macro_now_dt and not skip_macro_now:
        floor = macro_now_dt
    return floor
def _trial_order_flow_eligible_tasks(
    tasks_today: list,
    task_queue: list,
    current_date: date,
    *,
    daily_status: dict | None = None,
    members: list | None = None,
    machine_avail_dt: dict | None = None,
    machine_day_start: datetime | None = None,
    machine_handoff: dict | None = None,
    skills_dict: dict | None = None,
    abolish_all_scheduling_limits: bool = False,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
    min_dispatch_effective: int | None = None,
    assign_probe_ctx: dict | None = None,
    interactive_trial_pair_dates: dict | None = None,
    wip_snapshot: dict | None = None,
    pending_by_occ: dict[str, list[tuple[int, dict]]] | None = None,
    rows_by_tid: dict[str, list[dict]] | None = None,
    rows_by_rule_tid: dict[str, list[dict]] | None = None,
    window_left_cache: dict | None = None,
    pipeline_room_cache: dict[str, float] | None = None,
) -> list:
    # 特別ルール（工程間WIP上限）: L11 は EC→（検査＋巻返し）の前段 WIP が上限以上なら EC を配台しない。
    if wip_snapshot is not None:
        _wip_l11_global_val = wip_snapshot.get("l11_global")
        _wip_l11_by_bucket = wip_snapshot.setdefault("l11_by_bucket", {})
        wip_slit_before_sec = wip_snapshot.get("wip_slit_before_sec")
        wip_connection_before_sec = wip_snapshot.get("wip_connection_before_sec")
    else:
        _wip_snap = _stage2_eligible_wip_snapshot(task_queue)
        _wip_l11_global_val = _wip_snap.get("l11_global")
        _wip_l11_by_bucket = _wip_snap["l11_by_bucket"]
        wip_slit_before_sec = _wip_snap.get("wip_slit_before_sec")
        wip_connection_before_sec = _wip_snap.get("wip_connection_before_sec")

    from planning_core.dispatch_rules.hook_adapter import (
        build_eligible_l13_context,
        eligible_l13_connection_skip_with_context,
    )

    out = []
    _elig_acc_early = 0.0
    _elig_acc_rules = 0.0
    _elig_acc_window = 0.0
    _elig_acc_equip = 0.0
    _l13_ctx = build_eligible_l13_context()
    _b41_gate_cache: dict[str, bool] = {}
    _b61_gate_cache: dict[str, bool] = {}
    for task in tasks_today:
        if float(task.get("remaining_units") or 0) <= 1e-12:
            continue
        _abolish_for_task = _scheduling_limits_abolished_for_task(
            {"abolish_all_scheduling_limits": abolish_all_scheduling_limits},
            task,
        )
        _t_task0 = time_module.perf_counter()
        try:
            if (
                _interactive_dispatch_trial_env_active()
                and interactive_trial_pair_dates is not None
            ):
                tid_n = _interactive_norm_cell(str(task.get("task_id") or ""))
                mach_n = _interactive_norm_cell(str(task.get("machine_name") or ""))
                _pd = interactive_trial_pair_dates.get((tid_n, mach_n))
                if _pd is not None and current_date not in _pd:
                    continue
            # L11: 検査前WIPが限界以上なら EC をブロック（集計は AGGREGATE_MODE）。
            # 配台キューに B-2 検査または B-3 巻返し（異設備）後続が無い依頼は対象外
            # （同一 EC 機上的巻返しは段階1配台不要のため L11 で EC を止めない）。
            if isinstance(WIP_LIMIT_EC_BEFORE_INSP_ROLLS, int) and WIP_LIMIT_EC_BEFORE_INSP_ROLLS > 0:
                if task.get("roll_pipeline_ec"):
                    _tid_l11 = str(task.get("task_id") or "").strip()
                    if _tid_l11 and _b2_ec_insp_pair_in_queue(task_queue, _tid_l11):
                        if _wip_ec_l11_aggregate_is_global():
                            _wip_use = _wip_l11_global_val
                            _wip_cache_key = "global"
                        else:
                            _m = WIP_LIMIT_EC_BEFORE_INSP_AGGREGATE_MODE
                            if _m == "task_id":
                                _wip_bk = _tid_l11
                            else:
                                _wip_bk = _wip_l11_bucket_key_for_task_id(_tid_l11)
                            _wip_cache_key = f"{_m}:{_wip_bk}"
                            if _wip_cache_key not in _wip_l11_by_bucket:
                                if _m == "task_id":
                                    _wip_l11_by_bucket[_wip_cache_key] = (
                                        _wip_ec_before_insp_roll_count(
                                            task_queue, task_id_exact=_wip_bk
                                        )
                                    )
                                else:
                                    _wip_l11_by_bucket[_wip_cache_key] = (
                                        _wip_ec_before_insp_roll_count(
                                            task_queue, task_id_head=_wip_bk
                                        )
                                    )
                            _wip_use = _wip_l11_by_bucket[_wip_cache_key]
                        if (
                            _wip_use is not None
                            and _wip_use >= float(WIP_LIMIT_EC_BEFORE_INSP_ROLLS)
                        ):
                            continue
            # L10: SEC前WIPが限界以上ならスリットをブロック（SECは進めてWIP解消）
            if wip_slit_before_sec is not None and wip_slit_before_sec >= float(
                WIP_LIMIT_SLIT_BEFORE_SEC_ROLLS
            ):
                proc = _normalize_process_name_for_rule_match(task.get("machine"))
                mach = _normalize_equipment_match_key(task.get("machine_name"))
                if (
                    proc
                    == _normalize_process_name_for_rule_match(SPECIAL_WIP_SLIT_PROCESS)
                    and mach
                    == _normalize_equipment_match_key(SPECIAL_WIP_SLIT_MACHINE)
                    and _task_on_slit_sec_process_path(task)
                ):
                    continue

            # B-6 / L13: 接続 WIP 上限（rules/plan は eligible 1 回あたり 1 度だけ読込）
            if eligible_l13_connection_skip_with_context(
                task, wip_connection_before_sec, task_queue, _l13_ctx
            ):
                continue

            _tid_gate = str(task.get("task_id") or "").strip()
            if _tid_gate:
                if _tid_gate not in _b41_gate_cache:
                    _b41_gate_cache[_tid_gate] = _l10_b41_sec_blocked_by_slit_min_rolls(
                        task, task_queue, rows_by_tid=rows_by_tid
                    )
                if _b41_gate_cache[_tid_gate]:
                    continue
                if _tid_gate not in _b61_gate_cache:
                    _b61_gate_cache[_tid_gate] = _b61_sec_blocked_by_connection_min_rolls(
                        task, task_queue, rows_by_tid=rows_by_tid
                    )
                if _b61_gate_cache[_tid_gate]:
                    continue
            else:
                if _l10_b41_sec_blocked_by_slit_min_rolls(
                    task, task_queue, rows_by_tid=rows_by_tid
                ):
                    continue
                if _b61_sec_blocked_by_connection_min_rolls(
                    task, task_queue, rows_by_tid=rows_by_tid
                ):
                    continue
            if _task_blocked_by_same_request_dependency(
                task,
                task_queue,
                rows_by_rule_tid=rows_by_rule_tid,
                pipeline_room_cache=pipeline_room_cache,
            ):
                continue
            if _task_blocked_by_global_dispatch_trial_order(
                task,
                task_queue,
                current_date,
                daily_status=daily_status,
                members=members,
                machine_avail_dt=machine_avail_dt,
                machine_day_start=machine_day_start,
                machine_handoff=machine_handoff,
                skills_dict=skills_dict,
                abolish_all_scheduling_limits=_abolish_for_task,
                dispatch_interval_mirror=dispatch_interval_mirror,
                min_dispatch_effective=min_dispatch_effective,
            ):
                continue
        finally:
            _elig_acc_early += time_module.perf_counter() - _t_task0

        _t_er0 = time_module.perf_counter()
        _t_ew0 = time_module.perf_counter()
        # min_dto から全日カレンダー占有は除外済みでも」同日試行順の「ブロック」は my_o>m のみのため、
        # 試行順=min の占有行は残り」他試行順は永久坜止し得る。当日スロットゼロの行は候補外にれる。
        if daily_status is not None and members is not None:
            if _task_fully_machine_calendar_blocked_on_date(
                task, current_date, daily_status, members
            ):
                continue
            if _task_no_machining_window_left_from_avail_floor_cached(
                task,
                current_date,
                daily_status,
                members,
                machine_avail_dt,
                machine_day_start,
                machine_handoff=machine_handoff,
                skills_dict=skills_dict,
                abolish_all_scheduling_limits=_abolish_for_task,
                dispatch_interval_mirror=dispatch_interval_mirror,
                window_left_cache=window_left_cache,
            ):
                continue
        _tid_rp = str(task.get("task_id") or "").strip()
        if (
            task.get("roll_pipeline_inspection") or task.get("roll_pipeline_rewind")
        ) and (
            (
                pipeline_room_cache.get(_tid_rp)
                if pipeline_room_cache is not None and _tid_rp in pipeline_room_cache
                else _roll_pipeline_inspection_assign_room(task_queue, _tid_rp)
            )
            <= 1e-12
        ):
            continue
        _elig_acc_window += time_module.perf_counter() - _t_ew0
        _t_ee0 = time_module.perf_counter()
        machine = task["machine"]
        eq_line = str(
            task.get("equipment_line_key") or machine or ""
        ).strip() or machine
        _mocc_trial = _machine_occupancy_key_resolve(task, eq_line)
        if PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE:
            _b1_holder = _exclusive_b1_inspection_holder_for_machine(
                task_queue,
                _mocc_trial,
            )
            if _b1_holder is not None and _b1_holder is not task:
                continue
            from planning_core.core.plan_input import dispatch_trial_order_key_from_task

            _my_dispatch_ord = dispatch_trial_order_key_from_task(task)
        if _equipment_line_lower_dispatch_trial_still_pending(
            task_queue,
            _mocc_trial,
            _my_dispatch_ord,
            current_date,
            daily_status=daily_status,
            members=members,
            machine_avail_dt=machine_avail_dt,
            machine_day_start=machine_day_start,
            machine_handoff=machine_handoff,
            skills_dict=skills_dict,
            abolish_all_scheduling_limits=_abolish_for_task,
            dispatch_interval_mirror=dispatch_interval_mirror,
            assign_probe_ctx=assign_probe_ctx,
            pending_by_occ=pending_by_occ,
            window_left_cache=window_left_cache,
            candidate_task=task,
        ):
            continue
        _elig_acc_equip += time_module.perf_counter() - _t_ee0
        out.append(task)
    _dispatch_loop_profile_add("eligible_early", _elig_acc_early)
    _dispatch_loop_profile_add("eligible_rules", _elig_acc_rules)
    _dispatch_loop_profile_add("eligible_window", _elig_acc_window)
    _dispatch_loop_profile_add("eligible_equip_line", _elig_acc_equip)
    return out
def _assign_one_roll_trial_order_flow(
    task: dict,
    current_date: date,
    daily_status: dict,
    avail_dt: dict,
    machine_avail_dt: dict,
    task_queue: list,
    skills_dict: dict,
    members: list,
    req_map: dict,
    need_rules: list,
    surplus_map: dict,
    global_priority_override: dict,
    macro_run_date: date,
    macro_now_dt: datetime,
    preferred_team: tuple | None,
    _need_headcount_logged_orders: set,
    team_combo_presets: dict | None = None,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
    limited_equipment_mirror: DispatchIntervalMirror | None = None,
    machine_handoff: dict | None = None,
    timeline_events: list | None = None,
    stage35_overtime_only: bool = False,
) -> dict | None:
    """
    1ロール分の最良フォームを決定れる。設備空し・日開始下限を team_start に織り込む。
    preferred_team は与ごられ」かつ「同一日内の直後ロール」として成立れれみ」
    組み合わせ探索より優先して採用する（翌日には挝う越さない）。
    戻り値: team(tuple), start_dt, end_dt, breaks, eff, op, eff_time_per_unit, extra_max, rq_base, need_src_line, extra_src_line, machine, machine_name, eq_line, req_num, max_team_size
    """
    machine = task["machine"]
    machine_name = str(task.get("machine_name", "") or "").strip()
    machine_proc = str(machine or "").strip()
    eq_line = str(task.get("equipment_line_key") or machine or "").strip() or machine
    machine_occ_key = _machine_occupancy_key_resolve(task, eq_line)
    _gpo = global_priority_override or {}
    _all_limits_abolished = bool(_gpo.get("abolish_all_scheduling_limits"))
    _equipment_occupancy_abolished = (
        _new_dispatch_scheduling_limits_abolished(_gpo, task)
    )
    _mh = machine_handoff or {
        "last_tid": {},
        "last_eq": {},
        "last_process": {},
        "started_today": set(),
        "machining_today_occ": set(),
        "last_machining_dt": {},
        "last_machining_date": {},
        "last_lead_op": {},
        "last_machining_sub": {},
    }
    _eod_cont_exempt = _eod_same_request_continuation_exempt(
        machine_occ_key, task, _mh
    )

    plan_ro = _plan_sheet_required_op_optional(task)
    need_src_line = ""
    # need 条件は元依頼NO（rule_task_id）で照合する（枝番タスクは親で判定）。
    _need_tid = _rule_task_id(task)
    if TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY:
        req_num, need_src_line = resolve_need_required_op_explain(
            machine,
            machine_name,
            _need_tid,
            req_map,
            need_rules,
        )
        if plan_ro is not None and plan_ro != req_num:
            need_src_line = (need_src_line + "；") if need_src_line else ""
            need_src_line += f"計画シート必須人数{plan_ro}は未使用（need基準={req_num}）"
    else:
        if plan_ro is not None:
            req_num = plan_ro
            need_src_line = f"計画シート「必須OP(上書)」={req_num}"
        else:
            req_num, need_src_line = resolve_need_required_op_explain(
                machine,
                machine_name,
                _need_tid,
                req_map,
                need_rules,
            )
    if _gpo.get("ignore_need_minimum"):
        req_num = 1
        need_src_line = (
            (need_src_line + " → ") if need_src_line else ""
        ) + "メイン上書ignore_need_minimumでreq=1"

    # -------------------------------------------------------------------
    # 特別ルール L3（EC×EC機 湖南）: 製品名に NR28 を含む場合は 3名配台が必要
    # -------------------------------------------------------------------
    _l3_prod = str(task.get(TASK_COL_PRODUCT) or "")
    if (
        _normalize_process_name_for_rule_match(machine_proc)
        == _normalize_process_name_for_rule_match("EC")
        and _normalize_equipment_match_key(machine_name)
        == _normalize_equipment_match_key("EC機　湖南")
        and "NR28" in unicodedata.normalize("NFKC", _l3_prod)
    ):
        if req_num < 3:
            req_num = 3
            need_src_line = (need_src_line + " → ") if need_src_line else ""
            need_src_line += "特別ルールL3(NR28)で必須人数=3"

    # -------------------------------------------------------------------
    # 特別ルール L7（SEC×SEC機 湖南）: 依頼NOに「PN」が含まれている場合はOPが2名でも配台可能
    # ＝最低人数（req_num）を 2 まで緩和（他の強い必須条件がある場合はそちらを優先）
    # -------------------------------------------------------------------
    _l7_tid_nfkc = unicodedata.normalize("NFKC", str(task.get("task_id") or ""))
    if (
        _normalize_process_name_for_rule_match(machine_proc)
        == _normalize_process_name_for_rule_match("SEC")
        and _normalize_equipment_match_key(machine_name)
        == _normalize_equipment_match_key("SEC機　湖南")
        and "PN" in _l7_tid_nfkc
    ):
        if req_num > 2:
            req_num = 2
            need_src_line = (need_src_line + " → ") if need_src_line else ""
            need_src_line += "特別ルールL7(PN)で必須人数を2まで緩和"

    # -------------------------------------------------------------------
    # 特別ルール L2（スライス×スライス機1 湖南）:
    # ロール単位長さ=100m のときは原則 3 名で配台し、成立しない場合のみ速度 20m/分へフォールバック。
    # - remaining_units / unit_m はロール本数・ロール長(m)（既存仕様）
    # - 速度フォールバックを一度採用したタスクは以降も speed=20 を維持する（探索の無駄を避ける）
    # -------------------------------------------------------------------
    _l2_target = False
    try:
        _unit_m_i = int(task.get("unit_m") or 0)
    except (TypeError, ValueError):
        _unit_m_i = 0
    if _unit_m_i == 100:
        _p = _normalize_process_name_for_rule_match(machine_proc)
        _m = _normalize_equipment_match_key(machine_name)
        if (
            _p == _normalize_process_name_for_rule_match("スライス")
            and _m == _normalize_equipment_match_key("スライス機1　湖南")
        ):
            _l2_target = True
    _l2_req_before = int(req_num) if req_num is not None else 1
    _l2_enforced = False
    if _l2_target and not task.get("_special_l2_speed20_fallback"):
        if req_num < 3:
            req_num = 3
            _l2_enforced = True

    skill_meta_cache: dict = {}

    def skill_role_priority(mem):
        if _new_dispatch_skill_requirements_ignored(_gpo, task):
            return ("OP", 100)
        if mem not in skill_meta_cache:
            srow = skills_dict.get(mem, {})
            v = ""
            if machine_proc and machine_name:
                v = srow.get(f"{machine_proc}+{machine_name}", "")
            elif machine_name:
                v = srow.get(machine_name, "")
            elif machine_proc:
                v = srow.get(machine_proc, "")
            skill_meta_cache[mem] = parse_op_as_skill_cell(v)
        return skill_meta_cache[mem]

    capable_members = [m for m in avail_dt if skill_role_priority(m)[0] in ("OP", "AS")]
    capable_members.sort(key=lambda mm: (skill_role_priority(mm)[1], mm))
    capable_members = _filter_capable_members_b2_disjoint_teams(
        task, task_queue, capable_members
    )

    op_today = [m for m in capable_members if skill_role_priority(m)[0] == "OP"]
    limited_constraints = _new_dispatch_limited_operator_constraints(
        task,
        members,
        skill_role_priority,
        capable_members,
    )
    if limited_constraints is not None:
        req_num = limited_constraints["required_count"]
        need_src_line = (
            f"計画シート「{PLAN_COL_LIMITED_OP}」で必須人数={req_num}"
        )
        capable_members = limited_constraints["capable_members"]
        op_today = [
            m for m in capable_members if skill_role_priority(m)[0] == "OP"
        ]

    _gdp_must, _gdp_warns = _active_global_day_process_must_include(
        _gpo, task, current_date, capable_members, members
    )

    for _gw in _gdp_warns:
        logging.warning(_gw)
    fixed_team_anchor = (
        limited_constraints["fixed_team"]
        if limited_constraints is not None
        else list(_gdp_must)
    )
    if _gdp_must:
        logging.info(
            "メイングローバル(日付×工程): task=%s date=%s 工程=%s フォーム必須=%s",
            task.get("task_id"),
            current_date,
            _log_plain_label(machine),
            ",".join(_gdp_must),
        )
    if fixed_team_anchor:
        _nfix = len(fixed_team_anchor)
        if _nfix > req_num:
            need_src_line = (need_src_line + " → ") if need_src_line else ""
            need_src_line += f"グローバル(日付×工程)指定で最低{_nfix}人"
        req_num = max(req_num, _nfix)

    if limited_constraints is not None:
        extra_max_sheet = 0
        extra_src_line = f"計画シート「{PLAN_COL_LIMITED_OP}」で追加人数=0"
    else:
        extra_max_sheet, extra_src_line = resolve_need_surplus_extra_max_explain(
            machine,
            machine_name,
            task["task_id"],
            surplus_map,
            need_rules,
        )
    if TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW:
        extra_max_sheet = 0
        extra_src_line = (
            (extra_src_line + " → ") if extra_src_line else ""
        ) + "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROWで0"
    extra_max = (
        extra_max_sheet if TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS else 0
    )
    if (
        extra_max_sheet > 0
        and not TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
        and not TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW
    ):
        extra_src_line = (
            (extra_src_line + " → ") if extra_src_line else ""
        ) + "メインは基本人数のみ（余力枠は全配台後に未割当×スキルで追記）"
    max_team_size = min(req_num + extra_max, len(capable_members))
    if max_team_size < req_num:
        max_team_size = req_num
    rq_base = max(1, int(req_num))
    combo_key_assign = (
        f"{machine_proc}+{machine_name}"
        if machine_proc and machine_name
        else ""
    )
    preset_rows_assign = (
        limited_constraints["preset_rows"]
        if limited_constraints is not None
        else (
            (team_combo_presets or {}).get(combo_key_assign)
            if (team_combo_presets and combo_key_assign)
            else None
        )
    )

    # 組み合わせ表プリセットが存在する工程+機械は、探索で表外メンバーを混ぜず
    # プリセット記載メンバー（＋グローバル日付×工程指定の明示メンバー）に限定する。
    if TEAM_ASSIGN_COMBO_SHEET_RESTRICT_TO_PRESET_MEMBERS and preset_rows_assign:
        _allowed_members: set = set()
        for _prio, _req, _preset_team, _cid in preset_rows_assign:
            for _m in _preset_team:
                _allowed_members.add(_m)
        for _m in (fixed_team_anchor or []):
            _allowed_members.add(_m)
        _restricted = [m for m in capable_members if m in _allowed_members]
        _dropped = [m for m in capable_members if m not in _allowed_members]
        if _dropped:
            logging.info(
                "組み合わせ表限定: %s/%s は表記載メンバーのみに配台探索を限定（除外=%s）。",
                _log_plain_label(machine),
                _log_plain_label(machine_name),
                ",".join(_log_plain_label(m) for m in _dropped),
            )
        capable_members = _restricted
        op_today = [
            m for m in capable_members if skill_role_priority(m)[0] == "OP"
        ]
        max_team_size = min(req_num + extra_max, len(capable_members))
        if max_team_size < req_num:
            max_team_size = req_num

    _dto_head = task.get("dispatch_trial_order")
    if _dto_head is not None and _dto_head not in _need_headcount_logged_orders:
        _need_headcount_logged_orders.add(_dto_head)
        logging.info(
            "need人数(試行順優先フロー) order=%s task=%s 工程/機械=%s/%s "
            "req_num=%s [%s] extra_max=%s [%s] max_team候補=%s capable=%s人",
            _dto_head,
            task["task_id"],
            _log_plain_label(machine),
            _log_plain_label(machine_name),
            req_num,
            need_src_line,
            extra_max,
            extra_src_line,
            max_team_size,
            len(capable_members),
        )

    day_floor = _trial_order_flow_day_start_floor(
        task, current_date, macro_run_date, macro_now_dt, task_queue
    )
    machine_day_floor = datetime.combine(current_date, DEFAULT_START_TIME)
    b2_insp_ec_floor: datetime | None = None
    _tid_assign = str(task.get("task_id") or "").strip()
    _trace_assign_enabled = _trace_schedule_task_enabled(_tid_assign)
    def _trace_assign(msg: str, *args) -> None:
        if not _trace_assign_enabled:
            return
        _log_dispatch_trace_schedule(
            _tid_assign,
            "[配台トレース task=%s] " + msg,
            _tid_assign,
            *args,
        )
    if (
        (
            task.get("roll_pipeline_inspection")
            or task.get("roll_pipeline_rewind")
        )
        and _task_queue_has_roll_pipeline_ec_for_tid(task_queue, _tid_assign)
    ):
        b2_insp_ec_floor = _roll_pipeline_b2_inspection_ec_completion_floor_dt(
            task_queue, _tid_assign
        )

    _base_time_per_unit = float(task.get("base_time_per_unit") or 0.0)

    _prev_mach_before_co = machine_avail_dt.get(
        machine_occ_key, machine_day_floor
    )
    _mach_floor_eff, _co_segs, _co_abort = _resolve_machine_changeover_floor_segments(
        abolish_all_scheduling_limits=_equipment_occupancy_abolished,
        machine_occ_key=machine_occ_key,
        task_id=str(task.get("task_id") or "").strip(),
        eq_line=eq_line,
        machine_name=machine_name,
        machine_proc=machine_proc,
        machine_avail_dt=machine_avail_dt,
        machine_day_floor=machine_day_floor,
        current_date=current_date,
        machine_handoff=_mh,
        daily_status=daily_status,
        skills_dict=skills_dict,
        dispatch_interval_mirror=dispatch_interval_mirror,
        timeline_events=timeline_events,
        task_queue=task_queue,
        avail_dt=avail_dt,
    )
    if _co_abort:
        return None

    _l10_slit_nth_end_floor_dt = _l10_sec_start_floor_from_slit_timeline(
        task, timeline_events, task_queue
    )
    _b6_conn_nth_end_floor_dt = _b6_sec_start_floor_from_connection_timeline(
        task, timeline_events, task_queue
    )
    _sec_pair_gate_floor_dt: datetime | None = None
    _pair_gate_candidates = [
        x
        for x in (_l10_slit_nth_end_floor_dt, _b6_conn_nth_end_floor_dt)
        if x is not None
    ]
    if _pair_gate_candidates:
        _sec_pair_gate_floor_dt = max(_pair_gate_candidates)

    def _one_roll_from_team(
        team: tuple,
        min_n: int | None = None,
        max_n: int | None = None,
    ) -> dict | None:
        lo = req_num if min_n is None else min_n
        hi = max_team_size if max_n is None else max_n
        if len(team) < lo or len(team) > hi:
            _trace_assign(
                "候補坴下: フォーム人数外 team=%s size=%s req=%s max=%s",
                ",".join(str(x) for x in team),
                len(team),
                lo,
                hi,
            )
            return None
        op_list = [m for m in team if skill_role_priority(m)[0] == "OP"]
        if not op_list:
            _trace_assign(
                "候補坴下: OP丝在 team=%s",
                ",".join(str(x) for x in team),
            )
            return None
        if not all(m in daily_status for m in team):
            _trace_assign(
                "候補坴下: 当日勤怠キーなし team=%s",
                ",".join(str(x) for x in team),
            )
            return None
        team_start = max(avail_dt[m] for m in team)
        if not _equipment_occupancy_abolished:
            machine_free_dt = _mach_floor_eff
            if team_start < machine_free_dt:
                team_start = machine_free_dt
        if not _all_limits_abolished:
            if team_start < day_floor:
                team_start = day_floor
        if b2_insp_ec_floor is not None and team_start < b2_insp_ec_floor:
            team_start = b2_insp_ec_floor
        if (
            _sec_pair_gate_floor_dt is not None
            and team_start < _sec_pair_gate_floor_dt
        ):
            team_start = _sec_pair_gate_floor_dt
        if stage35_overtime_only:
            ot_floor = _stage35_overtime_regular_end_floor(
                team, daily_status, current_date
            )
            if ot_floor is None:
                _trace_assign(
                    "候補棄却: 残業帯開始下限なし team=%s",
                    ",".join(str(x) for x in team),
                )
                return None
            if team_start < ot_floor:
                team_start = ot_floor
            team_end_limit = min(daily_status[m]["end_dt"] for m in team)
        else:
            team_end_limit = min(daily_status[m]["end_dt"] for m in team)
            team_end_limit = _interactive_trial_relax_team_end_limit_to_eod(
                team_end_limit, current_date
            )
        if team_start >= team_end_limit:
            _trace_assign(
                "候補坴下: 開始>=終業 team=%s start=%s end_limit=%s",
                ",".join(str(x) for x in team),
                team_start,
                team_end_limit,
            )
            return None
        team_breaks = []
        for m in team:
            team_breaks.extend(daily_status[m]["breaks_dt"])
        team_breaks = merge_time_intervals(team_breaks)

        avg_eff = sum(daily_status[m]["efficiency"] for m in team) / len(team)
        if avg_eff <= 0:
            avg_eff = 0.01
        t_eff = parse_float_safe(task.get("task_eff_factor"), 1.0)
        if t_eff <= 0:
            t_eff = 1.0
        eff_time_per_unit = (
            _base_time_per_unit
            / avg_eff
            / t_eff
            * _surplus_team_time_factor(rq_base, len(team), extra_max)
        )
        _defer_min_contig = max(1, int(math.ceil(float(eff_time_per_unit))))

        def _refloor_trial_roll(ts: datetime) -> datetime:
            ts = max(ts, max(avail_dt[m] for m in team))
            if not _equipment_occupancy_abolished:
                mf = _mach_floor_eff
                if ts < mf:
                    ts = mf
            if not _all_limits_abolished:
                if ts < day_floor:
                    ts = day_floor
            if b2_insp_ec_floor is not None and ts < b2_insp_ec_floor:
                ts = b2_insp_ec_floor
            if (
                _sec_pair_gate_floor_dt is not None
                and ts < _sec_pair_gate_floor_dt
            ):
                ts = _sec_pair_gate_floor_dt
            if stage35_overtime_only:
                ot_fl = _stage35_overtime_regular_end_floor(
                    team, daily_status, current_date
                )
                if ot_fl is not None and ts < ot_fl:
                    ts = ot_fl
            return ts

        team_start_d = _defer_team_start_past_prebreak_and_end_of_day(
            task,
            team,
            team_start,
            team_end_limit,
            team_breaks,
            _refloor_trial_roll,
            min_contiguous_work_mins=_defer_min_contig,
            eod_same_request_continuation_exempt=_eod_cont_exempt,
        )
        if team_start_d is None:
            _trace_assign(
                "候補坴下: 休憩帯内・終業直後(尝残)で当日試行 team=%s",
                ",".join(str(x) for x in team),
            )
            return None
        team_start = team_start_d
        _roll_prep_extra: list[dict] = []
        if not _equipment_occupancy_abolished:
            team_start, _roll_prep_extra = _roll_prep_segments_for_assign(
                team_start=team_start,
                team_breaks=team_breaks,
                machine_handoff=_mh,
                machine_occ_key=machine_occ_key,
                current_date=current_date,
                task_id=str(task.get("task_id") or "").strip(),
                machine_proc=str(machine_proc or "").strip(),
                machine_name=str(machine_name or "").strip(),
                eq_line=eq_line,
                abolish_limits=False,
                prev_machining_end=_resolve_prev_machining_end_for_roll_prep(
                    _mh,
                    machine_occ_key,
                    (_mh.get("last_machining_dt") or {}).get(machine_occ_key),
                    machine_avail_dt,
                    machine_day_floor,
                    _prev_mach_before_co,
                ),
                prev_eq_line=str(
                    (_mh.get("last_eq") or {}).get(machine_occ_key, "") or ""
                ).strip(),
                machine_avail_dt=machine_avail_dt,
                machine_day_floor=machine_day_floor,
            )
            team_start = _refloor_trial_roll(team_start)
        protected_capacity = _candidate_capacity_after_equipment_protection(
            limited_equipment_mirror,
            machine_occ_key,
            team_start,
            float(max(1, int(math.ceil(eff_time_per_unit)))),
            1.0,
            team_breaks,
            team_end_limit,
        )
        if protected_capacity is None:
            return None
        team_start, _ = protected_capacity
        if team_start >= team_end_limit:
            _trace_assign(
                "候補坴下: デファー後に開始>=終業 team=%s start=%s end_limit=%s",
                ",".join(str(x) for x in team),
                team_start,
                team_end_limit,
            )
            return None
        if dispatch_interval_mirror is not None and _roll_prep_extra:
            for _pseg in _roll_prep_extra:
                _pst = _pseg.get("start_dt")
                _ped = _pseg.get("end_dt")
                _pok = str(
                    _pseg.get("machine_occupancy_key") or machine_occ_key
                ).strip()
                if (
                    isinstance(_pst, datetime)
                    and isinstance(_ped, datetime)
                    and _pok
                    and dispatch_interval_mirror.would_block_equipment(
                        _pok, _pst, _ped
                    )
                ):
                    _trace_assign(
                        "区間ミラー坴下(準備): eq=%s start=%s end=%s",
                        _pok,
                        _pst,
                        _ped,
                    )
                    return None

        _, avail_mins, _ = calculate_end_time(
            team_start, 9999, team_breaks, team_end_limit
        )
        # calculate_end_time への 1 ロール実分は int(eff_time_per_unit)（少なくとも 1 分）と一致させる。
        # float のまま割ると avail≈need の端数で cap=0 になり却下する（ログ: avail_mins=7, eff=7.797）。
        work_mins_needed = max(1, int(eff_time_per_unit))
        _trial_units_cap = int(avail_mins / work_mins_needed)
        if _trial_units_cap < 1:
            _trace_assign(
                "候補坴下: 実僝丝足 team=%s start=%s avail_mins=%s need_mins=%.2f",
                ",".join(str(x) for x in team),
                team_start,
                avail_mins,
                eff_time_per_unit,
            )
            return None
        _rem_eod_ceil = math.ceil(float(task.get("remaining_units") or 0))
        _eod_eff_th = (
            min(int(ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS), int(_rem_eod_ceil))
            if _rem_eod_ceil > 0
            else int(ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS)
        )
        if _eod_reject_capacity_units_below_threshold(
            _trial_units_cap,
            team_start,
            team_end_limit,
            eod_same_request_continuation_exempt=_eod_cont_exempt,
            remaining_units_ceil=_rem_eod_ceil,
        ):
            _trace_assign(
                "候補坴下: 終業直後で当日坎容ロール数は閾値未満 team=%s cap=%s th=%s start=%s",
                ",".join(str(x) for x in team),
                _trial_units_cap,
                _eod_eff_th,
                team_start,
            )
            return None
        _contig = _contiguous_work_minutes_until_next_break_or_limit(
            team_start, team_breaks, team_end_limit
        )
        if _contig < work_mins_needed:
            _trace_assign(
                "候補坴下: 休憩またねのため、連続実僝丝足 team=%s contiguous_min=%s need_mins=%s start=%s",
                ",".join(str(x) for x in team),
                _contig,
                work_mins_needed,
                team_start,
            )
            return None
        actual_end_dt, _, _ = calculate_end_time(
            team_start, work_mins_needed, team_breaks, team_end_limit
        )
        if dispatch_interval_mirror is not None and dispatch_interval_mirror.would_block_roll(
            machine_occ_key, team, team_start, actual_end_dt
        ):
            _trace_assign(
                "区間ミラー坴下: team=%s start=%s end=%s eq=%s",
                ",".join(str(x) for x in team),
                team_start,
                actual_end_dt,
                eq_line,
            )
            return None
        if _new_dispatch_limited_equipment_interval_blocked(
            task,
            _gpo,
            machine_avail_dt,
            limited_equipment_mirror,
            machine_occ_key,
            team_start,
            actual_end_dt,
        ):
            _trace_assign(
                "限定設備保護区間で棄却: team=%s start=%s end=%s eq=%s",
                ",".join(str(x) for x in team),
                team_start,
                actual_end_dt,
                machine_occ_key,
            )
            return None
        lead_op = min(op_list, key=lambda mm: (skill_role_priority(mm)[1], mm))
        team_prio_sum = sum(skill_role_priority(m)[1] for m in team)
        return {
            "team": team,
            "team_start": team_start,
            "actual_end_dt": actual_end_dt,
            "team_breaks": team_breaks,
            "avg_eff": avg_eff,
            "prio_sum": team_prio_sum,
            "op_list": op_list,
            "eff_time_per_unit": eff_time_per_unit,
            "lead_op": lead_op,
            "changeover_segments": list(_co_segs or []) + list(_roll_prep_extra or []),
            "startup_skill_role_priority": skill_role_priority,
        }

    # 特別指定: 同一日・連続ロールは剝回フォームを優先（翌日へは挝う越さない）。
    _hist = task.get("assigned_history") or []
    _last_hist_date = _hist[-1].get("date") if _hist else None
    _same_day_last_roll = _last_hist_date == current_date.strftime("%m/%d")
    _pref_team_size_ok = (
        True
        if TEAM_ASSIGN_COMBO_SHEET_MAY_EXCEED_NEED
        else len(preferred_team or ()) <= max_team_size
    )
    if preferred_team and _same_day_last_roll and _pref_team_size_ok:
        pt = tuple(preferred_team)
        _pref_pt_ok = (not fixed_team_anchor) or all(
            m in pt for m in fixed_team_anchor
        )
        if _pref_pt_ok and all(m in capable_members and m in avail_dt for m in pt):
            got = _one_roll_from_team(pt)
            if got is not None:
                _cid_pt = _lookup_combo_sheet_row_id_for_preset_team(
                    preset_rows_assign, pt
                )
                return {
                    **got,
                    "extra_max": extra_max,
                    "rq_base": rq_base,
                    "need_src_line": need_src_line,
                    "extra_src_line": extra_src_line,
                    "machine": machine,
                    "machine_name": machine_name,
                    "eq_line": eq_line,
                    "req_num": req_num,
                    "max_team_size": max_team_size,
                    "combo_sheet_row_id": _cid_pt,
                    "combo_preset_team": pt if _cid_pt is not None else None,
                    "startup_skill_role_priority": skill_role_priority,
                }

    team_candidates: list[dict] = []
    # 組み合わせ表プリセットは「成立したら坳 return」せう」組み合わせ探索とまとめで
    # team_start / スラック付しタプルで最良を決める（シート上の優先度順は試行順のみ）。
    if preset_rows_assign:
        for _prio, sheet_rs, preset_team, combo_row_id in preset_rows_assign:
            bounds = _combo_preset_team_size_bounds(
                tuple(preset_team), sheet_rs, max_team_size
            )
            if bounds is None:
                continue
            lo_pt, hi_pt = bounds
            if fixed_team_anchor and not all(m in preset_team for m in fixed_team_anchor):
                continue
            if not all(m in capable_members for m in preset_team):
                continue
            if sum(1 for m in preset_team if skill_role_priority(m)[0] == "OP") < 1:
                continue
            got = _one_roll_from_team(
                tuple(preset_team), min_n=lo_pt, max_n=hi_pt
            )
            if got is not None:
                team_candidates.append(
                    {
                        **got,
                        "combo_sheet_row_id": combo_row_id,
                        "combo_preset_team": tuple(preset_team),
                        "combo_preset_priority": _prio,
                    }
                )
    for tsize in range(req_num, max_team_size + 1):
        if fixed_team_anchor:
            _ft = list(fixed_team_anchor)
            others = [m for m in capable_members if m not in _ft]
            need_extra = tsize - len(_ft)
            if need_extra < 0:
                teams_iter = []
            elif need_extra == 0:
                teams_iter = [tuple(_ft)]
            elif len(others) >= need_extra:
                teams_iter = [
                    tuple(_ft + list(rest))
                    for rest in itertools.combinations(others, need_extra)
                ]
            else:
                teams_iter = []
        else:
            teams_iter = itertools.combinations(capable_members, tsize)

        for team in teams_iter:
            got = _one_roll_from_team(team)
            if got is not None:
                team_candidates.append(
                    {
                        **got,
                        "combo_sheet_row_id": None,
                        "combo_preset_team": None,
                        "combo_preset_priority": None,
                    }
                )

    if not team_candidates:
        # L2 フォールバック: 3名が成立しない場合は req を元に戻し、速度20m/分相当へ切替えて再探索。
        if (
            _l2_target
            and _l2_enforced
            and _l2_req_before < 3
            and not task.get("_special_l2_speed20_fallback")
        ):
            _orig_speed = task.get(TASK_COL_SPEED)
            _orig_btpu = task.get("base_time_per_unit")
            try:
                req_num = _l2_fallback_required_count(
                    _l2_req_before, limited_constraints
                )
            except (TypeError, ValueError):
                req_num = 1
            max_team_size = min(req_num + extra_max, len(capable_members))
            if max_team_size < req_num:
                max_team_size = req_num
            rq_base = max(1, int(req_num))
            # unit_m/speed(20m/分) = 1ロールあたり分
            _speed20 = 20.0
            task[TASK_COL_SPEED] = _speed20
            task["base_time_per_unit"] = float(_unit_m_i) / _speed20 if _unit_m_i > 0 else _base_time_per_unit
            _base_time_per_unit = float(task.get("base_time_per_unit") or _base_time_per_unit)

            team_candidates = []
            if preset_rows_assign:
                for _prio, sheet_rs, preset_team, combo_row_id in preset_rows_assign:
                    bounds = _combo_preset_team_size_bounds(
                        tuple(preset_team), sheet_rs, max_team_size
                    )
                    if bounds is None:
                        continue
                    lo_pt, hi_pt = bounds
                    if fixed_team_anchor and not all(m in preset_team for m in fixed_team_anchor):
                        continue
                    if not all(m in capable_members for m in preset_team):
                        continue
                    if sum(1 for m in preset_team if skill_role_priority(m)[0] == "OP") < 1:
                        continue
                    got = _one_roll_from_team(
                        tuple(preset_team), min_n=lo_pt, max_n=hi_pt
                    )
                    if got is not None:
                        team_candidates.append(
                            {
                                **got,
                                "combo_sheet_row_id": combo_row_id,
                                "combo_preset_team": tuple(preset_team),
                            }
                        )
            for tsize in range(req_num, max_team_size + 1):
                if fixed_team_anchor:
                    _ft = list(fixed_team_anchor)
                    others = [m for m in capable_members if m not in _ft]
                    need_extra = tsize - len(_ft)
                    if need_extra < 0:
                        teams_iter = []
                    elif need_extra == 0:
                        teams_iter = [tuple(_ft)]
                    elif len(others) >= need_extra:
                        teams_iter = [
                            tuple(_ft + list(rest))
                            for rest in itertools.combinations(others, need_extra)
                        ]
                    else:
                        teams_iter = []
                else:
                    teams_iter = itertools.combinations(capable_members, tsize)

                for team in teams_iter:
                    got = _one_roll_from_team(team)
                    if got is not None:
                        team_candidates.append(
                            {
                                **got,
                                "combo_sheet_row_id": None,
                                "combo_preset_team": None,
                            }
                        )

            if team_candidates:
                task["_special_l2_speed20_fallback"] = True
            else:
                # 速度切替でも成立しない → もとの値へ戻す
                task[TASK_COL_SPEED] = _orig_speed
                task["base_time_per_unit"] = _orig_btpu
                _base_time_per_unit = float(task.get("base_time_per_unit") or _base_time_per_unit)

        if not team_candidates:
            _mem_max_end: datetime | None = None
            for _m in capable_members:
                if _m not in daily_status:
                    continue
                _ed = daily_status[_m].get("end_dt")
                if isinstance(_ed, datetime):
                    _mem_max_end = (
                        _ed if _mem_max_end is None else max(_mem_max_end, _ed)
                    )
            if (
                len(capable_members) >= req_num
                and _mem_max_end is not None
                and isinstance(_mach_floor_eff, datetime)
                and _mach_floor_eff >= _mem_max_end
            ):
                logging.warning(
                    "段階2: 依頼NO=%s 日付=%s 工程/機械=%s/%s でフォーム候補は0件。"
                    "スキル革坈(OP/AS)は %s 人いしたは」設備の加工開始下限=%s は"
                    "当日の担当候補の退勤(%s)以降のため、この日は割当でしません。"
                    "master「機械カレンダー」で当該日・当該機械列に試行な記入はないか」"
                    "または剝工程の占有で設備下限は終業まで繰り上はっていないか確認してください"
                    "（配台ルール 3.2.1 機械カレンダー・トラブルシュート）。"
                    "参考: changeover剝の設備空し下限=%s 占有キー=%s",
                    task.get("task_id"),
                    current_date,
                    _log_plain_label(machine),
                    _log_plain_label(machine_name),
                    len(capable_members),
                    _mach_floor_eff.strftime("%Y-%m-%d %H:%M"),
                    _mem_max_end.strftime("%H:%M"),
                    _prev_mach_before_co.strftime("%Y-%m-%d %H:%M"),
                    machine_occ_key,
                )
            elif len(capable_members) < int(req_num or 1):
                task["_dispatch_block_no_op_on_working_days"] = True
            elif (
                len(capable_members) >= int(req_num or 1)
                and not op_today
            ):
                task["_dispatch_block_no_op_on_working_days"] = True
                if not task.get("_dispatch_no_op_warned"):
                    task["_dispatch_no_op_warned"] = True
                    _as_only = [
                        m
                        for m in capable_members
                        if skill_role_priority(m)[0] == "AS"
                    ]
                    logging.warning(
                        "段階2: 依頼NO=%s 工程/機械=%s/%s — 当日稼働のスキル候補 %s 人いるが "
                        "OP が 0 人のため 2 人以上編成できません（AS のみ: %s）。"
                        " master「skills」で当該工程×機械に OP を設定するか、"
                        " OP 担当の勤怠（休暇・公休等）を確認してください。",
                        task.get("task_id"),
                        _log_plain_label(machine),
                        _log_plain_label(machine_name),
                        len(capable_members),
                        "、".join(_as_only[:8])
                        + (f" 他{len(_as_only) - 8}人" if len(_as_only) > 8 else ""),
                    )
            _fail_reason = "other_no_team"
            if (
                len(capable_members) >= int(req_num or 1)
                and _mem_max_end is not None
                and isinstance(_mach_floor_eff, datetime)
                and _mach_floor_eff >= _mem_max_end
            ):
                _fail_reason = "mach_floor_after_shift_end"
            elif len(capable_members) < int(req_num or 1):
                _fail_reason = "capable_lt_req"
            elif (
                len(capable_members) >= int(req_num or 1)
                and not op_today
            ):
                _fail_reason = "no_op_on_working_day"
            if limited_constraints is not None:
                if _fail_reason == "mach_floor_after_shift_end":
                    _record_limited_operator_rejection(
                        task, "設備占有により開始可能時刻が選択者の終業後です"
                    )
                elif _fail_reason == "capable_lt_req":
                    _record_limited_operator_rejection(
                        task, "非出勤または二重配台により選択者全員が同時に利用できません"
                    )
                else:
                    _record_limited_operator_rejection(
                        task,
                        "休憩・終業・二重配台・設備占有の安全条件を満たしません",
                    )
            task["_debug_fail_attempts"] = (
                int(task.get("_debug_fail_attempts") or 0) + 1
            )
            task["_debug_last_fail_reason"] = _fail_reason
            task["_debug_last_fail_date"] = str(current_date)
            task["_debug_last_fail_capable"] = len(capable_members)
            task["_debug_last_fail_req_num"] = int(req_num or 1)
            task["_debug_last_fail_op_today"] = len(op_today)
            _interactive_append_team_shortage_op_as(
                task,
                current_date,
                machine,
                machine_name,
                capable_members,
                int(req_num) if req_num is not None else 1,
            )
            return None
    t_min = min(c["team_start"] for c in team_candidates)

    def _team_cand_key(c):
        _base = _team_assignment_sort_tuple(
            c["team"],
            c["team_start"],
            1,
            c["prio_sum"],
            t_min,
        )
        # シート「組合せ優先度」は試行順だけでなく、同条件のタイブレークにも使う
        # （開始・人数・短縮数などが同じなら、優先度が小さいプリセットを優先）
        _combo_rank = c.get("combo_preset_priority")
        try:
            _combo_rank_i = int(_combo_rank) if _combo_rank is not None else 10**9
        except (TypeError, ValueError):
            _combo_rank_i = 10**9
        try:
            if isinstance(_base, tuple) and len(_base) >= 1:
                return _base[:-1] + (_combo_rank_i, _base[-1])
        except Exception:
            pass
        return (_combo_rank_i, _base)

    best_c = min(team_candidates, key=_team_cand_key)
    if best_c.get("combo_sheet_row_id") is None and preset_rows_assign:
        _lcid = _lookup_combo_sheet_row_id_for_preset_team(
            preset_rows_assign, tuple(best_c["team"])
        )
        if _lcid is not None:
            best_c = {
                **best_c,
                "combo_sheet_row_id": _lcid,
                "combo_preset_team": tuple(best_c["team"]),
            }
    return {
        **best_c,
        "extra_max": extra_max,
        "rq_base": rq_base,
        "need_src_line": need_src_line,
        "extra_src_line": extra_src_line,
        "machine": machine,
        "machine_name": machine_name,
        "eq_line": eq_line,
        "req_num": req_num,
        "max_team_size": max_team_size,
        "combo_sheet_row_id": best_c.get("combo_sheet_row_id"),
        "combo_preset_team": best_c.get("combo_preset_team"),
        "startup_skill_role_priority": skill_role_priority,
    }
def _trial_order_hard_precheck_blocks_assign_probe(task: dict, task_queue: list) -> bool:
    """
    `_trial_order_flow_eligible_tasks` で `_task_blocked_by_global_dispatch_trial_order`
    より前に適用される条件のうち、1 ロール assign プローブ単体では表れないもの。
    これをプローブ前に適用し」同一依頼依存や B1 占有等で先頭 dto が実際には候補に入れないとき
    `_effective_min_dispatch_trial_order_from_pool` が dto=1 に張り付し、他行が global_dto で
    全日止まる膠着を防し。
    """
    if float(task.get("remaining_units") or 0) <= 1e-12:
        return True
    wip_ec_before_insp = None
    if isinstance(WIP_LIMIT_EC_BEFORE_INSP_ROLLS, int) and WIP_LIMIT_EC_BEFORE_INSP_ROLLS > 0:
        _m = WIP_LIMIT_EC_BEFORE_INSP_AGGREGATE_MODE
        if _wip_ec_l11_aggregate_is_global():
            wip_ec_before_insp = _wip_ec_before_insp_roll_count(task_queue)
        elif _m == "task_id":
            _tid = str(task.get("task_id") or "").strip()
            wip_ec_before_insp = _wip_ec_before_insp_roll_count(
                task_queue, task_id_exact=_tid
            )
        else:
            _bk = _wip_l11_bucket_key_for_task_id(str(task.get("task_id") or ""))
            wip_ec_before_insp = _wip_ec_before_insp_roll_count(
                task_queue, task_id_head=_bk
            )

    wip_slit_before_sec = None
    if (
        isinstance(WIP_LIMIT_SLIT_BEFORE_SEC_ROLLS, int)
        and WIP_LIMIT_SLIT_BEFORE_SEC_ROLLS > 0
    ):
        slit_done_total = 0.0
        sec_done_total = 0.0
        _slit_proc = _normalize_process_name_for_rule_match(SPECIAL_WIP_SLIT_PROCESS)
        _slit_mach = _normalize_equipment_match_key(SPECIAL_WIP_SLIT_MACHINE)
        _sec_proc = _normalize_process_name_for_rule_match(SPECIAL_WIP_SEC_PROCESS)
        _sec_mach = _normalize_equipment_match_key(SPECIAL_WIP_SEC_MACHINE)
        for _t in task_queue:
            proc = _normalize_process_name_for_rule_match(_t.get("machine"))
            mach = _normalize_equipment_match_key(_t.get("machine_name"))
            if not proc or not mach:
                continue
            init = float(_t.get("initial_remaining_units") or 0)
            rem = float(_t.get("remaining_units") or 0)
            done = max(0.0, init - rem)
            if done <= 1e-12:
                continue
            if proc == _slit_proc and mach == _slit_mach:
                slit_done_total += done
            elif proc == _sec_proc and mach == _sec_mach:
                sec_done_total += done
        wip_slit_before_sec = max(0.0, slit_done_total - sec_done_total)

    wip_connection_before_sec = None
    if (
        isinstance(WIP_LIMIT_CONNECTION_BEFORE_SEC_ROLLS, int)
        and WIP_LIMIT_CONNECTION_BEFORE_SEC_ROLLS > 0
    ):
        connection_done_total = 0.0
        sec_done_c = 0.0
        _conn_proc = _normalize_process_name_for_rule_match(
            SPECIAL_WIP_CONNECTION_PROCESS
        )
        _conn_mach = _normalize_equipment_match_key(SPECIAL_WIP_CONNECTION_MACHINE)
        _sec_proc_c = _normalize_process_name_for_rule_match(SPECIAL_WIP_SEC_PROCESS)
        _sec_mach_c = _normalize_equipment_match_key(SPECIAL_WIP_SEC_MACHINE)
        for _t in task_queue:
            proc = _normalize_process_name_for_rule_match(_t.get("machine"))
            mach = _normalize_equipment_match_key(_t.get("machine_name"))
            if not proc or not mach:
                continue
            init = float(_t.get("initial_remaining_units") or 0)
            rem = float(_t.get("remaining_units") or 0)
            done = max(0.0, init - rem)
            if done <= 1e-12:
                continue
            if proc == _conn_proc and mach == _conn_mach:
                connection_done_total += done
            elif proc == _sec_proc_c and mach == _sec_mach_c:
                sec_done_c += done
        wip_connection_before_sec = max(0.0, connection_done_total - sec_done_c)

    if (
        wip_ec_before_insp is not None
        and wip_ec_before_insp >= float(WIP_LIMIT_EC_BEFORE_INSP_ROLLS)
        and task.get("roll_pipeline_ec")
        and _b2_ec_insp_pair_in_queue(
            task_queue, str(task.get("task_id") or "").strip()
        )
    ):
        return True
    if wip_slit_before_sec is not None and wip_slit_before_sec >= float(
        WIP_LIMIT_SLIT_BEFORE_SEC_ROLLS
    ):
        proc = _normalize_process_name_for_rule_match(task.get("machine"))
        mach = _normalize_equipment_match_key(task.get("machine_name"))
        if (
            proc == _normalize_process_name_for_rule_match(SPECIAL_WIP_SLIT_PROCESS)
            and mach == _normalize_equipment_match_key(SPECIAL_WIP_SLIT_MACHINE)
            and _task_on_slit_sec_process_path(task)
        ):
            return True
    if wip_connection_before_sec is not None and wip_connection_before_sec >= float(
        WIP_LIMIT_CONNECTION_BEFORE_SEC_ROLLS
    ):
        proc = _normalize_process_name_for_rule_match(task.get("machine"))
        mach = _normalize_equipment_match_key(task.get("machine_name"))
        if (
            proc
            == _normalize_process_name_for_rule_match(SPECIAL_WIP_CONNECTION_PROCESS)
            and mach
            == _normalize_equipment_match_key(SPECIAL_WIP_CONNECTION_MACHINE)
        ):
            return True
    if _l10_b41_sec_blocked_by_slit_min_rolls(task, task_queue):
        return True
    if _b61_sec_blocked_by_connection_min_rolls(task, task_queue):
        return True
    if _task_blocked_by_same_request_dependency(task, task_queue):
        return True
    if (
        task.get("roll_pipeline_inspection") or task.get("roll_pipeline_rewind")
    ) and (
        _roll_pipeline_inspection_assign_room(
            task_queue, str(task.get("task_id", "") or "").strip()
        )
        <= 1e-12
    ):
        return True
    machine = task["machine"]
    eq_line = str(
        task.get("equipment_line_key") or machine or ""
    ).strip() or machine
    _mocc_trial = _machine_occupancy_key_resolve(task, eq_line)
    if PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE:
        _b1_holder = _exclusive_b1_inspection_holder_for_machine(task_queue, _mocc_trial)
        if _b1_holder is not None and _b1_holder is not task:
            return True
    return False
def _trial_order_assign_probe_fails(
    task: dict,
    current_date: date,
    daily_status: dict,
    ctx: dict,
) -> bool:
    """
    睾在の avail_dt / machine_avail_dt / machine_handoff のスナップショットで
    `_assign_one_roll_trial_order_flow` は None になるなら True。
    機械枠は坝分でも人・休憩・ミラー等で詰まり」グローバル試行順の値は先頭行に張り付しのを防し。
    副作用なし（need 人数ログ用 set は毎回空）。
    """
    if _trial_order_hard_precheck_blocks_assign_probe(task, ctx["task_queue"]):
        return True
    try:
        _t_probe0 = time_module.perf_counter()
        r = _assign_one_roll_trial_order_flow(
            task,
            current_date,
            daily_status,
            ctx["avail_dt"],
            ctx["machine_avail_dt"],
            ctx["task_queue"],
            ctx["skills_dict"],
            ctx["members"],
            ctx["req_map"],
            ctx["need_rules"],
            ctx["surplus_map"],
            ctx["global_priority_override"],
            ctx["macro_run_date"],
            ctx["macro_now_dt"],
            None,
            set(),
            team_combo_presets=ctx.get("team_combo_presets"),
            dispatch_interval_mirror=ctx.get("dispatch_interval_mirror"),
            limited_equipment_mirror=ctx.get("limited_equipment_mirror"),
            machine_handoff=ctx["machine_handoff"],
        )
        _dispatch_loop_profile_add(
            "assign_probe", time_module.perf_counter() - _t_probe0
        )
    except Exception as ex:
        logging.warning(
            "trial_order_assign_probe 例外のため、当該行は除外しない: task=%s err=%s",
            task.get("task_id"),
            ex,
        )
        return False
    return r is None
def _interactive_stage3_unmet_cap_m_on_date(
    task: dict,
    current_date: date,
    *,
    interactive_dispatch_targets: dict | None,
    interactive_trial_meters_done: dict | None,
) -> float:
    """
    段階3: 当該暦日の JSON 目標（interactive_dispatch_targets）に対する未達 m。
    同一機械・同一日で複数依頼があるとき、暦日キャップ未達を大きい順に先に割り当てるための sort 用。
    段階2同一パリティではループ内キャップを使わないため常に 0。
    """
    if not _interactive_dispatch_cap_enforced_in_schedule_loop():
        return 0.0
    if not interactive_dispatch_targets:
        return 0.0
    tid = _interactive_norm_cell(str(task.get("task_id") or ""))
    proc = _interactive_dispatch_target_process_key(task.get("machine"))
    mach = _interactive_norm_cell(str(task.get("machine_name") or ""))
    if not tid or not proc or not mach:
        return 0.0
    cap_key = (tid, proc, mach, current_date)
    try:
        cap_m = float(interactive_dispatch_targets.get(cap_key, 0.0))
    except (TypeError, ValueError):
        cap_m = 0.0
    if cap_m <= 1e-9:
        return 0.0
    done_dict = interactive_trial_meters_done or {}
    try:
        done_m = float(done_dict.get(cap_key, 0.0))
    except (TypeError, ValueError):
        done_m = 0.0
    return max(0.0, cap_m - done_m)
def _interactive_stage3_skip_macro_now_start_floor(
    task: dict,
    current_date: date,
) -> bool:
    """
    段階3: JSON 暦日キャップが未達の依頼は、data_extract 由来の macro_now_dt で
    当日開始下限を切り上げない（計画暦日どおり配台試行する）。
    """
    targets = _PLAN_IMPL_INTERACTIVE_DISPATCH_TARGETS
    meters_done = _PLAN_IMPL_INTERACTIVE_TRIAL_METERS_DONE
    if not (_interactive_dispatch_trial_env_active() and targets):
        return False
    if (
        _interactive_stage3_unmet_cap_m_on_date(
            task,
            current_date,
            interactive_dispatch_targets=targets,
            interactive_trial_meters_done=meters_done,
        )
        > 1e-9
    ):
        return True
    return False
def _interactive_trial_meters_done_by_timeline_calendar_date(
    timeline_events: list,
    task_queue: list,
) -> dict[tuple[str, str, str, date], float]:
    """
    タイムライン加工イベントの暦日をそのままキーに換算 m を集計する（cap_key 解決なし）。
    段階3の暦日未達一覧・実配台数量表示はこちらを正とする。
    """
    acc: dict[tuple[str, str, str, date], float] = {}
    _mach_evs = [e for e in (timeline_events or []) if _is_machining_timeline_event(e)]

    def _ev_sort_key(ev: dict):
        st = ev.get("start_dt")
        if isinstance(st, datetime):
            return st
        return datetime.max.replace(tzinfo=None)

    _mach_evs.sort(key=_ev_sort_key)
    for ev in _mach_evs:
        tid = _interactive_norm_cell(ev.get("task_id"))
        ev_machine = str(ev.get("machine") or "").strip()
        tsk = _resolve_task_dict_for_timeline_line(tid, ev_machine, task_queue)
        if tsk is None:
            tsk = next(
                (
                    t
                    for t in (task_queue or [])
                    if _interactive_norm_cell(str(t.get("task_id") or "")) == tid
                ),
                None,
            )
        if tsk is None:
            continue
        proc_n = _interactive_dispatch_target_process_key(tsk.get("machine"))
        mach_n = _interactive_norm_cell(tsk.get("machine_name"))
        d = _interactive_timeline_event_calendar_date(ev)
        if d is None or not tid or not proc_n or not mach_n:
            continue
        try:
            ud = float(ev.get("units_done") or 0)
            um = float(ev.get("unit_m") or tsk.get("unit_m") or 0)
        except (TypeError, ValueError):
            continue
        add_m = ud * um
        if add_m <= 1e-18:
            continue
        kk = (tid, proc_n, mach_n, d)
        acc[kk] = acc.get(kk, 0.0) + add_m
    return acc
def _stage2_aladdin_next_day_exclude_consumes_roll(task: dict, current_date: date) -> bool:
    """翌稼働日のみ、アラジン当日除外 m を消費し当該ロール割当をスキップする。"""
    apply_d = _STAGE2_ALADDIN_EXCLUDE_APPLY_DATE
    if apply_d is None or current_date != apply_d:
        return False
    if not task.get("aladdin_today_exclude_next_day_dialog"):
        return False
    try:
        rem = float(task.get("aladdin_next_day_exclude_remaining_m") or 0)
    except (TypeError, ValueError):
        rem = 0.0
    if rem <= 1e-12:
        return False
    try:
        um = float(task.get("unit_m") or 0)
    except (TypeError, ValueError):
        um = 0.0
    if um <= 1e-12:
        return False
    skip_m = min(rem, um)
    task["aladdin_next_day_exclude_remaining_m"] = max(0.0, rem - skip_m)
    logging.info(
        "段階2: アラジン当日・翌日除外を適用 依頼NO=%s 工程=%s 機械名=%s → 除外 %s m（配台日=%s）",
        task.get("task_id"),
        _log_plain_label(task.get("machine")),
        _log_plain_label(task.get("machine_name")),
        skip_m,
        current_date.isoformat(),
    )
    return True
def _trial_order_first_schedule_pass(
    current_date: date,
    tasks_today: list,
    task_queue: list,
    daily_status: dict,
    machine_avail_dt: dict,
    avail_dt: dict,
    timeline_events: list,
    skills_dict: dict,
    members: list,
    req_map: dict,
    need_rules: list,
    surplus_map: dict,
    global_priority_override: dict,
    macro_run_date: date,
    macro_now_dt: datetime,
    _need_headcount_logged_orders: set,
    team_combo_presets: dict | None = None,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
    limited_equipment_mirror: DispatchIntervalMirror | None = None,
    interactive_dispatch_targets: dict | None = None,
    interactive_trial_pair_dates: dict | None = None,
    interactive_trial_meters_done: dict | None = None,
    *,
    cap_drain_only: bool = False,
) -> bool:
    """
    ①当日候補を配台試行順の昇順に並きる（1 パス分）。
    **完全二相（§B-2 / §B-3）**: **フェーズ1**で **後続パイプライン行**（熱融着検査・巻返し）**を除し**候補（EC・他依頼・他工程）を試行順どより
    **`_drain_rolls_for_task`** し、**フェーズ2**は §B-2 検査＝§B-3 巻返し行のみ（**同一依頼の EC は全日で完走した後**に陝り候補化。
    EC 残はある日は `_trial_order_flow_eligible_tasks` で後続を外し、翌稼働日以降も EC のみ剝進れる。
    カレンダー通算で EC 完走後」`_run_b2_inspection_rewind_pass` は日付先頭から後続の値再走査れる）。
    EC と後続を **同一担当者で** 交互に詰ゝると EC はブロックされるため、従来はフェーズ1を先に詰ゝた。
    として後続は候補化した時点で **検査とともに実機械**のフェーズ1や **同一依頼の EC** は全日先に進むと」
    検査は `start_ge_end_initial`（設備空しは終業より後）で全日失敗する。§B-2/§B-3 後続はあるとしは
    「同一依頼EC・検査機と機械共有れるフェーズ1・後続」を **配台試行順**でマージし、
    坌順では **後続を EC より先に**」**しの他のフェーズ1** とあゝせで **配台試行順**で整列し
    **最大1ロールうつ**の値周回れる（マージ・rest とも一括ドレインしない。検査OPは他工程に
    同日坖り切られ start_ge_end_initial になるのを防し）。
    リワインド坴の後続行は坄ロールについで `_roll_pipeline_inspection_assign_room` よよよ
    `_roll_pipeline_b2_inspection_ec_completion_floor_dt`（EC ロール終了時刻下限）で整合する。
    試行順最尝の行の値は当日入らない場合でも」**坌もフェーズ内で次の試行順へ進み**他設備を埋ゝる。
    機械・人の空しはロールごとに更新れる（⑦⑧）。
    """
    _t_trial_pass0 = time_module.perf_counter()
    _mc_w0 = datetime.combine(current_date, DEFAULT_START_TIME)
    _t_mh_init0 = time_module.perf_counter()
    _mh_init = _machine_handoff_state_from_timeline_cached(timeline_events, current_date)
    _dispatch_loop_profile_add(
        "trial_pass_mh_init",
        time_module.perf_counter() - _t_mh_init0,
    )
    _gpo = global_priority_override or {}
    _assign_probe_ctx: dict | None = None
    if STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT:
        _assign_probe_ctx = {
            "avail_dt": avail_dt,
            "machine_avail_dt": machine_avail_dt,
            "task_queue": task_queue,
            "skills_dict": skills_dict,
            "members": members,
            "req_map": req_map,
            "need_rules": need_rules,
            "surplus_map": surplus_map,
            "global_priority_override": global_priority_override,
            "macro_run_date": macro_run_date,
            "macro_now_dt": macro_now_dt,
            "machine_handoff": _mh_init,
            "team_combo_presets": team_combo_presets,
            "dispatch_interval_mirror": dispatch_interval_mirror,
            "limited_equipment_mirror": limited_equipment_mirror,
        }
    _min_dispatch_eff: int | None = None
    _pool_min: list = []
    _t_pool0 = time_module.perf_counter()
    if STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT and _assign_probe_ctx:
        _pool_min = _tasks_in_min_pending_dispatch_pool(
            task_queue,
            current_date,
            daily_status=daily_status,
            members=members,
            machine_avail_dt=machine_avail_dt,
            machine_day_start=_mc_w0,
            machine_handoff=_mh_init,
            skills_dict=skills_dict,
            abolish_all_scheduling_limits=bool(
                _gpo.get("abolish_all_scheduling_limits")
            ),
            dispatch_interval_mirror=dispatch_interval_mirror,
            interactive_trial_pair_dates=interactive_trial_pair_dates,
        )
        _min_dispatch_eff = _effective_min_dispatch_trial_order_from_pool(
            _pool_min, current_date, daily_status, _assign_probe_ctx
        )
    _dispatch_loop_profile_add(
        "trial_pass_pool_min", time_module.perf_counter() - _t_pool0
    )
    _t_elig_pre0 = time_module.perf_counter()
    _wip_snap_pass = _stage2_eligible_wip_snapshot(task_queue)
    _pbo_index = _stage2_pending_by_machine_occ_index(task_queue, current_date)
    _rows_by_tid = _stage2_rows_by_task_id(task_queue)
    _rows_by_rule_tid = _stage2_rows_by_rule_task_id(task_queue)
    _window_left_cache: dict = {}
    _pipeline_room_cache: dict[str, float] = {}
    for _pt in tasks_today:
        _ptid = str(_pt.get("task_id") or "").strip()
        if _ptid and _ptid not in _pipeline_room_cache:
            _pipeline_room_cache[_ptid] = _roll_pipeline_inspection_assign_room(
                task_queue, _ptid
            )
    _dispatch_loop_profile_add(
        "trial_pass_eligible_precompute",
        time_module.perf_counter() - _t_elig_pre0,
    )
    _t_elig0 = time_module.perf_counter()
    eligible = _trial_order_flow_eligible_tasks(
        tasks_today,
        task_queue,
        current_date,
        daily_status=daily_status,
        members=members,
        machine_avail_dt=machine_avail_dt,
        machine_day_start=_mc_w0,
        machine_handoff=_mh_init,
        skills_dict=skills_dict,
        abolish_all_scheduling_limits=bool(_gpo.get("abolish_all_scheduling_limits")),
        dispatch_interval_mirror=dispatch_interval_mirror,
        min_dispatch_effective=_min_dispatch_eff,
        assign_probe_ctx=_assign_probe_ctx,
        interactive_trial_pair_dates=interactive_trial_pair_dates,
        wip_snapshot=_wip_snap_pass,
        pending_by_occ=_pbo_index,
        rows_by_tid=_rows_by_tid,
        rows_by_rule_tid=_rows_by_rule_tid,
        window_left_cache=_window_left_cache,
        pipeline_room_cache=_pipeline_room_cache,
    )
    if not eligible:
        if (
            cap_drain_only
            and _interactive_dispatch_cap_enforced_in_schedule_loop()
            and interactive_dispatch_targets
        ):
            eligible = [
                t
                for t in tasks_today
                if float(t.get("remaining_units") or 0) > 1e-12
            ]
        if not eligible:
            _dispatch_loop_profile_add(
                "trial_pass_total", time_module.perf_counter() - _t_trial_pass0
            )
            return False
    _dispatch_loop_profile_add(
        "trial_pass_eligible", time_module.perf_counter() - _t_elig0
    )
    eligible_sorted = sorted(
        eligible,
        key=lambda t: (
            dispatch_trial_order_key_from_task(t),
            # 特別ルール（特別ルール列挙.md）:
            # SEC×SEC機 湖南で「PN」を含む依頼NOは、同日に配台する場合は「JR」等より優先（同一試行順帯の先頭側へ寄せる）
            0
            if (
                _normalize_process_name_for_rule_match(t.get("machine"))
                == _normalize_process_name_for_rule_match("SEC")
                and _normalize_equipment_match_key(t.get("machine_name"))
                == _normalize_equipment_match_key("SEC機　湖南")
                and "PN" in unicodedata.normalize("NFKC", str(t.get("task_id") or ""))
            )
            else 1,
        ),
    )
    _mc_plan_end = _machine_calendar_planning_window_end_dt(
        current_date, daily_status, members
    )
    machine_handoff = {
        "last_tid": dict(_mh_init["last_tid"]),
        "last_eq": dict(_mh_init["last_eq"]),
        "last_process": dict(_mh_init.get("last_process") or {}),
        "started_today": set(_mh_init["started_today"]),
        "machining_today_occ": set(_mh_init.get("machining_today_occ") or set()),
        "last_machining_dt": dict(_mh_init.get("last_machining_dt") or {}),
        "last_machining_date": dict(_mh_init.get("last_machining_date") or {}),
        "last_lead_op": dict(_mh_init.get("last_lead_op") or {}),
        "last_machining_sub": dict(_mh_init.get("last_machining_sub") or {}),
        # 特別ルール（L9 等）向け: 直前ロールの試行順・製品厚み
        "last_dispatch_trial_order": dict(_mh_init.get("last_dispatch_trial_order") or {}),
        "last_product_thickness": dict(_mh_init.get("last_product_thickness") or {}),
    }

    def _l9_slice_continuity_key(t: dict) -> tuple:
        """
        L9: スライス×スライス機1 湖南で、直前ロールの試行順±10 かつ同厚みなら優先する。
        ただし原反投入日（start_date_req）・納期基準（due_basis_date）を満たさない場合は優先しない。
        """
        from planning_core.core.plan_input import dispatch_trial_order_key_from_task

        dto = dispatch_trial_order_key_from_task(t)
        proc = _normalize_process_name_for_rule_match(t.get("machine"))
        mach = _normalize_equipment_match_key(t.get("machine_name"))
        if not (
            proc == _normalize_process_name_for_rule_match("スライス")
            and mach == _normalize_equipment_match_key("スライス機1　湖南")
        ):
            return (
                dto,
                1,
                str(t.get("task_id") or ""),
                int(t.get("same_request_line_seq") or 0),
            )
        # 除外条件: 原反投入日（開始日下限）・納期基準を満たさない場合は優先しない
        sreq = t.get("start_date_req")
        if isinstance(sreq, date) and current_date < sreq:
            return (
                dto,
                1,
                str(t.get("task_id") or ""),
                int(t.get("same_request_line_seq") or 0),
            )
        due = t.get("due_basis_date")
        if isinstance(due, date) and current_date > due:
            return (
                dto,
                1,
                str(t.get("task_id") or ""),
                int(t.get("same_request_line_seq") or 0),
            )
        eqt = str(t.get("equipment_line_key") or t.get("machine") or "").strip() or (
            t.get("machine") or ""
        )
        occ = (_machine_occupancy_key_resolve(t, eqt) or "").strip()
        if not occ:
            return (
                dto,
                1,
                str(t.get("task_id") or ""),
                int(t.get("same_request_line_seq") or 0),
            )
        prev_dto = (machine_handoff.get("last_dispatch_trial_order") or {}).get(occ)
        prev_th = (machine_handoff.get("last_product_thickness") or {}).get(occ)
        th = t.get(PLAN_COL_PRODUCT_THICKNESS)
        prefer = False
        try:
            if (
                prev_dto is not None
                and abs(int(prev_dto) - int(dto)) <= 10
                and prev_th is not None
                and th is not None
            ):
                prefer = float(prev_th) == float(th)
        except Exception:
            prefer = False
        return (
            dto,
            0 if prefer else 1,
            str(t.get("task_id") or ""),
            int(t.get("same_request_line_seq") or 0),
        )

    from planning_core.core.process_machine_priority import (
        consecutive_prefer_sort_penalty,
        last_process_for_occupancy,
        load_priority_rules,
        priority_rank_for_task,
        should_defer_blocker_behind_consecutive_peers,
    )

    l14_rules = load_priority_rules()

    def _occ_of(t: dict) -> str:
        eqt = str(t.get("equipment_line_key") or t.get("machine") or "").strip()
        return (_machine_occupancy_key_resolve(t, eqt) or "").strip()

    def _phase1_sort_key(t: dict, peers: list | None = None) -> tuple:
        """段階3: 当該暦日キャップ未達 m が大きいタスクを先に（同一機械の依頼切替待ちを後ろ倒ししない）。"""
        dto_k, l9_pen, tid, seq = _l9_slice_continuity_key(t)
        occ = _occ_of(t)
        last_proc = last_process_for_occupancy(machine_handoff, occ)
        cons = consecutive_prefer_sort_penalty(
            t, machine_handoff, occupancy=occ, rules=l14_rules
        )
        rank = priority_rank_for_task(t, l14_rules)
        defer = 0
        if peers:
            if should_defer_blocker_behind_consecutive_peers(
                t,
                peers,
                occupancy=occ,
                last_process=last_proc,
                occupancy_of=_occ_of,
                rules=l14_rules,
            ):
                defer = 1
        rest = (defer, dto_k, l9_pen, cons, rank, tid, seq)
        unmet = _interactive_stage3_unmet_cap_m_on_date(
            t,
            current_date,
            interactive_dispatch_targets=interactive_dispatch_targets,
            interactive_trial_meters_done=interactive_trial_meters_done,
        )
        if _interactive_stage2_parity_active():
            return rest
        if unmet <= 1e-9:
            return (1, 0.0) + rest
        return (0, -unmet) + rest

    def _drain_rolls_for_task(
        task: dict, *, max_rolls: int | None = None
    ) -> bool:
        preferred_team: tuple | None = None
        made_local = False
        rolls_done = 0
        while float(task.get("remaining_units") or 0) > 1e-12:
            if max_rolls is not None and rolls_done >= max_rolls:
                break
            if _stage2_aladdin_next_day_exclude_consumes_roll(task, current_date):
                continue
            stage35_overtime_only = False
            _iv_cap = (
                _interactive_dispatch_cap_enforced_in_schedule_loop()
                and interactive_dispatch_targets is not None
                and interactive_trial_meters_done is not None
            )
            _tid_iv = _interactive_norm_cell(str(task.get("task_id") or ""))
            _proc_iv = _interactive_dispatch_target_process_key(task.get("machine"))
            _mach_iv = _interactive_norm_cell(str(task.get("machine_name") or ""))
            if _interactive_cap_schedule_blocked_before_earliest_target(
                interactive_dispatch_targets,
                _tid_iv,
                _proc_iv,
                _mach_iv,
                current_date,
            ):
                break
            _raw_cap_key = (_tid_iv, _proc_iv, _mach_iv, current_date)
            _cap_key = _raw_cap_key
            if _iv_cap and _raw_cap_key not in interactive_dispatch_targets:
                _resolved_ck = _interactive_dispatch_resolve_cap_key(
                    interactive_dispatch_targets=interactive_dispatch_targets,
                    interactive_trial_meters_done=interactive_trial_meters_done,
                    tid=_tid_iv,
                    proc=_proc_iv,
                    mach=_mach_iv,
                    current_date=current_date,
                )
                if _resolved_ck is not None:
                    _cap_key = _resolved_ck
            if _iv_cap and _cap_key in interactive_dispatch_targets:
                try:
                    _cap_m = float(interactive_dispatch_targets[_cap_key])
                    _done_m = float(interactive_trial_meters_done.get(_cap_key, 0.0))
                except (TypeError, ValueError):
                    _cap_m = 0.0
                    _done_m = 0.0
                if _done_m >= _cap_m - 1e-5:
                    if _overtime_simulation_dispatch_trial_active():
                        stage35_overtime_only = True
                    else:
                        break
                if not stage35_overtime_only:
                    try:
                        _um_lim = float(task.get("unit_m") or 0)
                    except (TypeError, ValueError):
                        _um_lim = 0.0
                    if _um_lim > 1e-12:
                        _rem_key_m = max(0.0, _cap_m - _done_m)
                        _rem_m = _rem_key_m
                        _tot_task = float(parse_float_safe(task.get("total_qty_m"), 0.0))
                        if (
                            _tot_task > 1e-12
                            and interactive_dispatch_targets is not None
                            and interactive_trial_meters_done is not None
                        ):
                            _glob_done_m = 0.0
                            for _gk in interactive_dispatch_targets:
                                if (
                                    isinstance(_gk, tuple)
                                    and len(_gk) == 4
                                    and _gk[0] == _tid_iv
                                    and _gk[1] == _proc_iv
                                    and _gk[2] == _mach_iv
                                ):
                                    try:
                                        _glob_done_m += float(
                                            interactive_trial_meters_done.get(_gk, 0.0)
                                        )
                                    except (TypeError, ValueError):
                                        pass
                            _rem_task_m = max(0.0, _tot_task - _glob_done_m)
                            _rem_m = min(_rem_key_m, _rem_task_m)
                        if _rem_m + 1e-9 < _um_lim:
                            break
            _t_assign0 = time_module.perf_counter()
            res = _assign_one_roll_trial_order_flow(
                task,
                current_date,
                daily_status,
                avail_dt,
                machine_avail_dt,
                task_queue,
                skills_dict,
                members,
                req_map,
                need_rules,
                surplus_map,
                global_priority_override,
                macro_run_date,
                macro_now_dt,
                preferred_team,
                _need_headcount_logged_orders,
                team_combo_presets,
                dispatch_interval_mirror=dispatch_interval_mirror,
                limited_equipment_mirror=limited_equipment_mirror,
                machine_handoff=machine_handoff,
                timeline_events=timeline_events,
                stage35_overtime_only=stage35_overtime_only,
            )
            _dispatch_loop_profile_add(
                "assign_one_roll", time_module.perf_counter() - _t_assign0
            )
            if res is None:
                break
            done_units = 1
            if task.get("roll_pipeline_inspection") or task.get(
                "roll_pipeline_rewind"
            ):
                _rp_room = _roll_pipeline_inspection_assign_room(
                    task_queue, str(task.get("task_id", "") or "").strip()
                )
                if _rp_room <= 1e-12:
                    break
                done_units = min(
                    1, int(min(_rp_room, math.ceil(task["remaining_units"])))
                )
            if done_units <= 0:
                break
            best_team = tuple(res["team"])
            lead_op = res["lead_op"]
            sub_members = [m for m in best_team if m != lead_op]
            best_start = res["team_start"]
            best_end = res["actual_end_dt"]
            best_breaks = res["team_breaks"]
            best_eff = res["avg_eff"]
            rq_base = res["rq_base"]
            extra_max = res["extra_max"]
            eq_line = res["eq_line"]
            machine_occ_key = _machine_occupancy_key_resolve(task, eq_line)
            _te_disp = parse_float_safe(task.get("task_eff_factor"), 1.0)
            if _te_disp <= 0:
                _te_disp = 1.0

            total_u = (
                math.ceil(task["total_qty_m"] / task["unit_m"]) if task["unit_m"] else 0
            )
            rem_u_before = math.ceil(task["remaining_units"])
            already_done = total_u - rem_u_before
            try:
                tot_qty = parse_float_safe(task.get("total_qty_m"), 0.0)
                done_qty = parse_float_safe(task.get("done_qty_reported"), 0.0)
                if tot_qty > 0:
                    pct_macro = max(
                        0, min(100, int(round((done_qty / tot_qty) * 100)))
                    )
                else:
                    pct_macro = 0
            except Exception:
                pct_macro = 0

            _mach_sub_line = ", ".join(
                str(s).strip() for s in sub_members if s and str(s).strip()
            )
            _co_append = list(res.get("changeover_segments") or [])
            _t_timeline0 = time_module.perf_counter()
            _append_changeover_segments_to_timeline(
                timeline_events,
                dispatch_interval_mirror,
                avail_dt,
                daily_status,
                current_date=current_date,
                task_id=str(task.get("task_id") or ""),
                machine_occ_key=machine_occ_key,
                segments=_co_append,
                machining_lead_op=str(lead_op or "").strip() or None,
                machining_sub_str=_mach_sub_line or None,
                machine_handoff=machine_handoff,
                skill_role_priority=res.get("startup_skill_role_priority"),
                machine_name_for_startup=str(res.get("machine_name") or "").strip()
                or None,
            )
            _mach_timeline_ev = {
                "date": current_date,
                "task_id": task["task_id"],
                "machine": eq_line,
                "machine_occupancy_key": machine_occ_key,
                "op": lead_op,
                "sub": ", ".join(sub_members),
                "start_dt": best_start,
                "end_dt": best_end,
                "breaks": best_breaks,
                "units_done": done_units,
                "already_done_units": already_done,
                "total_units": total_u,
                "pct_macro": pct_macro,
                "eff_time_per_unit": task["base_time_per_unit"]
                / best_eff
                / _te_disp
                * _surplus_team_time_factor(
                    rq_base, len(best_team), extra_max
                ),
                "unit_m": task["unit_m"],
                "total_qty_m": float(parse_float_safe(task.get("total_qty_m"), 0.0)),
                "event_kind": TIMELINE_EVENT_MACHINING,
            }
            timeline_events.append(_mach_timeline_ev)
            _stage2_dispatch_track_timeline_event(
                _mach_timeline_ev, len(timeline_events)
            )
            _dispatch_loop_profile_add(
                "timeline_commit", time_module.perf_counter() - _t_timeline0
            )
            if dispatch_interval_mirror is not None:
                dispatch_interval_mirror.register_from_event(timeline_events[-1])
            if (
                _iv_cap
                and interactive_dispatch_targets is not None
                and interactive_trial_meters_done is not None
                and _cap_key in interactive_dispatch_targets
            ):
                try:
                    _um_iv_acc = float(task.get("unit_m") or 0)
                except (TypeError, ValueError):
                    _um_iv_acc = 0.0
                interactive_trial_meters_done[_cap_key] = float(
                    interactive_trial_meters_done.get(_cap_key, 0.0)
                ) + float(done_units) * _um_iv_acc
            task["remaining_units"] = max(
                0.0,
                float(task.get("remaining_units") or 0) - float(done_units),
            )
            op_main = (lead_op or "").strip()
            subs_part = ",".join(
                s.strip() for s in sub_members if s and str(s).strip()
            )
            team_s = f"{op_main}, {subs_part}" if subs_part else op_main
            req_num_run = int(res.get("req_num") or 0)
            extra_max_run = int(res.get("extra_max") or 0)
            need_surplus_assigned = (
                TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
                and extra_max_run > 0
                and len(best_team) > req_num_run
            )
            names_ordered: list[str] = []
            if op_main:
                names_ordered.append(op_main)
            for _m in sub_members:
                if _m and str(_m).strip():
                    names_ordered.append(str(_m).strip())
            surplus_member_names = (
                names_ordered[req_num_run:]
                if need_surplus_assigned
                and len(names_ordered) > req_num_run
                else []
            )
            task["assigned_history"].append(
                {
                    "date": current_date.strftime("%m/%d"),
                    "team": team_s,
                    "done_m": int(done_units * task["unit_m"]),
                    "start_dt": best_start,
                    "end_dt": best_end,
                    "need_surplus_assigned": need_surplus_assigned,
                    "combo_sheet_row_id": res.get("combo_sheet_row_id"),
                    "surplus_member_names": surplus_member_names,
                }
            )
            for m in best_team:
                avail_dt[m] = best_end
            if _machine_occupancy_tracking_required(_gpo, task_queue):
                machine_avail_dt[machine_occ_key] = best_end
                _bump_machine_avail_after_roll_for_calendar(
                    current_date,
                    machine_occ_key,
                    machine_avail_dt,
                    machine_calendar_plan_end=_mc_plan_end,
                    machine_day_floor=_mc_w0,
                )
            _register_limited_equipment_interval(
                limited_equipment_mirror,
                task,
                machine_occ_key,
                best_start,
                best_end,
            )
            machine_handoff["last_tid"][machine_occ_key] = str(
                task.get("task_id") or ""
            ).strip()
            machine_handoff["last_eq"][machine_occ_key] = eq_line
            from planning_core.core.process_machine_priority import remember_last_process

            remember_last_process(
                machine_handoff,
                machine_occ_key,
                process_name=str(task.get("machine") or ""),
                eq_line=eq_line,
            )
            machine_handoff["started_today"].add(machine_occ_key)
            machine_handoff["machining_today_occ"].add(machine_occ_key)
            machine_handoff["last_machining_dt"][machine_occ_key] = best_end
            machine_handoff["last_machining_date"][machine_occ_key] = current_date
            machine_handoff["last_lead_op"][machine_occ_key] = lead_op
            machine_handoff.setdefault("last_machining_sub", {})
            machine_handoff["last_machining_sub"][machine_occ_key] = _mach_sub_line
            machine_handoff.setdefault("last_dispatch_trial_order", {})
            machine_handoff.setdefault("last_product_thickness", {})
            from planning_core.core.plan_input import dispatch_trial_order_key_from_task

            machine_handoff["last_dispatch_trial_order"][machine_occ_key] = (
                dispatch_trial_order_key_from_task(task)
            )
            machine_handoff["last_product_thickness"][machine_occ_key] = task.get(
                PLAN_COL_PRODUCT_THICKNESS
            )
            if _trace_schedule_task_enabled(task.get("task_id")):
                _log_dispatch_trace_schedule(
                    task.get("task_id"),
                    "[配台トレース task=%s] ロール確定 メイン day=%s machine=%s machine_name=%s "
                    "start=%s end=%s 採用人数=%s req_num=%s メイン探索extra_max=%s "
                    "余剰人数適用(メイン)=%s team=%s",
                    task.get("task_id"),
                    current_date,
                    eq_line,
                    str(task.get("machine_name") or "").strip(),
                    best_start,
                    best_end,
                    len(best_team),
                    req_num_run,
                    extra_max_run,
                    need_surplus_assigned,
                    team_s,
                )
            preferred_team = best_team
            made_local = True
            rolls_done += 1
        return made_local

    def _is_b2_follower_phase2_row(t: dict) -> bool:
        """§B-2/§B-3 後続行。EC 行がキューから落ちていてもマージ・二相割付の対象に含める。"""
        return bool(
            t.get("roll_pipeline_inspection") or t.get("roll_pipeline_rewind")
        )

    phase1_tasks = [t for t in eligible_sorted if not _is_b2_follower_phase2_row(t)]
    phase2_tasks = [t for t in eligible_sorted if _is_b2_follower_phase2_row(t)]

    phase2_tids: set[str] = {
        str(t.get("task_id") or "").strip()
        for t in phase2_tasks
        if str(t.get("task_id") or "").strip()
    }
    phase2_mocc: set[str] = set()
    for t in phase2_tasks:
        _tm = t.get("machine")
        _eqt = str(t.get("equipment_line_key") or _tm or "").strip() or (_tm or "")
        _pk = (_machine_occupancy_key_resolve(t, _eqt) or "").strip()
        if _pk:
            phase2_mocc.add(_pk)

    phase1_interleave: list = []
    phase1_rest: list = []
    for t in phase1_tasks:
        _tid1 = str(t.get("task_id") or "").strip()
        _tm = t.get("machine")
        _eqt = str(t.get("equipment_line_key") or _tm or "").strip() or (_tm or "")
        _mk = (_machine_occupancy_key_resolve(t, _eqt) or "").strip()
        _same_tid_ec = bool(t.get("roll_pipeline_ec") and _tid1 and _tid1 in phase2_tids)
        _share_m = bool(_mk and _mk in phase2_mocc)
        if _same_tid_ec or _share_m:
            phase1_interleave.append(t)
        else:
            phase1_rest.append(t)

    def _b2_merged_sort_key(t: dict) -> tuple:
        # 坌も配台試行順では後続（検査・巻返し）を EC より先に回し、熱融着のタイムラインを
        # 同日早い段階で坖りに行し（§B-2 担当者分離で EC と検査は別メンバー想定）。
        _fol = bool(
            t.get("roll_pipeline_inspection") or t.get("roll_pipeline_rewind")
        )
        return (
            dispatch_trial_order_key_from_task(t),
            0 if _fol else 1,
            str(t.get("task_id") or ""),
            int(t.get("same_request_line_seq") or 0),
        )

    pass_made = False
    if cap_drain_only and (
        _interactive_dispatch_cap_enforced_in_schedule_loop()
        and interactive_dispatch_targets
    ):
        _cap_drain_source = (
            [
                t
                for t in tasks_today
                if float(t.get("remaining_units") or 0) > 1e-12
            ]
            if cap_drain_only
            else list(phase1_tasks)
        )
        _cap_drain_selected = [
            t
            for t in _cap_drain_source
            if _interactive_stage3_unmet_cap_m_on_date(
                t,
                current_date,
                interactive_dispatch_targets=interactive_dispatch_targets,
                interactive_trial_meters_done=interactive_trial_meters_done,
            )
            > 1e-9
        ]
        _cap_drain_tasks = sorted(
            _cap_drain_selected,
            key=lambda t: _phase1_sort_key(t, _cap_drain_selected),
        )
        _cap_drain_ids = {id(t) for t in _cap_drain_tasks}
        for _ct in _cap_drain_tasks:
            if _drain_rolls_for_task(_ct):
                pass_made = True
        if _cap_drain_ids:
            phase1_interleave = [
                t for t in phase1_interleave if id(t) not in _cap_drain_ids
            ]
            phase1_rest = [t for t in phase1_rest if id(t) not in _cap_drain_ids]
            phase1_tasks = [t for t in phase1_tasks if id(t) not in _cap_drain_ids]
        _dispatch_loop_profile_add(
            "trial_pass_total", time_module.perf_counter() - _t_trial_pass0
        )
        return pass_made
    if phase2_tasks:
        if (
            _interactive_dispatch_cap_enforced_in_schedule_loop()
            and interactive_dispatch_targets
            and phase1_rest
        ):
            # 段階3（非パリティ）: B2 ラウンドロビン（1ロールずつ）だと同一機械で未達大の依頼と交替し
            # 開始が遅れる。phase1_rest は暦日キャップ優先で一括ドレインする。
            for task in sorted(phase1_rest, key=lambda t: _phase1_sort_key(t, phase1_rest)):
                if _drain_rolls_for_task(task):
                    pass_made = True
            phase1_rest = []
        elif phase1_rest or phase1_interleave:
            # 段階2: §B-2 後続が載る日は、同一依頼の先行（例: スリット残）を
            # 1ロール交互より先にドレインする。interleave（機械共有の EC 等）も対象。
            for task in sorted(
                phase1_rest + phase1_interleave,
                key=lambda t: _phase1_sort_key(t, phase1_rest + phase1_interleave),
            ):
                if _drain_rolls_for_task(task):
                    pass_made = True
            phase1_rest = [
                t
                for t in phase1_rest
                if float(t.get("remaining_units") or 0) > 1e-12
            ]
            phase1_interleave = [
                t
                for t in phase1_interleave
                if float(t.get("remaining_units") or 0) > 1e-12
            ]
        merged_b2 = sorted(
            phase1_interleave + phase2_tasks,
            key=lambda t: _phase1_sort_key(t, phase1_interleave + phase2_tasks),
        )
        _merged_row_ids = {id(x) for x in merged_b2}

        def _b2_rr_key(t: dict) -> tuple:
            return _phase1_sort_key(t, merged_b2 + phase1_rest)

        while True:
            round_made = False
            all_rr = sorted(merged_b2 + phase1_rest, key=_b2_rr_key)
            for task in all_rr:
                if float(task.get("remaining_units") or 0) <= 1e-12:
                    continue
                if _drain_rolls_for_task(task, max_rolls=1):
                    round_made = True
            if not round_made:
                break
            pass_made = True
    else:
        for task in sorted(phase1_tasks, key=lambda t: _phase1_sort_key(t, phase1_tasks)):
            if _drain_rolls_for_task(task):
                pass_made = True
    _dispatch_loop_profile_add(
        "trial_pass_total", time_module.perf_counter() - _t_trial_pass0
    )
    return pass_made
def _run_b2_inspection_rewind_pass(
    sorted_dates: list,
    attendance_data: dict,
    task_queue: list,
    timeline_events: list,
    skills_dict: dict,
    members: list,
    equipment_list: list,
    req_map: dict,
    need_rules: list,
    surplus_map: dict,
    global_priority_override: dict,
    macro_run_date: date,
    macro_now_dt: datetime,
    _need_headcount_logged_orders: set,
    team_combo_presets: dict | None = None,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
    interactive_dispatch_targets: dict | None = None,
    interactive_trial_pair_dates: dict | None = None,
    interactive_trial_meters_done: dict | None = None,
) -> bool:
    """
    §B-2 / §B-3: EC 坴を先に全日で進ゝた後」検査＝巻返し坴のみを日付先頭から再走査して配台れる。
    timeline_events を人・設備のブロックテーブルとして使い」日跨ねの占有を保挝れる。
    """
    target_tids: set[str] = set()
    for t in task_queue:
        if float(t.get("remaining_units") or 0) <= 1e-12:
            continue
        if not (t.get("roll_pipeline_inspection") or t.get("roll_pipeline_rewind")):
            continue
        tid = str(t.get("task_id", "") or "").strip()
        if not tid:
            continue
        # EC 行がキューに無い＝完走後に欠落した依頼は、EC 済みとみなして後続専用パスへ含める
        # （L10 B-4.1 のスリット行欠落と同趣旨。assign_room は既に UNCAPPED）。
        if not _task_queue_has_roll_pipeline_ec_for_tid(task_queue, tid):
            target_tids.add(tid)
            continue
        if not _pipeline_ec_fully_done_for_tid(task_queue, tid):
            continue
        target_tids.add(tid)
    if not target_tids:
        return False

    _gpo = global_priority_override or {}
    _any_progress = False
    _machine_day_start_cache: dict[date, datetime] = {}
    for current_date in sorted_dates:
        daily_status = attendance_data.get(current_date)
        if not daily_status:
            continue
        machine_avail_dt: dict = {}
        avail_dt: dict = {}
        for m in members:
            if m not in daily_status:
                continue
            st = daily_status[m]
            if st.get("eligible_for_assignment", st.get("is_working", False)):
                avail_dt[m] = st["start_dt"]
        if not avail_dt:
            continue

        _machine_day_start = _machine_day_start_cache.get(current_date)
        if _machine_day_start is None:
            _machine_day_start = datetime.combine(current_date, DEFAULT_START_TIME)
            _machine_day_start_cache[current_date] = _machine_day_start
        _seed_avail_from_timeline_for_date(
            timeline_events,
            current_date,
            machine_avail_dt,
            avail_dt,
            _machine_day_start,
        )
        if not _machine_occupancy_tracking_required(_gpo, task_queue):
            machine_avail_dt.clear()
        _mc_plan_end_b2 = _machine_calendar_planning_window_end_dt(
            current_date, daily_status, members
        )
        _apply_machine_calendar_floor_for_date(
            current_date,
            machine_avail_dt,
            equipment_list,
            _machine_day_start,
            machine_calendar_plan_end=_mc_plan_end_b2,
        )

        tasks_today = [
            t
            for t in task_queue
            if float(t.get("remaining_units") or 0) > 1e-12
            and (
                t.get("roll_pipeline_inspection")
                or t.get("roll_pipeline_rewind")
            )
            and str(t.get("task_id", "") or "").strip() in target_tids
            and t.get("start_date_req") <= current_date
        ]
        if not tasks_today:
            continue

        _sched_max_passes = max(96, max(1, len(tasks_today)) * 15)
        _sched_pi = 0
        while _sched_pi < _sched_max_passes:
            _sched_pi += 1
            _made = _trial_order_first_schedule_pass(
                current_date,
                tasks_today,
                task_queue,
                daily_status,
                machine_avail_dt,
                avail_dt,
                timeline_events,
                skills_dict,
                members,
                req_map,
                need_rules,
                surplus_map,
                global_priority_override,
                macro_run_date,
                macro_now_dt,
                _need_headcount_logged_orders,
                team_combo_presets,
                dispatch_interval_mirror=dispatch_interval_mirror,
                interactive_dispatch_targets=interactive_dispatch_targets,
                interactive_trial_pair_dates=interactive_trial_pair_dates,
                interactive_trial_meters_done=interactive_trial_meters_done,
            )
            if not _made:
                break
            _any_progress = True
    return _any_progress
def _timeline_event_team_names_set(ev: dict) -> set:
    names: set = set()
    op = str(ev.get("op") or "").strip()
    if op:
        names.add(op)
    sub = str(ev.get("sub") or "").strip()
    if sub:
        for s in sub.split(","):
            t = s.strip()
            if t:
                names.add(t)
    return names
def _timeline_event_fallback_task_from_want(
    ev: dict, task_queue: list, want: set | None, tid_n: str
) -> dict | None:
    """
    task_queue が依頼NO 1 行に潰れているとき、want の (工程, 機械) とイベントの設備キーで行を復元する。
    """
    if not want:
        return None
    ev_eq_n = _normalize_equipment_match_key(str(ev.get("machine") or ""))
    ev_oc_n = _normalize_equipment_match_key(str(ev.get("machine_occupancy_key") or ""))
    best_k = None
    best_sc = -1
    for k in want:
        if not isinstance(k, tuple) or len(k) != 4:
            continue
        if _interactive_norm_cell(str(k[0])) != tid_n:
            continue
        km = _interactive_norm_cell(str(k[2]))
        kmnk = _normalize_equipment_match_key(km)
        sc = 0
        if kmnk and (
            kmnk == ev_eq_n
            or kmnk == ev_oc_n
            or (ev_eq_n and (kmnk in ev_eq_n or ev_eq_n in kmnk))
            or (ev_oc_n and (kmnk in ev_oc_n or ev_oc_n in kmnk))
        ):
            sc += 80
        if sc > best_sc:
            best_sc = sc
            best_k = k
    if best_k is None or best_sc < 1:
        return None
    proc_t = str(best_k[1])
    for t in task_queue:
        if _interactive_norm_cell(str(t.get("task_id") or "")) != tid_n:
            continue
        if _interactive_dispatch_target_process_key(t.get("machine")) == proc_t:
            return t
    for t in task_queue:
        if _interactive_norm_cell(str(t.get("task_id") or "")) != tid_n:
            continue
        out = dict(t)
        out["machine"] = proc_t
        out["machine_name"] = str(best_k[2])
        return out
    return None
def _timeline_event_task_match_score(
    ev: dict, t: dict, want: set | None, tid_n: str
) -> int:
    """タイムライン加工イベント ev とタスク行 t の適合度（高いほど同一行とみなしやすい）。"""
    eq = str(ev.get("machine") or "").strip()
    t_eq = str(t.get("equipment_line_key") or t.get("machine") or "").strip()
    sc = 0
    if t_eq and t_eq == eq:
        sc += 100
    ne, nt = _normalize_equipment_match_key(eq), _normalize_equipment_match_key(t_eq)
    if ne and nt and ne == nt:
        sc += 50
    ev_occ = _normalize_equipment_match_key(str(ev.get("machine_occupancy_key") or ""))
    if ev_occ:
        try:
            t_occ = str(_machine_occupancy_key_resolve(t, t_eq) or "").strip()
        except Exception:
            t_occ = ""
        if _normalize_equipment_match_key(t_occ) == ev_occ:
            sc += 80
    mn = str(t.get("machine_name") or "").strip()
    nm = _normalize_equipment_match_key(mn)
    if ne and nm:
        if ne == nm:
            sc += 40
        elif ne in nm or nm in ne:
            sc += 18
    if want:
        tp = _interactive_dispatch_target_process_key(t.get("machine"))
        for k in want:
            if not isinstance(k, tuple) or len(k) != 4:
                continue
            if _interactive_norm_cell(str(k[0])) != tid_n:
                continue
            if str(k[1]) != tp:
                continue
            km = _interactive_norm_cell(str(k[2]))
            kmnk = _normalize_equipment_match_key(km)
            if kmnk and (
                kmnk == ne
                or kmnk == ev_occ
                or (ne and (kmnk in ne or ne in kmnk))
                or (ev_occ and (kmnk in ev_occ or ev_occ in kmnk))
            ):
                sc += 75
                break
    return sc
def _task_dict_for_timeline_event(
    ev: dict, task_queue: list, want: set | None = None
) -> dict | None:
    tid = str(ev.get("task_id") or "").strip()
    if not tid:
        return None
    tid_n = _interactive_norm_cell(tid)
    eq = str(ev.get("machine") or "").strip()
    ev_occ = str(ev.get("machine_occupancy_key") or "").strip()
    norm_ev_occ = _normalize_equipment_match_key(ev_occ) if ev_occ else ""
    norm_eq_ev = _normalize_equipment_match_key(eq)

    candidates: list[dict] = []
    for t in task_queue:
        if _interactive_norm_cell(str(t.get("task_id") or "")) != tid_n:
            continue
        candidates.append(t)
    if not candidates:
        return None

    for t in candidates:
        t_eq = str(t.get("equipment_line_key") or t.get("machine") or "").strip()
        if t_eq == eq:
            return t

    for t in candidates:
        t_eq = str(t.get("equipment_line_key") or t.get("machine") or "").strip()
        if norm_eq_ev and _normalize_equipment_match_key(t_eq) == norm_eq_ev:
            return t

    if norm_ev_occ:
        for t in candidates:
            t_eq = str(t.get("equipment_line_key") or t.get("machine") or "").strip()
            try:
                t_occ = str(_machine_occupancy_key_resolve(t, t_eq) or "").strip()
            except Exception:
                t_occ = ""
            if t_occ and _normalize_equipment_match_key(t_occ) == norm_ev_occ:
                return t

    multi_proc = False
    if want:
        pst = set()
        for k in want:
            if isinstance(k, tuple) and len(k) == 4:
                if _interactive_norm_cell(str(k[0])) == tid_n:
                    pst.add(str(k[1]))
        multi_proc = len(pst) > 1
    else:
        neqs = {
            _normalize_equipment_match_key(
                str(c.get("equipment_line_key") or c.get("machine") or "")
            )
            for c in candidates
        }
        neqs.discard("")
        multi_proc = len(neqs) > 1

    # 同一依頼NOに複数工程の target があるとき、占有キー一致などで SEC が高得点になり
    # スリット設備のイベントでも SEC 行が選ばれ meters_done が SEC に積まれるのを防ぐ。
    # イベント設備キーと want の (工程, 機械) の突合せをスコアリングより先に採用する。
    if multi_proc and want:
        fb_first = _timeline_event_fallback_task_from_want(ev, task_queue, want, tid_n)
        if fb_first is not None:
            return fb_first

    scored = [(_timeline_event_task_match_score(ev, t, want, tid_n), t) for t in candidates]
    scored.sort(key=lambda x: -x[0])
    best_sc, best_t = scored[0]

    if best_sc >= 50:
        return best_t
    if len(candidates) == 1 and not multi_proc:
        return candidates[0]
    fb = _timeline_event_fallback_task_from_want(ev, task_queue, want, tid_n)
    if fb is not None:
        return fb
    if best_sc > 0 and not multi_proc:
        return best_t
    return None
def _interactive_trial_recompute_meters_done_from_timeline(
    timeline_events: list,
    task_queue: list,
    targets: dict | None,
) -> dict[tuple[str, str, str, date], float]:
    """
    インタラクティブ試行: timeline の加工イベントから (依頼NO, 工程名, 機械名, 配台日) キー別の換算mを再集計する。

    イベント暦日と結果_配台表 JSON の配台日が一致しない場合は、
    `_interactive_dispatch_resolve_cap_key` とフォールバックでキーを解決する。
    """
    acc: dict[tuple[str, str, str, date], float] = {}
    if not targets:
        return acc
    want = set(targets.keys())
    _mach_evs = [e for e in (timeline_events or []) if _is_machining_timeline_event(e)]

    def _ev_sort_key(ev: dict):
        st = ev.get("start_dt")
        if isinstance(st, datetime):
            return st
        return datetime.max.replace(tzinfo=None)

    _mach_evs.sort(key=_ev_sort_key)
    for ev in _mach_evs:
        tid = _interactive_norm_cell(ev.get("task_id"))
        tsk = _task_dict_for_timeline_event(ev, task_queue, want)
        if tsk is None:
            continue
        proc_n = _interactive_dispatch_target_process_key(tsk.get("machine"))
        mach_n = _interactive_norm_cell(tsk.get("machine_name"))
        d = _interactive_timeline_event_calendar_date(ev)
        if d is None:
            continue
        try:
            ud = float(ev.get("units_done") or 0)
            um = float(tsk.get("unit_m") or 0)
        except (TypeError, ValueError):
            continue
        add_m = ud * um
        if add_m <= 1e-18:
            continue
        k_direct = (tid, proc_n, mach_n, d)
        if k_direct in want:
            kk = k_direct
        else:
            kk = _interactive_dispatch_resolve_cap_key(
                interactive_dispatch_targets=targets,
                interactive_trial_meters_done=acc,
                tid=tid,
                proc=proc_n,
                mach=mach_n,
                current_date=d,
            )
            if kk is None:
                kk = _interactive_fallback_meter_target_key_for_recompute(
                    tid, proc_n, mach_n, d, want
                )
        if kk is None:
            continue
        acc[kk] = acc.get(kk, 0.0) + add_m
    return acc
def _timeline_break_intervals_for_occ(
    timeline_events: list, day_d: date, occ: str
) -> list[tuple[datetime, datetime]]:
    occ_n = str(occ or "").strip()
    breaks_acc: list = []
    for e in timeline_events:
        if e.get("date") != day_d:
            continue
        if str(e.get("machine_occupancy_key") or "").strip() != occ_n:
            continue
        for item in e.get("breaks") or []:
            if not isinstance(item, (list, tuple)) or len(item) < 2:
                continue
            bs, be = item[0], item[1]
            if isinstance(bs, datetime) and isinstance(be, datetime):
                breaks_acc.append((bs, be))
    return merge_time_intervals(breaks_acc)
def _daily_startup_snap_end_before_first_machining(
    timeline_events: list,
    day_d: date,
    occ: str,
    startup_end: datetime,
    first_machining_start: datetime,
    align_tol: timedelta,
) -> datetime | None:
    """
    日次始業帯を寄せるときの終了時刻（半開区間 [start, end) の end）。

    - 通常: 先頭加工開始
    - 日次始業終了と先頭加工の間に暦休憩がある、または先頭加工が休憩再開準備直後:
      その間で先頭加工に最も近い休憩の開始（休憩直前）
    """
    m_st = first_machining_start
    if not isinstance(m_st, datetime) or not isinstance(startup_end, datetime):
        return None
    best_bs: datetime | None = None
    best_be: datetime | None = None
    for bs, be in _timeline_break_intervals_for_occ(timeline_events, day_d, occ):
        if be <= startup_end + align_tol:
            continue
        if bs >= m_st - align_tol:
            continue
        if be > startup_end and bs < m_st:
            if best_be is None or be > best_be:
                best_bs = bs
                best_be = be
    if best_bs is not None:
        return best_bs
    return m_st
def _repair_timeline_daily_startup_snapped_to_first_machining(
    timeline_events: list,
    task_queue: list,
    attendance_data: dict,
    skills_dict: dict,
    global_priority_override: dict | None,
) -> int:
    """
    当日・同一占有キーで、日次始業準備の終了が先頭の加工開始より前に空きがあるとき、
    壁時計の長さを保って寄せる（表示・ミラー整合用の後処理）。

    寄せ先の終了は通常は先頭加工開始。日次始業と先頭加工の間に暦休憩があるときは
    直前休憩の開始（休憩直前）。休憩再開準備は先頭加工直前のまま維持する。
    """
    gpo = global_priority_override or {}
    if not timeline_events:
        return 0
    keys: set[tuple[date, str]] = set()
    for e in timeline_events:
        if _timeline_event_kind(e) != TIMELINE_EVENT_MACHINE_DAILY_STARTUP:
            continue
        d = e.get("date")
        occ = str(e.get("machine_occupancy_key") or "").strip()
        if isinstance(d, date) and occ:
            keys.add((d, occ))
    n_adj = 0
    align_tol = timedelta(seconds=90)
    for day_d, occ in keys:
        daily_status = attendance_data.get(day_d)
        if not daily_status:
            continue
        su_candidates = [
            e
            for e in timeline_events
            if e.get("date") == day_d
            and str(e.get("machine_occupancy_key") or "").strip() == occ
            and _timeline_event_kind(e) == TIMELINE_EVENT_MACHINE_DAILY_STARTUP
        ]
        mach_sorted = sorted(
            (
                e
                for e in timeline_events
                if e.get("date") == day_d
                and str(e.get("machine_occupancy_key") or "").strip() == occ
                and _is_machining_timeline_event(e)
            ),
            key=lambda e: (e.get("start_dt") or datetime.max),
        )
        if not su_candidates or not mach_sorted:
            continue
        su_ev = min(
            su_candidates,
            key=lambda e: e.get("start_dt") or datetime.min,
        )
        first_m = mach_sorted[0]
        st0 = su_ev.get("start_dt")
        ed0 = su_ev.get("end_dt")
        m_st = first_m.get("start_dt")
        if (
            not isinstance(st0, datetime)
            or not isinstance(ed0, datetime)
            or not isinstance(m_st, datetime)
        ):
            continue
        if ed0 <= st0 or st0 >= m_st:
            continue
        snap_end = _daily_startup_snap_end_before_first_machining(
            timeline_events, day_d, occ, ed0, m_st, align_tol
        )
        if snap_end is None:
            continue
        if ed0 >= snap_end - align_tol:
            continue
        dur = ed0 - st0
        if dur <= timedelta(0):
            continue
        new_ed = snap_end
        new_st = new_ed - dur
        if new_st >= m_st:
            continue
        blocked = False
        for e in timeline_events:
            if e is su_ev:
                continue
            if e.get("date") != day_d:
                continue
            if str(e.get("machine_occupancy_key") or "").strip() != occ:
                continue
            o_st = e.get("start_dt")
            o_ed = e.get("end_dt")
            if not isinstance(o_st, datetime) or not isinstance(o_ed, datetime):
                continue
            if new_st < o_ed and o_st < new_ed:
                blocked = True
                break
        if blocked:
            continue
        su_ev["start_dt"] = new_st
        su_ev["end_dt"] = new_ed
        task = _task_dict_for_timeline_event(first_m, task_queue)
        machine_name = str((task or {}).get("machine_name") or "").strip()
        if not machine_name:
            machine_name = str(su_ev.get("machine") or "").strip()
        if task and machine_name:
            skill_meta_cache: dict = {}

            def skill_role_priority(mem: str):
                if gpo.get("ignore_skill_requirements"):
                    return ("OP", 100)
                mm = str(mem or "").strip()
                if not mm:
                    return ("", 9999)
                if mm not in skill_meta_cache:
                    srow = skills_dict.get(mm, {})
                    mp = str(task.get("machine") or "").strip()
                    mn = str(task.get("machine_name") or "").strip()
                    if mp and mn:
                        v = srow.get(f"{mp}+{mn}", "")
                    elif mn:
                        v = srow.get(mn, "")
                    elif mp:
                        v = srow.get(mp, "")
                    else:
                        v = ""
                    skill_meta_cache[mm] = parse_op_as_skill_cell(v)
                return skill_meta_cache[mm]

            su_ev["op"] = ""
            su_ev["sub"] = ""
            _daily_startup_fill_segment_staff(
                su_ev,
                machine_name=machine_name,
                lead_op=str(first_m.get("op") or "").strip(),
                sub_csv=str(first_m.get("sub") or "").strip(),
                skill_role_priority=skill_role_priority,
                daily_status=daily_status,
                avail_dt={},
                dispatch_interval_mirror=None,
            )
        else:
            su_ev["op"] = ""
            su_ev["sub"] = ""
        br_acc: list = []
        op_s = str(su_ev.get("op") or "").strip()
        sub_s = str(su_ev.get("sub") or "").strip()
        for nm in (op_s, *[x.strip() for x in sub_s.split(",") if x.strip()]):
            if nm and nm in daily_status:
                br_acc.extend(daily_status[nm].get("breaks_dt") or [])
        su_ev["breaks"] = merge_time_intervals(br_acc)
        n_adj += 1
    return n_adj
def _member_overlaps_busy(
    busy_map: dict, member: str, st: datetime, ed: datetime
) -> bool:
    for bs, be in busy_map.get(member, ()):
        if st < be and bs < ed:
            return True
    return False
def append_surplus_staff_after_main_dispatch(
    timeline_events: list,
    attendance_data: dict,
    skills_dict: dict,
    members: list,
    task_queue: list,
    req_map: dict,
    need_rules: list,
    surplus_map: dict,
    global_priority_override: dict | None,
) -> int:
    """
    need「配台時追加人数＝余力時追加人数」行の上限まで」メイン割付で採用ししれなかった枠を追記れる。
    坄タイムラインブロックについで」しの時間帯に他ブロックへ未坂加（区間重なりなし）で
    eligible かつ OP/AS スキルの者をサブに追加れる。
    日次始業（event_kind は加工以外）は本処理の対象外（余剰サブは加工にのみ追記）。
    """
    gpo = global_priority_override or {}
    if not surplus_map or TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW:
        return 0

    busy: dict[str, list[tuple[datetime, datetime]]] = defaultdict(list)
    for e in timeline_events:
        st = e.get("start_dt")
        ed = e.get("end_dt")
        if not isinstance(st, datetime) or not isinstance(ed, datetime):
            continue
        for name in _timeline_event_team_names_set(e):
            busy[name].append((st, ed))

    appended_total = 0
    sorted_evs = sorted(
        (
            e
            for e in timeline_events
            if isinstance(e.get("start_dt"), datetime)
            and isinstance(e.get("end_dt"), datetime)
        ),
        key=lambda x: (x.get("date"), x.get("start_dt") or datetime.min),
    )

    for ev in sorted_evs:
        d = ev.get("date")
        if d is None or d not in attendance_data:
            continue
        if not _is_machining_timeline_event(ev):
            continue
        daily_status = attendance_data[d]
        task = _task_dict_for_timeline_event(ev, task_queue)
        if task is None:
            continue
        machine = task.get("machine")
        machine_name = str(task.get("machine_name") or "").strip()
        # need/surplus 条件は元依頼NO（rule_task_id）で照合する（枝番タスクは親で判定）。
        tid = _rule_task_id(task)

        if TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY:
            req_num = resolve_need_required_op(
                str(machine or "").strip(),
                machine_name,
                tid,
                req_map,
                need_rules,
            )
        else:
            ro = task.get("required_op")
            if ro is not None:
                try:
                    riv = int(ro)
                    if riv >= 1:
                        req_num = riv
                    else:
                        req_num = resolve_need_required_op(
                            str(machine or "").strip(),
                            machine_name,
                            tid,
                            req_map,
                            need_rules,
                        )
                except (TypeError, ValueError):
                    req_num = resolve_need_required_op(
                        str(machine or "").strip(),
                        machine_name,
                        tid,
                        req_map,
                        need_rules,
                    )
            else:
                req_num = resolve_need_required_op(
                    str(machine or "").strip(),
                    machine_name,
                    tid,
                    req_map,
                    need_rules,
                )
        if gpo.get("ignore_need_minimum"):
            req_num = 1

        extra_max_sheet = resolve_need_surplus_extra_max(
            str(machine or "").strip(),
            machine_name,
            tid,
            surplus_map,
            need_rules,
        )
        if extra_max_sheet <= 0:
            continue

        names = _timeline_event_team_names_set(ev)
        team_size = len(names)
        cap_add = req_num + extra_max_sheet - team_size
        if cap_add <= 0:
            continue

        skill_meta_cache: dict = {}

        def skill_role_priority(mem):
            if gpo.get("ignore_skill_requirements"):
                return ("OP", 100)
            if mem not in skill_meta_cache:
                srow = skills_dict.get(mem, {})
                machine_proc = str(machine or "").strip()
                v = ""
                if machine_proc and machine_name:
                    v = srow.get(f"{machine_proc}+{machine_name}", "")
                elif machine_name:
                    v = srow.get(machine_name, "")
                elif machine_proc:
                    v = srow.get(machine_proc, "")
                skill_meta_cache[mem] = parse_op_as_skill_cell(v)
            return skill_meta_cache[mem]

        capable = []
        for mem in members:
            if mem not in daily_status:
                continue
            st_ent = daily_status[mem]
            if not st_ent.get(
                "eligible_for_assignment", st_ent.get("is_working", False)
            ):
                continue
            if skill_role_priority(mem)[0] not in ("OP", "AS"):
                continue
            capable.append(mem)
        capable.sort(key=lambda mm: (skill_role_priority(mm)[1], mm))

        st = ev["start_dt"]
        ed = ev["end_dt"]
        candidates = [
            m
            for m in capable
            if m not in names and not _member_overlaps_busy(busy, m, st, ed)
        ]
        candidates.sort(
            key=lambda mm: (
                0 if skill_role_priority(mm)[0] == "AS" else 1,
                skill_role_priority(mm)[1],
                mm,
            )
        )

        chosen = candidates[:cap_add]
        if not chosen:
            continue

        team_size_before = team_size
        final_team_size = team_size_before + len(chosen)
        highlight_surplus = final_team_size > req_num

        old_sub = str(ev.get("sub") or "").strip()
        parts = [s.strip() for s in old_sub.split(",") if s.strip()]
        parts.extend(chosen)
        ev["sub"] = ", ".join(parts)
        for m in chosen:
            busy[m].append((st, ed))
        appended_total += len(chosen)

        op_sync = str(ev.get("op") or "").strip()
        subs_sync = ",".join(
            s.strip()
            for s in str(ev.get("sub") or "").split(",")
            if s.strip()
        )
        team_sync = f"{op_sync}, {subs_sync}" if subs_sync else op_sync

        _hist = task.get("assigned_history")
        if _hist:
            for h in _hist:
                if (
                    h.get("start_dt") == st
                    and h.get("end_dt") == ed
                ):
                    if highlight_surplus:
                        h["need_surplus_assigned"] = True
                    h["team"] = team_sync
                    prev_pd = h.get("post_dispatch_surplus_names") or []
                    h["post_dispatch_surplus_names"] = prev_pd + [
                        str(x) for x in chosen
                    ]
                    break

        if _trace_schedule_task_enabled(tid):
            _log_dispatch_trace_schedule(
                tid,
                "[配台トレース task=%s] 余力追記(メイン完了後) day=%s machine=%s machine_name=%s "
                "start=%s end=%s 追記人数=%s 追記剝人数=%s 追記後人数=%s req_num=%s "
                "need追加枠(シート)=%s 履歴黄(余剰人数超靎)=%s 追記メンバー=%s",
                tid,
                d,
                str(machine or "").strip(),
                machine_name,
                st,
                ed,
                len(chosen),
                team_size_before,
                final_team_size,
                req_num,
                extra_max_sheet,
                highlight_surplus,
                ",".join(chosen),
            )

    return appended_total

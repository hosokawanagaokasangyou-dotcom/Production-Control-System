# -*- coding: utf-8 -*-
# planning_core.core.stage2_impl — body only (loaded via _core exec chain)
def _generate_plan_impl(
    tasks_df_override=None,
    stage2_output_root=None,
    skip_remove_prior_stage2_workbooks=False,
    return_output_paths=False,
    tasks_df_raw_input_baseline=None,
    result_pattern_shift_label=None,
    *,
    interactive_relax_intraday: bool = False,
    interactive_dispatch_targets: dict | None = None,
    interactive_result_dispatch_json_rows: list | None = None,
    interactive_result_dispatch_json_columns: list | None = None,
):
    # 配台トレース（設定シート A3:A26 のみ）はメンバー0人等で早期 return しても
    # execution_log に残るよご skills 読込より剝で確定・ログれる。
    global TRACE_SCHEDULE_TASK_IDS, DEBUG_DISPATCH_ONLY_TASK_IDS
    global _LAST_INTERACTIVE_TRIAL_METERS_DONE_SNAPSHOT, _LAST_INTERACTIVE_STAGE3_META
    global _LAST_INTERACTIVE_TRIAL_META_MISS_SHORTFALL
    global _LAST_INTERACTIVE_TRIAL_PLAN_TARGETS_SNAPSHOT
    global _LAST_INTERACTIVE_REMAINING_TASKS_AT_CALENDAR_END
    global _PLAN_IMPL_INTERACTIVE_DISPATCH_TARGETS, _PLAN_IMPL_INTERACTIVE_TRIAL_METERS_DONE
    global _PLAN_IMPL_INTERACTIVE_DISPATCH_TARGETS, _PLAN_IMPL_INTERACTIVE_TRIAL_METERS_DONE
    if interactive_relax_intraday or interactive_dispatch_targets is not None:
        _INTERACTIVE_TRIAL_OP_SHORTAGE.clear()
        _INTERACTIVE_TRIAL_AS_SHORTAGE.clear()
        _LAST_INTERACTIVE_TRIAL_METERS_DONE_SNAPSHOT.clear()
        _LAST_INTERACTIVE_TRIAL_META_MISS_SHORTFALL.clear()
        _LAST_INTERACTIVE_REMAINING_TASKS_AT_CALENDAR_END.clear()
        _LAST_INTERACTIVE_STAGE3_META.clear()
    _wb_trace = _excel_plan_input_wb()
    _ids_from_sheet = _read_trace_schedule_task_ids_from_config_sheet(_wb_trace)
    TRACE_SCHEDULE_TASK_IDS = frozenset(
        str(x).strip() for x in _ids_from_sheet if str(x).strip()
    )
    if _ids_from_sheet:
        _preview = _ids_from_sheet[:25]
        _suffix = " …" if len(_ids_from_sheet) > 25 else ""
        logging.info(
            "設定シート「%s」A3:A26: トレース用依頼NOを %s 件読み込み（%s%s）",
            APP_CONFIG_SHEET_NAME,
            len(_ids_from_sheet),
            ", ".join(_preview),
            _suffix,
        )
    else:
        logging.info(
            "設定シート「%s」A3:A26: トレース用依頼NOは無し（空またはシート無し）",
            APP_CONFIG_SHEET_NAME,
        )
    if TRACE_SCHEDULE_TASK_IDS:
        logging.info(
            "配台トレース: 有効 task_id = %s（設定シート A3:A26）",
            ", ".join(sorted(TRACE_SCHEDULE_TASK_IDS)),
        )
    else:
        logging.info(
            "配台トレース: 対象なし（[配台トレース …] ログは出ません）"
        )
    _ids_debug_dispatch_raw = _read_debug_dispatch_task_ids_from_config_sheet(_wb_trace)
    _dbg_norm: list[str] = []
    for _dx in _ids_debug_dispatch_raw:
        _dt = planning_task_id_str_from_scalar(_dx)
        if _dt:
            _dbg_norm.append(_dt)
    DEBUG_DISPATCH_ONLY_TASK_IDS = frozenset(_dbg_norm)
    if DEBUG_DISPATCH_ONLY_TASK_IDS:
        logging.warning(
            "デバッグ配台: 「%s」B3:B26 により配台対象を %s 件の依頼NOに限定しました: %s",
            APP_CONFIG_SHEET_NAME,
            len(DEBUG_DISPATCH_ONLY_TASK_IDS),
            ", ".join(sorted(DEBUG_DISPATCH_ONLY_TASK_IDS)),
        )
        _show_stage2_debug_dispatch_mode_dialog(sorted(DEBUG_DISPATCH_ONLY_TASK_IDS))
    if TRACE_TEAM_ASSIGN_TASK_ID:
        logging.info(
            "環境変数 TRACE_TEAM_ASSIGN_TASK_ID=%r → フォーム割当トレース有効",
            TRACE_TEAM_ASSIGN_TASK_ID,
        )

    _reset_dispatch_trace_per_task_logfiles()

    _t_s2_entry = time_module.perf_counter()
    (
        skills_dict,
        members,
        equipment_list,
        req_map,
        need_rules,
        surplus_map,
        need_combo_col_index,
    ) = load_skills_and_needs()
    _t_combo0 = _log_stage2_phase_timing("load_skills_and_needs", _t_s2_entry)
    team_combo_presets = load_team_combination_presets_from_master()
    if team_combo_presets:
        _nrules = sum(len(v) for v in team_combo_presets.values())
        logging.info(
            "組み合わせ表: 工程+機械キー %s 種類・編集行 %s を配台プリセットとして読み込みました。",
            len(team_combo_presets),
            _nrules,
        )
    elif TEAM_ASSIGN_USE_MASTER_COMBO_SHEET:
        logging.info(
            "組み合わせ表: プリセット無し（シート欠如・空・または読込失敗）。従来のフォーム探索のみ。"
        )
    if not members:
        master_abs = _master_workbook_path_resolved()
        logging.error(
            "段階2を中断しました: メンバーは0人です（マスタの skills は空」または読み込み失敗）。"
            " 期待パス: %s （カレント: %s）。テストコード直下に master.xlsm を置し」"
            "planning_core のカレントはしのフォルダになるよご python\\ 配置を確認してください。"
            " この状態では production_plan / member_schedule は出力されません。",
            master_abs,
            os.getcwd(),
        )
        return
    if _dispatch_postpone_only_policy_active():
        _mp_prereq = _master_workbook_path_resolved()
        _ctx = (
            "段階3配台試行"
            if _interactive_dispatch_trial_env_active()
            else "段階2"
        )
        _validate_master_dispatch_prerequisites(
            _mp_prereq, members, equipment_list, context_label=_ctx
        )
        if _interactive_dispatch_trial_env_active() and _interactive_stage2_parity_active():
            _LAST_INTERACTIVE_STAGE3_META["dispatch_failure_policy"] = (
                "machine_calendar_or_attendance_missing_only"
            )
            _LAST_INTERACTIVE_STAGE3_META["extend_attendance_calendar"] = (
                _stage2_extend_attendance_calendar_enabled()
            )
    global _MACHINE_CALENDAR_BLOCKS_BY_DATE
    global _MACHINE_CALENDAR_INTERACTIVE_DEFINED_SLOTS_BY_DATE
    global _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE
    global _STAGE2_MACHINE_DAILY_STARTUP_REQ_BY_MACHINE
    global _STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE
    global _STAGE2_REQUEST_SWITCH_PREP_BY_MACHINE
    global _STAGE2_BREAK_RESUME_PREP_BY_PROC_MACHINE
    global _STAGE2_BREAK_RESUME_PREP_BY_MACHINE
    global _STAGE2_POST_MACHINING_CLEANUP_BY_PROC_MACHINE
    global _STAGE2_POST_MACHINING_CLEANUP_BY_MACHINE
    global _STAGE2_REQUEST_INTERVAL_BUFFER_BY_PROC_MACHINE
    global _STAGE2_REQUEST_INTERVAL_BUFFER_BY_MACHINE
    global _STAGE2_REGULAR_SHIFT_START
    global _STAGE2_DATA_EXTRACTION_DATETIME
    global DEFAULT_START_TIME, DEFAULT_END_TIME
    try:
        _trial_env = _interactive_trial_calendar_legacy_active()
        _MACHINE_CALENDAR_BLOCKS_BY_DATE = load_machine_calendar_occupancy_blocks(
            _master_workbook_path_resolved(),
            equipment_list,
            interactive_only_asterisk_occupancy=_trial_env,
        )
    except Exception as e:
        logging.warning(
            "機械カレンダー: 読込例外のため、占有なしとして続行しした (%s)", e
        )
        _MACHINE_CALENDAR_BLOCKS_BY_DATE = {}
        _MACHINE_CALENDAR_INTERACTIVE_DEFINED_SLOTS_BY_DATE = {}
    _t_cal0 = _log_stage2_phase_timing("load_team_combination_presets", _t_combo0)
    try:
        (
            _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE,
            _STAGE2_MACHINE_DAILY_STARTUP_REQ_BY_MACHINE,
        ) = load_machine_daily_startup_settings(
            _master_workbook_path_resolved()
        )
    except Exception as e:
        logging.warning(
            "機械日次始業準備設定: 読込例外のため、無視しした (%s)", e
        )
        _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE = {}
        _STAGE2_MACHINE_DAILY_STARTUP_REQ_BY_MACHINE = {}
    try:
        (
            _STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE,
            _STAGE2_REQUEST_SWITCH_PREP_BY_MACHINE,
            _STAGE2_BREAK_RESUME_PREP_BY_PROC_MACHINE,
            _STAGE2_BREAK_RESUME_PREP_BY_MACHINE,
            _STAGE2_POST_MACHINING_CLEANUP_BY_PROC_MACHINE,
            _STAGE2_POST_MACHINING_CLEANUP_BY_MACHINE,
            _STAGE2_REQUEST_INTERVAL_BUFFER_BY_PROC_MACHINE,
            _STAGE2_REQUEST_INTERVAL_BUFFER_BY_MACHINE,
        ) = load_request_switch_prep_settings(_master_workbook_path_resolved())
    except Exception as e:
        logging.warning(
            "依頼切替準備設定: 読込例外のため無視します (%s)", e
        )
        _STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE = {}
        _STAGE2_REQUEST_SWITCH_PREP_BY_MACHINE = {}
        _STAGE2_BREAK_RESUME_PREP_BY_PROC_MACHINE = {}
        _STAGE2_BREAK_RESUME_PREP_BY_MACHINE = {}
        _STAGE2_POST_MACHINING_CLEANUP_BY_PROC_MACHINE = {}
        _STAGE2_POST_MACHINING_CLEANUP_BY_MACHINE = {}
        _STAGE2_REQUEST_INTERVAL_BUFFER_BY_PROC_MACHINE = {}
        _STAGE2_REQUEST_INTERVAL_BUFFER_BY_MACHINE = {}
    _t_ds0 = _log_stage2_phase_timing("load_machine_calendar_and_master_settings", _t_cal0)
    _master_path_stage2 = _master_workbook_path_resolved()
    if any(int(v or 0) > 0 for v in _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE.values()):
        _a12s_chk, _a12e_chk = _read_master_main_factory_operating_times(_master_path_stage2)
        if _a12s_chk is None or _a12e_chk is None:
            raise PlanningValidationError(
                "日次始業準備に有効な準備時間（分）が登録されているため、"
                "master.xlsm メインの A12（工場稼働開始）・B12（工場稼働終了）を正しく設定してください。"
                "（欠損・開始>=終了・読込不可のときは配台を中止します）"
            )
    try:
        _rs_a15, _ = _read_master_main_regular_shift_times(_master_path_stage2)
        _STAGE2_REGULAR_SHIFT_START = _rs_a15
        if _rs_a15 is not None:
            logging.info(
                "定常枠: master メイン A15=%s（結果シートの定常外着色・出勤簿生成の参照）",
                _rs_a15.strftime("%H:%M"),
            )
    except Exception as e:
        logging.warning("定常開始(A15) 読込失敗: 結果の定常外着色等で参照しません (%s)", e)
        _STAGE2_REGULAR_SHIFT_START = None
    if _MACHINE_CALENDAR_BLOCKS_BY_DATE:
        _n_iv = sum(
            len(ivs)
            for _dm in _MACHINE_CALENDAR_BLOCKS_BY_DATE.values()
            for ivs in _dm.values()
        )
        logging.info(
            "機械カレンダー: %s 日分・設備占有ブロック合計 %s 件を配台に反映しました。",
            len(_MACHINE_CALENDAR_BLOCKS_BY_DATE),
            _n_iv,
        )
    if interactive_relax_intraday:
        _MACHINE_CALENDAR_BLOCKS_BY_DATE = {}
        _MACHINE_CALENDAR_INTERACTIVE_DEFINED_SLOTS_BY_DATE = {}
        DEFAULT_START_TIME = time(0, 0)
        DEFAULT_END_TIME = time(23, 59)
        logging.warning(
            "インタラクティブ配台試行: 工場枠を全日(00:00-23:59)に拡大し、"
            "機械カレンダー占有を無視します（サブプロセス終了で既定値に戻ります）。"
        )
    elif _interactive_trial_calendar_legacy_active():
        # デスクトップ配台試行（dispatch_interactive_trial）既定: 全日緩和は使わず master 工場枠＋同日延長
        _mp_cal = _master_path_stage2
        _ns, _ne = _read_master_main_factory_operating_times(_mp_cal)
        if _ns is not None and _ne is not None:
            DEFAULT_START_TIME = _ns
            DEFAULT_END_TIME = _ne
            logging.info(
                "インタラクティブ配台試行: 工場稼働は master メイン A12/B12（%s～%s）。",
                _ns.strftime("%H:%M"),
                _ne.strftime("%H:%M"),
            )
        DEFAULT_END_TIME = time(23, 59)
        logging.info(
            "インタラクティブ配台試行: 同一暦日内の割付のため終業を 23:59 まで拡張（"
            "開始は上記のまま。暦日をまたぐ加工はエラーで試行中止）。"
        )
    reset_gemini_usage_tracker()
    _clear_stage2_blocking_message_file()
    if (
        not TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
        and not TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW
    ):
        logging.info(
            "need配台時追加人数: メイン割付は基本必須人数のみ。"
            "余力は全シミュレーション後」未割当かつスキル保有社をサブに追記しました。"
            "（メインで増員探索れる従来挙動: TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS=1）"
        )

    # 段階2の基準日時は「マクロ実行時刻」ではなく加工計画DATA「データ抽出時間」（なければ「抽出時間」→「データ抽出日」）
    data_extract_dt, plan_base_dt_column = _extract_data_extraction_datetime()
    _STAGE2_DATA_EXTRACTION_DATETIME = data_extract_dt
    base_now_dt = data_extract_dt if data_extract_dt is not None else datetime.now()
    # 表示・ファイル名・メタ用の「データ抽出」文字列は、正規化前の抽出時刻（加工計画DATA上の値）を維持する。
    # base_now_dt は当日判定や探索下限の都合で正規化されうるため、用途を混ぜない。
    data_extract_dt_display_str = (
        data_extract_dt.strftime("%Y/%m/%d %H:%M:%S") if data_extract_dt is not None else "—"
    )
    # データ抽出時刻が「夕方」でも、同日の配台探索が終業前まで進むよう macro_now の時刻を正規化する（既定）。
    # 旧挙動が必要な場合のみ STAGE2_MACRO_NOW_USE_DATA_EXTRACT_CLOCK=1。
    if (
        data_extract_dt is not None
        and not STAGE2_MACRO_NOW_USE_DATA_EXTRACT_CLOCK
        and isinstance(base_now_dt, datetime)
    ):
        _orig_macro = base_now_dt
        base_now_dt = datetime.combine(base_now_dt.date(), DEFAULT_START_TIME)
        if _orig_macro != base_now_dt:
            logging.info(
                "計画基準日時(正規化): データ抽出の日付は維持し、時刻を %s に寄せました（元=%s）。"
                "旧挙動は STAGE2_MACRO_NOW_USE_DATA_EXTRACT_CLOCK=1。",
                base_now_dt.strftime("%Y/%m/%d %H:%M:%S"),
                _orig_macro.strftime("%Y/%m/%d %H:%M:%S"),
            )
    run_date = base_now_dt.date()
    calendar_today = run_date
    if _stage2_truthy_env("PM_AI_STAGE2_SKIP_TODAY_DISPATCH"):
        run_date = run_date + timedelta(days=1)
        logging.info(
            "段階2: PM_AI_STAGE2_SKIP_TODAY_DISPATCH により当日は配台せず、計画開始日を %s にずらしました。",
            run_date.isoformat(),
        )
    data_extract_dt_str = (
        data_extract_dt_display_str if data_extract_dt is not None else "—"
    )
    logging.info(
        "計画基準日時: %s（%s）",
        base_now_dt.strftime("%Y/%m/%d %H:%M:%S"),
        plan_base_dt_column if data_extract_dt is not None else "現在時刻フォールバック",
    )

    _t_att0 = _log_stage2_phase_timing("extract_data_extraction_datetime", _t_ds0)
    attendance_data, ai_log_data = load_attendance_and_analyze(members)
    _t_gpo0 = _log_stage2_phase_timing("load_attendance_and_analyze", _t_att0)
    global_priority_raw = load_main_sheet_global_priority_override_text()
    global_priority_override = analyze_global_priority_override_comment(
        global_priority_raw,
        members,
        run_date.year,
        ai_sheet_sink=ai_log_data,
    )
    _factory_closure_dates: set[date] = set()
    for _iso in global_priority_override.get("factory_closure_dates") or []:
        _d = parse_optional_date(_iso)
        if _d is not None:
            _factory_closure_dates.add(_d)
    if _factory_closure_dates:
        apply_factory_closure_dates_to_attendance(
            attendance_data, members, _factory_closure_dates
        )
        logging.info(
            "メイン・グローバルコメント: 工場休業扱いの日付 → %s",
            ", ".join(str(x) for x in sorted(_factory_closure_dates)),
        )
    if apply_overtime_simulation_overrides(attendance_data):
        ai_log_data["残業シミュレーション"] = (
            f"PM_AI_OVERTIME_SIMULATION_JSON を適用 ({os.environ.get(ENV_OVERTIME_SIMULATION_JSON, '')})"
        )
    ai_log_data["メイン_グローバル_工場休業日(解析)"] = (
        ", ".join(str(x) for x in sorted(_factory_closure_dates))
        if _factory_closure_dates
        else "（なし）"
    )
    _sn = str(global_priority_override.get("scheduler_notes_ja") or "").strip()
    if _sn:
        ai_log_data["メイン_グローバル_未適用メモ(AI)"] = _sn[:2000]

    sorted_dates = sorted(list(attendance_data.keys()))
    # 結果シートは「基準日（データ抽出時間／抽出時間／データ抽出日）」以降のみ表示・計画対象とする
    sorted_dates = [d for d in sorted_dates if d >= run_date]
    if not sorted_dates:
        logging.error("当日以降の処理対象日付はありません。")
        return

    # タスク入力: ブック内「配台計画_タスク入力」（段階1で出力→取り込み後に編集）
    try:
        if tasks_df_override is not None:
            tasks_df = tasks_df_override.copy()
        else:
            tasks_df = load_planning_tasks_df()
    except PlanningValidationError:
        raise
    except Exception as e:
        logging.error(f"配台計画タスクシート読み込みエラー: {e}")
        return

    if DEBUG_DISPATCH_ONLY_TASK_IDS:
        _n_tasks_before = len(tasks_df)
        _dbg_mask = tasks_df.apply(
            lambda row: planning_task_id_str_from_plan_row(row)
            in DEBUG_DISPATCH_ONLY_TASK_IDS,
            axis=1,
        )
        tasks_df = tasks_df.loc[_dbg_mask].copy()
        _n_tasks_after = len(tasks_df)
        logging.warning(
            "デバッグ配台: 「%s」の行を %s → %s に絞り込みました。",
            PLAN_INPUT_SHEET_NAME,
            _n_tasks_before,
            _n_tasks_after,
        )
        if _n_tasks_after == 0:
            logging.error(
                "デバッグ配台: B3:B26 の依頼NOに一致する行はありません。段階2を中断しました。"
            )
            return

    if global_priority_raw.strip():
        snip = global_priority_raw[:2500]
        if len(global_priority_raw) > 2500:
            snip += "…"
        ai_log_data["メイン_再優先特別記載(原文)"] = snip
    else:
        ai_log_data["メイン_再優先特別記載(原文)"] = (
            "（空」またはメインシートに「グローバルコメント」見出しは見つかりません）"
        )
    ai_log_data["メイン_再優先特別記載(AI)"] = json.dumps(
        global_priority_override, ensure_ascii=False
    )
    if global_priority_override.get("ignore_skill_requirements"):
        logging.warning(
            "メイン再優先特記: スキル覝件を無視して配台しした。%s",
            global_priority_override.get("interpretation_ja", ""),
        )
    if global_priority_override.get("ignore_need_minimum"):
        logging.warning(
            "メイン再優先特記: フォーム人数を1坝に固定しした（need・行の必須OP上書きより優先）。%s",
            global_priority_override.get("interpretation_ja", ""),
        )
    if global_priority_override.get("abolish_all_scheduling_limits"):
        logging.warning(
            "メイン再優先特記: 設備専有・原板同日開始・指定開始時刻・マクロ実行時刻下限を適用しません。%s",
            global_priority_override.get("interpretation_ja", ""),
        )

    # 「当日」判定と最早開始時刻には基準日時（データ抽出時間→抽出時間→データ抽出日）を使う。
    # base_now_dt は「同日を広く探索する」目的で時刻が工場開始へ正規化されうるが、
    # 抽出時刻より前に配台しない制約（要件: 抽出時間を基準に遡らない）には正規化前を使う。
    macro_now_dt = data_extract_dt if isinstance(data_extract_dt, datetime) else base_now_dt
    macro_run_date = macro_now_dt.date()
    ai_task_by_tid = analyze_task_special_remarks(
        tasks_df, reference_year=run_date.year,         ai_sheet_sink=ai_log_data
    )
    _t_tq0 = _log_stage2_phase_timing("analyze_global_priority_and_task_remarks", _t_gpo0)
    if _stage2_truthy_env("PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH"):
        logging.info(
            "段階2: PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH — 実加工数が正の行は配台キューに入れません（当日完了と想定）。"
        )
    task_queue = build_task_queue_from_planning_df(
        tasks_df,
        run_date,
        req_map,
        ai_task_by_tid,
        global_priority_override,
        equipment_list,
    )
    # 開始日は非稼働日の場合は」直後の稼働日へ補正（例: 4/4, 4/5 は非稼働なら 4/3 へ）
    working_days = [
        d for d in sorted_dates
        if any(attendance_data[d][m]["is_working"] for m in attendance_data[d])
    ]
    if _stage2_truthy_env("PM_AI_STAGE2_SKIP_TODAY_DISPATCH") and working_days:
        if run_date not in working_days:
            _orig_run_for_snap = run_date
            for _wd_snap in working_days:
                if _wd_snap >= run_date:
                    run_date = _wd_snap
                    logging.info(
                        "段階2: 計画開始日 %s は非稼働のため %s に繰り上げました（SKIP_TODAY）。",
                        _orig_run_for_snap.isoformat(),
                        run_date.isoformat(),
                    )
                    break
            sorted_dates = [d for d in sorted_dates if d >= run_date]
    if working_days:
        for t in task_queue:
            req_d = t.get("start_date_req")
            if not isinstance(req_d, date):
                continue
            if req_d in working_days:
                continue
            prev_work = None
            for wd in working_days:
                if wd <= req_d:
                    prev_work = wd
                else:
                    break
            if prev_work is not None:
                if str(t.get("task_id", "")).strip() == DEBUG_TASK_ID:
                    logging.info(
                        "DEBUG[task=%s] start_date_req を非稼働日補正: %s -> %s",
                        DEBUG_TASK_ID,
                        req_d,
                        prev_work,
                    )
                t["start_date_req"] = prev_work
    conflict_rows = collect_planning_conflicts_by_excel_row(tasks_df, ai_task_by_tid)
    _try_write_plan_input_global_parse_and_conflicts_one_save(
        global_priority_override,
        data_extract_dt_str,
        len(tasks_df),
        conflict_rows,
        tasks_df=tasks_df,
    )

    if not task_queue:
        logging.warning(
            f"有効なタスクはありません。「{PLAN_INPUT_SHEET_NAME}」の「依頼NO」「工程名」「{TASK_COL_QTY}」"
            "または完了区分・実出来高残作により残量は無い行のみの可能性はありした。"
        )

    # 配台試行順: シート列は权っていれみしれを採用。欠損時は §B 帯・紝期・need 列順でソートし EC 隣接後に 1..n
    _apply_dispatch_trial_order_for_generate_plan(
        task_queue, req_map, need_rules, need_combo_col_index
    )
    # 段階3: 手動修正 JSON で正の「当日配台数量」がある配台日の最古を start_date_req 下限とする
    # （原反投入・既定開始より前にはしない: max(既存, 最古暦日)。非稼働日は直前稼働日へ寄せる）
    # 段階2同一パリティ試行では適用しない（配台試行順・原反投入日と段階2を揃える。JSON 暦日はキャップのみ）。
    if (
        _interactive_dispatch_trial_env_active()
        and not _interactive_stage2_parity_active()
        and isinstance(interactive_result_dispatch_json_rows, list)
        and interactive_result_dispatch_json_rows
    ):
        for _t_iv in task_queue:
            _min_j = _interactive_min_positive_dispatch_date_from_json_rows(
                interactive_result_dispatch_json_rows,
                str(_t_iv.get("task_id") or ""),
                str(_t_iv.get("machine") or ""),
                str(_t_iv.get("machine_name") or ""),
            )
            if _min_j is None:
                continue
            _cur_iv = _t_iv.get("start_date_req")
            if isinstance(_cur_iv, date):
                _t_iv["start_date_req"] = max(_cur_iv, _min_j)
            else:
                _t_iv["start_date_req"] = max(run_date, _min_j)
        if working_days:
            for _t_iv in task_queue:
                req_d = _t_iv.get("start_date_req")
                if not isinstance(req_d, date):
                    continue
                if req_d in working_days:
                    continue
                prev_work = None
                for wd in working_days:
                    if wd <= req_d:
                        prev_work = wd
                    else:
                        break
                if prev_work is not None:
                    _t_iv["start_date_req"] = prev_work
    if DEBUG_TASK_ID:
        dbg_items = [t for t in task_queue if str(t.get("task_id", "")).strip() == DEBUG_TASK_ID]
        if dbg_items:
            t0 = dbg_items[0]
            logging.info(
                "DEBUG[task=%s] queue基準: start_date_req=%s due_basis=%s answer_due=%s specified_due=%s specified_due_ov=%s due_source=%s priority=%s in_progress=%s remark=%s",
                DEBUG_TASK_ID,
                t0.get("start_date_req"),
                t0.get("due_basis_date"),
                t0.get("answer_due_date"),
                t0.get("specified_due_date"),
                t0.get("specified_due_override"),
                t0.get("due_source"),
                t0.get("priority"),
                t0.get("in_progress"),
                t0.get("has_special_remark"),
            )
        else:
            logging.info("DEBUG[task=%s] task_queueに存在しません（完了/残量0/依頼NO厳密一致の可能性）。", DEBUG_TASK_ID)
    timeline_events = []

    _t_dispatch0 = _log_stage2_phase_timing(
        "build_task_queue_and_prepare_dispatch", _t_tq0
    )
    _reset_machine_handoff_timeline_cache()
    _reset_dispatch_loop_profile()
    _dispatch_sched_pass_total = 0
    _dispatch_day_timing: list[dict] = []

    # ---------------------------------------------------------
    # 日毎のスケジューリングループ
    # 勤怠自動拡張は _stage2_extend_attendance_calendar_enabled() が True のときのみ（段階2標準・段階3パリティは False）。
    # STAGE2_RETRY_SHIFT_DUE_ON_PARTIAL_REMAINING は True のときのみ: 紝期基準を靎ねでも残はある依頼についで
    # due_basis +1・当該依頼の割当戻し・先頭から再実行。坄再試行剝に勤怠拡張分はマスタ日付へ巻し戻れ。
    # 既定 False のため、通常は 1 パス（カレンダー通し 1 回）のみ。
    # ---------------------------------------------------------
    _master_attendance_date_set = frozenset(attendance_data.keys())
    _master_plan_dates_template = list(sorted_dates)
    _calendar_last_plan_day = _master_plan_dates_template[-1]

    for t in task_queue:
        t["remaining_units"] = float(t.get("initial_remaining_units") or 0)
        t["assigned_history"].clear()
    timeline_events.clear()

    _dispatch_interval_mirror: DispatchIntervalMirror | None = None
    if DISPATCH_INTERVAL_MIRROR_ENFORCE:
        _dispatch_interval_mirror = DispatchIntervalMirror()
        logging.info(
            "DISPATCH_INTERVAL_MIRROR_ENFORCE: 設備・人の占有を区間ミラーで追跡しした"
            "（無効化は 設定_環境変数 等で DISPATCH_INTERVAL_MIRROR_ENFORCE=0）。"
        )

    if STAGE2_SERIAL_DISPATCH_BY_TASK_ID:
        logging.info(
            "依頼NO直列配台: 有効（STAGE2_SERIAL_DISPATCH_BY_TASK_ID）。"
            " 坄日はアクティブな依頼NOの行の値は候補のため、当該依頼は詰まると他依頼は一切進みません。"
        )
    else:
        logging.info(
            "依頼NO直列配台: 無効。start_date を満たれ全行は当日候補になり」配台試行順・設備ルールで順庝付けしした。"
        )

    _due_shift_retry_count_by_request: dict[str, int] = {}
    _due_shift_exhausted_requests: set[str] = set()
    _due_shift_cap_warned_tids: set[str] = set()
    _interactive_trial_pair_dates = None
    _interactive_trial_meters_done: dict[tuple[str, str, str, date], float] = {}
    _PLAN_IMPL_INTERACTIVE_DISPATCH_TARGETS = (
        interactive_dispatch_targets
        if _interactive_dispatch_trial_env_active()
        else None
    )
    _PLAN_IMPL_INTERACTIVE_TRIAL_METERS_DONE = _interactive_trial_meters_done
    if _overtime_simulation_dispatch_trial_active():
        _apply_stage35_stage3_meters_floor(task_queue, _interactive_trial_meters_done)
    if (
        _interactive_dispatch_trial_env_active()
        and not _interactive_stage2_parity_active()
        and interactive_dispatch_targets
    ):
        # 段階3: JSON の暦日集合だけへ eligible を絞ると、依存で遅れた暦日が集合に無く
        # 全日スキップされうる（例: V5-4）。日次目標 m は interactive_dispatch_targets のキャップで担保。
        # 従来の暦日絞り: PM_AI_INTERACTIVE_TRIAL_PAIR_DATES=1
        if (os.environ.get("PM_AI_INTERACTIVE_TRIAL_PAIR_DATES") or "").strip().lower() in (
            "1",
            "true",
            "yes",
            "on",
        ):
            _interactive_trial_pair_dates = _interactive_trial_pair_dates_from_targets(
                interactive_dispatch_targets
            )
    _outer_retry_round = 0
    while True:
        _dispatch_trace_begin_outer_round(_outer_retry_round)
        _need_headcount_logged_orders: set = set()
        if _outer_retry_round > 0:
            _purge_attendance_days_not_in_set(
                attendance_data, _master_attendance_date_set
            )
            sorted_dates[:] = list(_master_plan_dates_template)

        for t in task_queue:
            t.pop("_partial_retry_calendar_blocked", None)

        if STAGE2_SERIAL_DISPATCH_BY_TASK_ID:
            _serial_order_tids = _serial_dispatch_order_task_ids(task_queue)
        else:
            _serial_order_tids = []

        _plan_day_iter = (
            _iter_plan_dates_extending(sorted_dates, attendance_data, task_queue)
            if _stage2_extend_attendance_calendar_enabled()
            else sorted_dates
        )
        global _STAGE2_ALADDIN_EXCLUDE_APPLY_DATE
        _STAGE2_ALADDIN_EXCLUDE_APPLY_DATE = _stage2_dialog_target_plan_day(
            run_date,
            sorted_dates,
            skip_today=_stage2_truthy_env("PM_AI_STAGE2_SKIP_TODAY_DISPATCH"),
        )
        _full_calendar_without_deadline_restart = True
        for current_date in _plan_day_iter:
            _t_day0 = time_module.perf_counter()
            _dispatch_loop_profile_begin_day(current_date.isoformat())
            _day_sched_passes = 0
            daily_status = attendance_data[current_date]
            # 設備ととの空し時刻（同一設備の坌時並行割当を防止）
            machine_avail_dt = {}
            
            avail_dt = {}
            for m in members:
                if m not in daily_status:
                    continue
                st = daily_status[m]
                if st.get("eligible_for_assignment", st.get("is_working", False)):
                    avail_dt[m] = st["start_dt"]

            _machine_day_start = datetime.combine(current_date, DEFAULT_START_TIME)
            _machine_calendar_plan_end = _machine_calendar_planning_window_end_dt(
                current_date, daily_status, members
            )
            if avail_dt:
                _day_timeline_ev = (
                    _STAGE2_DISPATCH_EVENTS_BY_DATE.get(current_date)
                    if _STAGE2_DISPATCH_EVENTS_BY_DATE is not None
                    else None
                )
                _t_seed0 = time_module.perf_counter()
                _seed_avail_from_timeline_for_date(
                    timeline_events,
                    current_date,
                    machine_avail_dt,
                    avail_dt,
                    _machine_day_start,
                    events_today=_day_timeline_ev,
                )
                _dispatch_loop_profile_add(
                    "day_seed_avail", time_module.perf_counter() - _t_seed0
                )
                _t_mcal0 = time_module.perf_counter()
                _apply_machine_calendar_floor_for_date(
                    current_date,
                    machine_avail_dt,
                    equipment_list,
                    _machine_day_start,
                    machine_calendar_plan_end=_machine_calendar_plan_end,
                )
                _dispatch_loop_profile_add(
                    "day_machine_calendar", time_module.perf_counter() - _t_mcal0
                )

            if not avail_dt:
                logging.info("DEBUG[day=%s] 稼働メンバー0のため、割付スキップ", current_date)
                _dispatch_day_timing.append(
                    {
                        "date": current_date.isoformat(),
                        "elapsed_sec": round(
                            time_module.perf_counter() - _t_day0, 3
                        ),
                        "sched_passes": 0,
                        "sched_max_passes": 0,
                        "tasks_today": 0,
                        "skipped_no_members": True,
                        "profile": _dispatch_loop_profile_finish_day(
                            time_module.perf_counter() - _t_day0,
                            skipped_no_members=True,
                        ),
                    }
                )
                continue
    
            tasks_today = [t for t in task_queue if t['remaining_units'] > 0 and t['start_date_req'] <= current_date]
            if STAGE2_SERIAL_DISPATCH_BY_TASK_ID and _serial_order_tids:
                _tasks_today_before_serial = len(tasks_today)
                _active_serial_tid = None
                for _tid in _serial_order_tids:
                    if any(
                        float(x.get("remaining_units") or 0) > 1e-12
                        for x in task_queue
                        if str(x.get("task_id", "") or "").strip() == _tid
                    ):
                        _active_serial_tid = _tid
                        break
                if _active_serial_tid is not None:
                    tasks_today = [
                        t
                        for t in tasks_today
                        if str(t.get("task_id", "") or "").strip() == _active_serial_tid
                    ]
                _serial_pos = (
                    _serial_order_tids.index(_active_serial_tid) + 1
                    if _active_serial_tid in _serial_order_tids
                    else 0
                )
                _pending_rows = sum(1 for t in task_queue if t["remaining_units"] > 0)
                logging.info(
                    "依頼NO直列配台 day=%s アクティブ依頼NO=%s 直列リスト佝置=%s/%s "
                    "当日候補行数(直列剝)=%s 直列後=%s キュー残行(全日)=%s",
                    current_date,
                    _active_serial_tid if _active_serial_tid is not None else "—",
                    _serial_pos if _serial_pos else "—",
                    len(_serial_order_tids),
                    _tasks_today_before_serial,
                    len(tasks_today),
                    _pending_rows,
                )
            pending_total = sum(1 for t in task_queue if t["remaining_units"] > 0)
            if not tasks_today:
                earliest_wait = min(
                    [t["start_date_req"] for t in task_queue if t["remaining_units"] > 0],
                    default=None,
                )
                logging.info(
                    "DEBUG[day=%s] 割付対象タスク0件 pending_total=%s earliest_start_date_req=%s",
                    current_date,
                    pending_total,
                    earliest_wait,
                )
                _day_elapsed_skip = time_module.perf_counter() - _t_day0
                _dispatch_day_timing.append(
                    {
                        "date": current_date.isoformat(),
                        "elapsed_sec": round(_day_elapsed_skip, 3),
                        "sched_passes": 0,
                        "sched_max_passes": 0,
                        "tasks_today": 0,
                        "skipped_empty_tasks": True,
                        "profile": _dispatch_loop_profile_finish_day(
                            _day_elapsed_skip,
                            skipped_empty_tasks=True,
                            pending_total=pending_total,
                        ),
                    }
                )
                continue
            elif DEBUG_TASK_ID:
                has_dbg_today = any(str(t.get("task_id", "")).strip() == DEBUG_TASK_ID for t in tasks_today)
                if current_date.isoformat() == "2026-04-03" or has_dbg_today:
                    logging.info(
                        "DEBUG[day=%s] avail_members=%s tasks_today=%s (task=%s 含む=%s)",
                        current_date,
                        len(avail_dt),
                        len(tasks_today),
                        DEBUG_TASK_ID,
                        has_dbg_today,
                    )
            
            _sched_max_passes = max(96, max(1, len(tasks_today)) * 15)
            if (
                _interactive_dispatch_cap_enforced_in_schedule_loop()
                and interactive_dispatch_targets
                and STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST
            ):
                _t_cap_pass0 = time_module.perf_counter()
                _trial_order_first_schedule_pass(
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
                    dispatch_interval_mirror=_dispatch_interval_mirror,
                    interactive_dispatch_targets=interactive_dispatch_targets,
                    interactive_trial_pair_dates=_interactive_trial_pair_dates,
                    interactive_trial_meters_done=_interactive_trial_meters_done,
                    cap_drain_only=True,
                )
                _cap_pass_sec = time_module.perf_counter() - _t_cap_pass0
                _dispatch_loop_profile_add("trial_pass_wall_cap", _cap_pass_sec)
                if _STAGE2_DISPATCH_LOOP_PROFILE_DAY is not None:
                    _STAGE2_DISPATCH_LOOP_PROFILE_DAY["trial_pass_count"] = (
                        _STAGE2_DISPATCH_LOOP_PROFILE_DAY.get("trial_pass_count", 0)
                        + 1
                    )
            _sched_pi = 0
            while _sched_pi < _sched_max_passes:
                _sched_pi += 1
                _day_sched_passes += 1
                _dispatch_sched_pass_total += 1
                _sched_made_progress = False
                _t_sched_pass0 = time_module.perf_counter()
                if STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST:
                    _sched_made_progress = _trial_order_first_schedule_pass(
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
                        dispatch_interval_mirror=_dispatch_interval_mirror,
                        interactive_dispatch_targets=interactive_dispatch_targets,
                        interactive_trial_pair_dates=_interactive_trial_pair_dates,
                        interactive_trial_meters_done=_interactive_trial_meters_done,
                    )
                    _sched_pass_sec = time_module.perf_counter() - _t_sched_pass0
                    _dispatch_loop_profile_add("trial_pass_wall", _sched_pass_sec)
                    if _STAGE2_DISPATCH_LOOP_PROFILE_DAY is not None:
                        _STAGE2_DISPATCH_LOOP_PROFILE_DAY["trial_pass_count"] = (
                            _STAGE2_DISPATCH_LOOP_PROFILE_DAY.get("trial_pass_count", 0)
                            + 1
                        )
                        _STAGE2_DISPATCH_LOOP_PROFILE_DAY.setdefault(
                            "sched_pass_secs", []
                        ).append(_sched_pass_sec)
                if not STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST:
                    _t_legacy0 = time_module.perf_counter()
                    _mh_legacy_day = _machine_handoff_state_from_timeline(
                        timeline_events, current_date
                    )
                    machine_handoff_legacy = {
                        "last_tid": dict(_mh_legacy_day["last_tid"]),
                        "last_eq": dict(_mh_legacy_day["last_eq"]),
                        "started_today": set(_mh_legacy_day["started_today"]),
                        "machining_today_occ": set(
                            _mh_legacy_day.get("machining_today_occ") or set()
                        ),
                        "last_machining_dt": dict(
                            _mh_legacy_day.get("last_machining_dt") or {}
                        ),
                        "last_machining_date": dict(
                            _mh_legacy_day.get("last_machining_date") or {}
                        ),
                        "last_lead_op": dict(_mh_legacy_day.get("last_lead_op") or {}),
                        "last_machining_sub": dict(
                            _mh_legacy_day.get("last_machining_sub") or {}
                        ),
                    }
                    _assign_probe_ctx_legacy: dict | None = None
                    _min_dispatch_eff_legacy: int | None = None
                    if STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT:
                        _assign_probe_ctx_legacy = {
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
                            "machine_handoff": machine_handoff_legacy,
                            "team_combo_presets": team_combo_presets,
                            "dispatch_interval_mirror": _dispatch_interval_mirror,
                        }
                        _pool_legacy = _tasks_in_min_pending_dispatch_pool(
                            task_queue,
                            current_date,
                            daily_status=daily_status,
                            members=members,
                            machine_avail_dt=machine_avail_dt,
                            machine_day_start=_machine_day_start,
                            machine_handoff=machine_handoff_legacy,
                            skills_dict=skills_dict,
                            abolish_all_scheduling_limits=bool(
                                global_priority_override.get(
                                    "abolish_all_scheduling_limits"
                                )
                            ),
                            dispatch_interval_mirror=_dispatch_interval_mirror,
                            interactive_trial_pair_dates=_interactive_trial_pair_dates,
                        )
                        _min_dispatch_eff_legacy = (
                            _effective_min_dispatch_trial_order_from_pool(
                                _pool_legacy,
                                current_date,
                                daily_status,
                                _assign_probe_ctx_legacy,
                            )
                        )
                    for task in sorted(
                        [t for t in tasks_today if float(t.get("remaining_units") or 0) > 1e-12],
                        key=lambda t: _day_schedule_task_sort_key(
                            t, task_queue, need_combo_col_index
                        ),
                    ):
                        if _task_blocked_by_same_request_dependency(task, task_queue):
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _log_dispatch_trace_schedule(
                                    task.get("task_id"),
                                    "[配台トレース task=%s] スキップ: 同一依頼NOの先行工程待う day=%s machine=%s rem=%.4f",
                                    task.get("task_id"),
                                    current_date,
                                    task.get("machine"),
                                    float(task.get("remaining_units") or 0),
                                )
                            continue
                        if (
                            task.get("roll_pipeline_inspection")
                            or task.get("roll_pipeline_rewind")
                        ) and (
                            _roll_pipeline_inspection_assign_room(
                                task_queue, str(task.get("task_id", "")).strip()
                            )
                            <= 1e-12
                        ):
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _tid_tr = str(task.get("task_id", "") or "").strip()
                                _ec_d = _pipeline_ec_roll_done_units(task_queue, _tid_tr)
                                _in_d = _pipeline_b2_follower_roll_done_units(
                                    task_queue, _tid_tr
                                )
                                _log_dispatch_trace_schedule(
                                    _tid_tr,
                                    "[配台トレース task=%s] スキップ: §B-2/§B-3 後続ロール枠ゼロ day=%s machine=%s "
                                    "ec累計完了R=%.4f 後続累計完了R=%.4f rem_follower=%.4f",
                                    _tid_tr,
                                    current_date,
                                    task.get("machine"),
                                    _ec_d,
                                    _in_d,
                                    float(task.get("remaining_units") or 0),
                                )
                            continue
                        machine = task["machine"]
                        eq_line = str(
                            task.get("equipment_line_key") or machine or ""
                        ).strip() or machine
                        machine_occ_key = _machine_occupancy_key_resolve(task, eq_line)
                        if PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE:
                            _b1_holder = _exclusive_b1_inspection_holder_for_machine(
                                task_queue,
                                machine_occ_key,
                            )
                            if _b1_holder is not None and _b1_holder is not task:
                                if _trace_schedule_task_enabled(task.get("task_id")):
                                    _log_dispatch_trace_schedule(
                                        task.get("task_id"),
                                        "[配台トレース task=%s] スキップ: 同一設備の検査占有中 day=%s "
                                        "占有者依頼NO=%s 占有者試行順=%s",
                                        task.get("task_id"),
                                        current_date,
                                        _b1_holder.get("task_id"),
                                        _b1_holder.get("dispatch_trial_order"),
                                    )
                                continue
                        if DEBUG_TASK_ID and str(task.get("task_id", "")).strip() == DEBUG_TASK_ID:
                            logging.info(
                                "DEBUG[task=%s] day=%s 開始判定: start_date_req=%s remaining_units=%s machine=%s",
                                DEBUG_TASK_ID,
                                current_date,
                                task.get("start_date_req"),
                                task.get("remaining_units"),
                                task.get("machine"),
                            )
                        if task.get("has_done_deadline_override"):
                            logging.info(
                                "DEBUG[完了日指定] 依頼NO=%s 日付=%s start_date_req=%s due_basis=%s 指定納期(上書き)=%s 進杗=%s/%s",
                                task.get("task_id"),
                                current_date,
                                task.get("start_date_req"),
                                task.get("due_basis_date"),
                                task.get("specified_due_override"),
                                task.get("done_qty_reported"),
                                task.get("total_qty_m"),
                            )
                        try:
                            _my_dispatch_ord = int(
                                task.get("dispatch_trial_order") or 10**9
                            )
                        except (TypeError, ValueError):
                            _my_dispatch_ord = 10**9
                        if _task_blocked_by_global_dispatch_trial_order(
                            task,
                            task_queue,
                            current_date,
                            daily_status=daily_status,
                            members=members,
                            machine_avail_dt=machine_avail_dt,
                            machine_day_start=_machine_day_start,
                            machine_handoff=machine_handoff_legacy,
                            skills_dict=skills_dict,
                            abolish_all_scheduling_limits=bool(
                                global_priority_override.get(
                                    "abolish_all_scheduling_limits"
                                )
                            ),
                            dispatch_interval_mirror=_dispatch_interval_mirror,
                            min_dispatch_effective=_min_dispatch_eff_legacy,
                        ):
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _log_dispatch_trace_schedule(
                                    task.get("task_id"),
                                    "[配台トレース task=%s] スキップ: より尝さい配台試行順に未完了あり "
                                    "day=%s my_order=%s",
                                    task.get("task_id"),
                                    current_date,
                                    _my_dispatch_ord,
                                )
                            continue
                        if _task_fully_machine_calendar_blocked_on_date(
                            task, current_date, daily_status, members
                        ):
                            continue
                        if _task_no_machining_window_left_from_avail_floor(
                            task,
                            current_date,
                            daily_status,
                            members,
                            machine_avail_dt,
                            _machine_day_start,
                            machine_handoff=machine_handoff_legacy,
                            skills_dict=skills_dict,
                            abolish_all_scheduling_limits=bool(
                                global_priority_override.get(
                                    "abolish_all_scheduling_limits"
                                )
                            ),
                            dispatch_interval_mirror=_dispatch_interval_mirror,
                        ):
                            continue
                        if _equipment_line_lower_dispatch_trial_still_pending(
                            task_queue,
                            machine_occ_key,
                            _my_dispatch_ord,
                            current_date,
                            daily_status=daily_status,
                            members=members,
                            machine_avail_dt=machine_avail_dt,
                            machine_day_start=_machine_day_start,
                            machine_handoff=machine_handoff_legacy,
                            skills_dict=skills_dict,
                            abolish_all_scheduling_limits=bool(
                                global_priority_override.get(
                                    "abolish_all_scheduling_limits"
                                )
                            ),
                            dispatch_interval_mirror=_dispatch_interval_mirror,
                        ):
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _log_dispatch_trace_schedule(
                                    task.get("task_id"),
                                    "[配台トレース task=%s] スキップ: 同一設備で配台試行順は先の行は未完了 "
                                    "day=%s eq_line=%s my_order=%s",
                                    task.get("task_id"),
                                    current_date,
                                    eq_line,
                                    _my_dispatch_ord,
                                )
                            continue
                        machine_name = str(task.get("machine_name", "") or "").strip()
                        machine_proc = str(machine or "").strip()
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
                                need_src_line = (
                                    (need_src_line + "；") if need_src_line else ""
                                )
                                need_src_line += (
                                    f"計画シート必須人数{plan_ro}は未使用（need基準={req_num}）"
                                )
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
                        if global_priority_override.get("ignore_need_minimum"):
                            req_num = 1
                            need_src_line = (
                                (need_src_line + " → ")
                                if need_src_line
                                else ""
                            ) + "メイン上書ignore_need_minimumでreq=1"
    
                        # メンバー×設備スキル（parse_op_as_skill_cell: 尝さい優先度ろど先にフォーム候補へ採用）
                        # skills 読込時に「機械名」独立キーへエイリアスれるため、工程名+機械名は両方ある行では
                        # 複坈キー「工程名+機械名」のみを見る（別工程の坌坝機械の OP は浝れ込まないよごにれる）。
                        skill_meta_cache = {}
                        _gpo = global_priority_override
    
                        def skill_role_priority(mem):
                            if _gpo.get("ignore_skill_requirements"):
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
    
                        capable_members = [m for m in avail_dt.keys() if skill_role_priority(m)[0] in ("OP", "AS")]
                        capable_members.sort(key=lambda mm: (skill_role_priority(mm)[1], mm))
                        capable_members = _filter_capable_members_b2_disjoint_teams(
                            task, task_queue, capable_members
                        )
                        if task.get("has_done_deadline_override"):
                            machine_free_dbg = _machine_effective_floor_for_assign(
                                machine_occ_key,
                                str(task.get("task_id") or "").strip(),
                                eq_line,
                                str(task.get("machine_name") or "").strip(),
                                machine_avail_dt,
                                machine_handoff_legacy,
                                _machine_day_start,
                                bool(_gpo.get("abolish_all_scheduling_limits")),
                                current_date=current_date,
                                daily_status=daily_status,
                                skills_dict=skills_dict,
                                machine_proc=machine_proc,
                            )
                            logging.info(
                                "DEBUG[完了日指定] 依頼NO=%s 設備=%s req_num=%s capable_members=%s machine_free=%s",
                                task.get("task_id"),
                                eq_line,
                                req_num,
                                len(capable_members),
                                machine_free_dbg,
                            )
    
                        pref_raw = str(task.get("preferred_operator_raw") or "").strip()
                        op_today = [m for m in capable_members if skill_role_priority(m)[0] == "OP"]
                        pref_mem = (
                            _resolve_preferred_op_to_member(pref_raw, op_today, members)
                            if pref_raw
                            else None
                        )
                        if pref_raw and pref_mem is None and op_today:
                            logging.info(
                                "担当OP指定: 当日のOP候補に一致せう制約なし task=%s raw=%r",
                                task.get("task_id"),
                                pref_raw,
                            )

                        _gdp_must, _gdp_warns = _active_global_day_process_must_include(
                            _gpo,
                            task,
                            current_date,
                            capable_members,
                            members,
                        )
                        for _gw in _gdp_warns:
                            logging.warning(_gw)
                        fixed_team_anchor = _merge_global_day_process_and_pref_anchor(
                            _gdp_must, pref_mem, capable_members
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
                                need_src_line = (
                                    (need_src_line + " → ")
                                    if need_src_line
                                    else ""
                                )
                                need_src_line += (
                                    f"グローバル(日付×工程)指定で最低{_nfix}人"
                                )
                            req_num = max(req_num, _nfix)
    
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
                                (extra_src_line + " → ")
                                if extra_src_line
                                else ""
                            ) + "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROWで0"
                        extra_max = (
                            extra_max_sheet
                            if TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
                            else 0
                        )
                        if (
                            extra_max_sheet > 0
                            and not TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
                            and not TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW
                        ):
                            extra_src_line = (
                                (extra_src_line + " → ")
                                if extra_src_line
                                else ""
                            ) + "メインは基本人数のみ（余力枠は全配台後に未割当×スキルで追記）"
                        max_team_size = min(req_num + extra_max, len(capable_members))
                        if max_team_size < req_num:
                            max_team_size = req_num
                        rq_base = max(1, int(req_num))
    
                        _dto_head = task.get("dispatch_trial_order")
                        if (
                            _dto_head is not None
                            and _dto_head not in _need_headcount_logged_orders
                        ):
                            _need_headcount_logged_orders.add(_dto_head)
                            logging.info(
                                "need人数(配台試行順初回) order=%s task=%s 工程/機械=%s/%s "
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
    
                        trace_assign = bool(TRACE_TEAM_ASSIGN_TASK_ID) and (
                            str(task.get("task_id", "")).strip() == TRACE_TEAM_ASSIGN_TASK_ID
                        )
                        if trace_assign:
                            logging.info(
                                "TRACE配台[%s] %s 工程/機械=%s / %s req_num=%s extra_max=%s → max_team=%s "
                                "capable(n=%s)=%s ignore_need1=%s ignore_skill=%s abolish=%s 担当OP指定=%s→%s",
                                task["task_id"],
                                current_date,
                                _log_plain_label(machine),
                                _log_plain_label(machine_name),
                                req_num,
                                extra_max,
                                max_team_size,
                                len(capable_members),
                                capable_members,
                                global_priority_override.get("ignore_need_minimum"),
                                global_priority_override.get("ignore_skill_requirements"),
                                global_priority_override.get("abolish_all_scheduling_limits"),
                                pref_raw,
                                pref_mem,
                            )
    
                        team_candidates: list[dict] = []
                        combo_key = (
                            f"{machine_proc}+{machine_name}"
                            if machine_proc and machine_name
                            else ""
                        )
                        preset_rows = (
                            (team_combo_presets or {}).get(combo_key)
                            if (team_combo_presets and combo_key)
                            else None
                        )
                        if TEAM_ASSIGN_COMBO_SHEET_RESTRICT_TO_PRESET_MEMBERS and preset_rows:
                            _allowed_members = set()
                            for _p, _r, _pteam, _cid in preset_rows:
                                for _m in _pteam:
                                    _allowed_members.add(_m)
                            for _m in (fixed_team_anchor or []):
                                _allowed_members.add(_m)
                            if pref_mem:
                                _allowed_members.add(pref_mem)
                            _dropped = [
                                m for m in capable_members if m not in _allowed_members
                            ]
                            if _dropped:
                                logging.info(
                                    "組み合わせ表限定(legacy): %s/%s は表記載メンバーのみに配台探索を限定（除外=%s）。",
                                    _log_plain_label(machine),
                                    _log_plain_label(machine_name),
                                    ",".join(_log_plain_label(m) for m in _dropped),
                                )
                            capable_members = [
                                m for m in capable_members if m in _allowed_members
                            ]
                            op_today = [
                                m
                                for m in capable_members
                                if skill_role_priority(m)[0] == "OP"
                            ]
                            max_team_size = min(
                                req_num + extra_max, len(capable_members)
                            )
                            if max_team_size < req_num:
                                max_team_size = req_num
                        _prev_mach_raw_legacy = machine_avail_dt.get(
                            machine_occ_key, _machine_day_start
                        )
                        (
                            _mach_floor_legacy,
                            _co_segs_legacy,
                            _abort_legacy,
                        ) = _resolve_machine_changeover_floor_segments(
                            abolish_all_scheduling_limits=bool(
                                _gpo.get("abolish_all_scheduling_limits")
                            ),
                            machine_occ_key=machine_occ_key,
                            task_id=str(task.get("task_id") or "").strip(),
                            eq_line=eq_line,
                            machine_name=str(task.get("machine_name") or "").strip(),
                            machine_proc=machine_proc,
                            machine_avail_dt=machine_avail_dt,
                            machine_day_floor=_machine_day_start,
                            current_date=current_date,
                            machine_handoff=machine_handoff_legacy,
                            daily_status=daily_status,
                            skills_dict=skills_dict,
                            dispatch_interval_mirror=_dispatch_interval_mirror,
                            timeline_events=timeline_events,
                            task_queue=task_queue,
                            avail_dt=avail_dt,
                        )
                        if _abort_legacy:
                            continue
                        # プリセットは成立分をまとめて候補に載せ」下の組み合わせ探索とまとめで最良を決める。
                        if preset_rows:
                            for _prio, sheet_rs, preset_team, combo_row_id in preset_rows:
                                pteam = tuple(preset_team)
                                bounds = _combo_preset_team_size_bounds(
                                    pteam, sheet_rs, max_team_size
                                )
                                if bounds is None:
                                    continue
                                if fixed_team_anchor and not all(
                                    m in pteam for m in fixed_team_anchor
                                ):
                                    continue
                                if pref_mem is not None and pref_mem not in pteam:
                                    continue
                                if not all(m in capable_members for m in pteam):
                                    continue
                                _append_legacy_dispatch_candidate_for_team(
                                    task,
                                    pteam,
                                    avail_dt,
                                    machine_avail_dt,
                                    daily_status,
                                    current_date,
                                    macro_run_date,
                                    macro_now_dt,
                                    skill_role_priority,
                                    eq_line,
                                    rq_base,
                                    extra_max,
                                    global_priority_override,
                                    team_candidates,
                                    combo_sheet_row_id=combo_row_id,
                                    combo_preset_team=pteam,
                                    dispatch_interval_mirror=_dispatch_interval_mirror,
                                    machine_handoff=machine_handoff_legacy,
                                    machine_day_floor=_machine_day_start,
                                    machine_floor_cached=_mach_floor_legacy,
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
                                        for rest in itertools.combinations(
                                            others, need_extra
                                        )
                                    ]
                                else:
                                    teams_iter = []
                            elif (
                                pref_mem is not None
                                and pref_mem in capable_members
                                and skill_role_priority(pref_mem)[0] == "OP"
                            ):
                                others = [m for m in capable_members if m != pref_mem]
                                if tsize == 1:
                                    teams_iter = [(pref_mem,)]
                                elif len(others) >= tsize - 1:
                                    teams_iter = [
                                        tuple([pref_mem] + list(rest))
                                        for rest in itertools.combinations(others, tsize - 1)
                                    ]
                                else:
                                    logging.info(
                                        "担当OP指定: フォーム人数を満たせないため、指定を無視 task=%s size=%s raw=%r",
                                        task.get("task_id"),
                                        tsize,
                                        pref_raw,
                                    )
                                    teams_iter = itertools.combinations(capable_members, tsize)
                            else:
                                teams_iter = itertools.combinations(capable_members, tsize)
    
                            for team in teams_iter:
                                op_list = [m for m in team if skill_role_priority(m)[0] == "OP"]
                                if not op_list:
                                    continue
    
                                team_start = max(avail_dt[m] for m in team)
                                if not _gpo.get("abolish_all_scheduling_limits"):
                                    # 同一設備は1時点で1タスクのみ（設備空し＋日次始業/依頼切替の準備・後始末）
                                    machine_free_dt = _mach_floor_legacy
                                    if team_start < machine_free_dt:
                                        team_start = machine_free_dt
                                    # 原板投入日と同日の開始は 12:45 以降（試行順優先フローと一致）
                                    if task.get("same_day_raw_start_limit") and current_date == task["start_date_req"]:
                                        min_start_dt = datetime.combine(
                                            current_date, task["same_day_raw_start_limit"]
                                        )
                                        if team_start < min_start_dt:
                                            team_start = min_start_dt
                                    if current_date == task["start_date_req"] and task.get("earliest_start_time"):
                                        min_user_t = datetime.combine(
                                            current_date, task["earliest_start_time"]
                                        )
                                        if team_start < min_user_t:
                                            team_start = min_user_t
                                    # 当日は「マクロ実行した時刻」より剝に開始でしない
                                    if current_date == macro_run_date and team_start < macro_now_dt:
                                        team_start = macro_now_dt
                                team_end_limit = min(daily_status[m]['end_dt'] for m in team)
                                team_end_limit = _interactive_trial_relax_team_end_limit_to_eod(
                                    team_end_limit, current_date
                                )

                                if team_start >= team_end_limit:
                                    continue

                                team_breaks = []
                                for m in team:
                                    team_breaks.extend(daily_status[m]['breaks_dt'])
                                team_breaks = merge_time_intervals(team_breaks)

                                avg_eff = sum(daily_status[m]['efficiency'] for m in team) / len(team)
                                if avg_eff <= 0:
                                    avg_eff = 0.01
                                t_eff = parse_float_safe(task.get("task_eff_factor"), 1.0)
                                if t_eff <= 0:
                                    t_eff = 1.0
                                eff_time_per_unit = (
                                    task["base_time_per_unit"]
                                    / avg_eff
                                    / t_eff
                                    * _surplus_team_time_factor(rq_base, len(team), extra_max)
                                )
                                _defer_min_contig = max(1, int(math.ceil(float(eff_time_per_unit))))
                                _eod_cont_exempt_il = (
                                    _eod_same_request_continuation_exempt(
                                        machine_occ_key, task, machine_handoff_legacy
                                    )
                                )

                                def _refloor_legacy_inline(ts):
                                    ts = max(ts, max(avail_dt[m] for m in team))
                                    if not _gpo.get("abolish_all_scheduling_limits"):
                                        _mfd = _mach_floor_legacy
                                        if ts < _mfd:
                                            ts = _mfd
                                        if task.get(
                                            "same_day_raw_start_limit"
                                        ) and current_date == task["start_date_req"]:
                                            _msd = datetime.combine(
                                                current_date,
                                                task["same_day_raw_start_limit"],
                                            )
                                            if ts < _msd:
                                                ts = _msd
                                        if current_date == task[
                                            "start_date_req"
                                        ] and task.get("earliest_start_time"):
                                            _mut = datetime.combine(
                                                current_date,
                                                task["earliest_start_time"],
                                            )
                                            if ts < _mut:
                                                ts = _mut
                                        if (
                                            current_date == macro_run_date
                                            and ts < macro_now_dt
                                        ):
                                            ts = macro_now_dt
                                    return ts
    
                                _ts_adj = _defer_team_start_past_prebreak_and_end_of_day(
                                    task,
                                    tuple(team),
                                    team_start,
                                    team_end_limit,
                                    team_breaks,
                                    _refloor_legacy_inline,
                                    min_contiguous_work_mins=_defer_min_contig,
                                    eod_same_request_continuation_exempt=_eod_cont_exempt_il,
                                )
                                if _ts_adj is None:
                                    continue
                                team_start = _ts_adj
                                _roll_prep_inline: list[dict] = []
                                if not _gpo.get("abolish_all_scheduling_limits"):
                                    team_start, _roll_prep_inline = (
                                        _roll_prep_segments_for_assign(
                                            team_start=team_start,
                                            team_breaks=team_breaks,
                                            machine_handoff=machine_handoff_legacy,
                                            machine_occ_key=machine_occ_key,
                                            current_date=current_date,
                                            task_id=str(task.get("task_id") or "").strip(),
                                            machine_proc=machine_proc,
                                            machine_name=str(
                                                task.get("machine_name") or ""
                                            ).strip(),
                                            eq_line=eq_line,
                                            abolish_limits=False,
                                            prev_machining_end=_resolve_prev_machining_end_for_roll_prep(
                                                machine_handoff_legacy,
                                                machine_occ_key,
                                                (
                                                    machine_handoff_legacy.get(
                                                        "last_machining_dt"
                                                    )
                                                    or {}
                                                ).get(machine_occ_key),
                                                machine_avail_dt,
                                                _machine_day_start,
                                                _prev_mach_raw_legacy,
                                            ),
                                            prev_eq_line=str(
                                                (
                                                    machine_handoff_legacy.get(
                                                        "last_eq"
                                                    )
                                                    or {}
                                                ).get(machine_occ_key, "")
                                                or ""
                                            ).strip(),
                                            machine_avail_dt=machine_avail_dt,
                                            machine_day_floor=_machine_day_start,
                                        )
                                    )
                                    team_start = _refloor_legacy_inline(team_start)
                                if team_start >= team_end_limit:
                                    continue
                                if (
                                    _dispatch_interval_mirror is not None
                                    and _roll_prep_inline
                                ):
                                    _prep_mirror_block = False
                                    for _pseg in _roll_prep_inline:
                                        _pst = _pseg.get("start_dt")
                                        _ped = _pseg.get("end_dt")
                                        _pok = str(
                                            _pseg.get("machine_occupancy_key")
                                            or machine_occ_key
                                        ).strip()
                                        if (
                                            isinstance(_pst, datetime)
                                            and isinstance(_ped, datetime)
                                            and _pok
                                            and _dispatch_interval_mirror.would_block_equipment(
                                                _pok, _pst, _ped
                                            )
                                        ):
                                            _prep_mirror_block = True
                                            break
                                    if _prep_mirror_block:
                                        continue
    
                                _, avail_mins, _ = calculate_end_time(team_start, 9999, team_breaks, team_end_limit)
    
                                units_can_do = int(avail_mins / eff_time_per_unit)
                                if units_can_do == 0:
                                    continue
    
                                units_today = min(units_can_do, math.ceil(task['remaining_units']))
                                if _eod_reject_capacity_units_below_threshold(
                                    units_today,
                                    team_start,
                                    team_end_limit,
                                    eod_same_request_continuation_exempt=_eod_cont_exempt_il,
                                    remaining_units_ceil=math.ceil(
                                        float(task.get("remaining_units") or 0)
                                    ),
                                ):
                                    continue
                                work_mins_needed = int(units_today * eff_time_per_unit)
                                if (
                                    _contiguous_work_minutes_until_next_break_or_limit(
                                        team_start, team_breaks, team_end_limit
                                    )
                                    < work_mins_needed
                                ):
                                    continue
                                actual_end_dt, _, _ = calculate_end_time(team_start, work_mins_needed, team_breaks, team_end_limit)
    
                                team_prio_sum = sum(skill_role_priority(m)[1] for m in team)
                                if (
                                    _dispatch_interval_mirror is not None
                                    and _dispatch_interval_mirror.would_block_roll(
                                        machine_occ_key,
                                        team,
                                        team_start,
                                        actual_end_dt,
                                    )
                                ):
                                    continue
                                team_candidates.append(
                                    {
                                        "team": team,
                                        "team_start": team_start,
                                        "actual_end_dt": actual_end_dt,
                                        "units_today": units_today,
                                        "team_breaks": team_breaks,
                                        "avg_eff": avg_eff,
                                        "prio_sum": team_prio_sum,
                                        "op_list": op_list,
                                        "eff_time_per_unit": eff_time_per_unit,
                                        "combo_sheet_row_id": None,
                                        "combo_preset_team": None,
                                        "combo_preset_priority": None,
                                        "roll_prep_segments": _roll_prep_inline,
                                    }
                                )
    
                        best_team = None
                        best_info = {
                            "start_dt": datetime.max,
                            "units_today": 0,
                            "prio_sum": 10**9,
                        }
                        t_min = (
                            min(c["team_start"] for c in team_candidates)
                            if team_candidates
                            else None
                        )
    
                        def _team_cand_key(c):
                            _base = _team_assignment_sort_tuple(
                                c["team"],
                                c["team_start"],
                                c["units_today"],
                                c["prio_sum"],
                                t_min,
                            )
                            _combo_rank = c.get("combo_preset_priority")
                            try:
                                _combo_rank_i = (
                                    int(_combo_rank) if _combo_rank is not None else 10**9
                                )
                            except (TypeError, ValueError):
                                _combo_rank_i = 10**9
                            try:
                                if isinstance(_base, tuple) and len(_base) >= 1:
                                    return _base[:-1] + (_combo_rank_i, _base[-1])
                            except Exception:
                                pass
                            return (_combo_rank_i, _base)
    
                        if team_candidates:
                            best_c = min(team_candidates, key=_team_cand_key)
                            if best_c.get("combo_sheet_row_id") is None and preset_rows:
                                _lcid_l = _lookup_combo_sheet_row_id_for_preset_team(
                                    preset_rows, tuple(best_c["team"])
                                )
                                if _lcid_l is not None:
                                    best_c = {
                                        **best_c,
                                        "combo_sheet_row_id": _lcid_l,
                                        "combo_preset_team": tuple(best_c["team"]),
                                    }
                            if pref_mem and pref_mem in best_c["op_list"]:
                                lead_op = pref_mem
                            else:
                                lead_op = min(
                                    best_c["op_list"],
                                    key=lambda mm: (skill_role_priority(mm)[1], mm),
                                )
                            best_team = best_c["team"]
                            best_info = {
                                "start_dt": best_c["team_start"],
                                "end_dt": best_c["actual_end_dt"],
                                "op": lead_op,
                                "units_today": best_c["units_today"],
                                "breaks": best_c["team_breaks"],
                                "eff": best_c["avg_eff"],
                                "prio_sum": best_c["prio_sum"],
                            }
    
                        if trace_assign:
                            _tk = _team_assign_trace_tuple_label()
                            tid = task["task_id"]
                            for tsize in range(req_num, max_team_size + 1):
                                sub = [c for c in team_candidates if len(c["team"]) == tsize]
                                if not sub:
                                    logging.info(
                                        "TRACE配台[%s] %s tsize=%s → この人数で成立れるフォームなし",
                                        tid,
                                        current_date,
                                        tsize,
                                    )
                                else:
                                    sm = min(sub, key=_team_cand_key)
                                    logging.info(
                                        "TRACE配台[%s] %s tsize=%s 人数内最良: members=%s "
                                        "start=%s units_today=%s prio_sum=%s eff_t/unit=%.6f "
                                        "比較ルール=%s ※全日最早開始=%s を基準に辞書式で尝さい方は採用",
                                        tid,
                                        current_date,
                                        tsize,
                                        sm["team"],
                                        sm["team_start"],
                                        sm["units_today"],
                                        sm["prio_sum"],
                                        sm["eff_time_per_unit"],
                                        _tk,
                                        t_min.isoformat(sep=" ") if t_min else "—",
                                    )
    
                        if trace_assign and best_team is not None:
                            logging.info(
                                "TRACE配台[%s] %s ★採用 n=%s members=%s start=%s units_today=%s prio_sum=%s",
                                task["task_id"],
                                current_date,
                                len(best_team),
                                best_team,
                                best_info["start_dt"],
                                best_info["units_today"],
                                best_info["prio_sum"],
                            )
                            if len(best_team) == 1 and max_team_size > req_num:
                                if TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF:
                                    logging.info(
                                        "TRACE配台[%s] %s 1人採用（TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF）: "
                                        "より大しい人数で有効なフォームなし（OP丝足・0短縮・開始>=終了等）。",
                                        task["task_id"],
                                        current_date,
                                    )
                                else:
                                    logging.info(
                                        "TRACE配台[%s] %s 1人採用: 人数を増やれと開始は遅れ」"
                                        "スラック外では開始優先で1人は選べた可能性。"
                                        "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=%s」または従来の人数最優先は環境変数参照。",
                                        task["task_id"],
                                        current_date,
                                        TEAM_ASSIGN_START_SLACK_WAIT_MINUTES,
                                    )
    
                        if best_team:
                            if len(best_team) > req_num:
                                logging.info(
                                    "配台採用人数>req_num task=%s day=%s order=%s 工程/機械=%s/%s "
                                    "採用=%s人 req_num=%s extra_max=%s max_team=%s [%s] [%s]",
                                    task["task_id"],
                                    current_date,
                                    task.get("dispatch_trial_order"),
                                    _log_plain_label(machine),
                                    _log_plain_label(machine_name),
                                    len(best_team),
                                    req_num,
                                    extra_max,
                                    max_team_size,
                                    need_src_line,
                                    extra_src_line,
                                )
                            sub_members = [m for m in best_team if m != best_info["op"]]
                            done_units = best_info["units_today"]
                            if task.get("roll_pipeline_inspection") or task.get(
                                "roll_pipeline_rewind"
                            ):
                                _rp_room = _roll_pipeline_inspection_assign_room(
                                    task_queue, str(task.get("task_id", "")).strip()
                                )
                                done_units = min(
                                    int(done_units),
                                    int(min(_rp_room, math.ceil(task["remaining_units"]))),
                                )
                            else:
                                done_units = int(done_units)
                            if done_units <= 0:
                                if _trace_schedule_task_enabled(task.get("task_id")):
                                    _rp_log = None
                                    if task.get(
                                        "roll_pipeline_inspection"
                                    ) or task.get("roll_pipeline_rewind"):
                                        _rp_log = _roll_pipeline_inspection_assign_room(
                                            task_queue,
                                            str(task.get("task_id", "") or "").strip(),
                                        )
                                    _log_dispatch_trace_schedule(
                                        task.get("task_id"),
                                        "[配台トレース task=%s] スキップ: フォーム採用後の実効ユニット0 "
                                        "day=%s machine=%s best_units_today=%s rp_room=%s rem=%.4f",
                                        task.get("task_id"),
                                        current_date,
                                        machine,
                                        best_info.get("units_today"),
                                        _rp_log,
                                        float(task.get("remaining_units") or 0),
                                    )
                                continue
                            if done_units < best_info["units_today"]:
                                team_end_limit = min(
                                    daily_status[m]["end_dt"] for m in best_team
                                )
                                _teff = parse_float_safe(task.get("task_eff_factor"), 1.0)
                                if _teff <= 0:
                                    _teff = 1.0
                                _eff_t = (
                                    task["base_time_per_unit"]
                                    / best_info["eff"]
                                    / _teff
                                    * _surplus_team_time_factor(rq_base, len(best_team), extra_max)
                                )
                                _wm = int(done_units * _eff_t)
                                _end_dt, _, _ = calculate_end_time(
                                    best_info["start_dt"],
                                    _wm,
                                    best_info["breaks"],
                                    team_end_limit,
                                )
                                best_info = dict(best_info)
                                best_info["end_dt"] = _end_dt
                                best_info["units_today"] = done_units
    
                            total_u = math.ceil(task['total_qty_m'] / task['unit_m']) if task['unit_m'] else 0
                            rem_u_before = math.ceil(task['remaining_units'])
                            already_done = total_u - rem_u_before
                            
                            # 「マクロ実行時点」の完了率（予定の進杗ではなく」実加工数ベース）
                            try:
                                tot_qty = parse_float_safe(task.get('total_qty_m'), 0.0)
                                done_qty = parse_float_safe(task.get('done_qty_reported'), 0.0)
                                if tot_qty > 0:
                                    pct_macro = max(0, min(100, int(round((done_qty / tot_qty) * 100))))
                                else:
                                    pct_macro = 0
                            except Exception:
                                pct_macro = 0
                            
                            _te_disp = parse_float_safe(task.get("task_eff_factor"), 1.0)
                            if _te_disp <= 0:
                                _te_disp = 1.0
                            _legacy_mach_sub = ", ".join(
                                str(s).strip()
                                for s in sub_members
                                if s and str(s).strip()
                            )
                            _co_append_l = list(_co_segs_legacy or []) + list(
                                best_info.get("roll_prep_segments") or []
                            )
                            _append_changeover_segments_to_timeline(
                                timeline_events,
                                _dispatch_interval_mirror,
                                avail_dt,
                                daily_status,
                                current_date=current_date,
                                task_id=str(task.get("task_id") or ""),
                                machine_occ_key=machine_occ_key,
                                segments=_co_append_l,
                                machining_lead_op=str(
                                    best_info.get("op") or ""
                                ).strip()
                                or None,
                                machining_sub_str=_legacy_mach_sub or None,
                                machine_handoff=machine_handoff_legacy,
                                skill_role_priority=skill_role_priority,
                                machine_name_for_startup=str(machine_name or "").strip()
                                or None,
                            )
                            _legacy_mach_ev = {
                                "date": current_date, "task_id": task['task_id'], "machine": eq_line,
                                "machine_occupancy_key": machine_occ_key,
                                "op": best_info["op"], "sub": ", ".join(sub_members),
                                "start_dt": best_info["start_dt"], "end_dt": best_info["end_dt"],
                                "breaks": best_info["breaks"], "units_done": done_units,
                                "already_done_units": already_done,
                                "total_units": total_u,
                                "pct_macro": pct_macro,
                                "eff_time_per_unit": task["base_time_per_unit"]
                                / best_info["eff"]
                                / _te_disp
                                * _surplus_team_time_factor(rq_base, len(best_team), extra_max),
                                "unit_m": task['unit_m'],
                                "total_qty_m": float(parse_float_safe(task.get("total_qty_m"), 0.0)),
                                "event_kind": TIMELINE_EVENT_MACHINING,
                            }
                            timeline_events.append(_legacy_mach_ev)
                            _stage2_dispatch_track_timeline_event(
                                _legacy_mach_ev, len(timeline_events)
                            )
                            if _dispatch_interval_mirror is not None:
                                _dispatch_interval_mirror.register_from_event(
                                    timeline_events[-1]
                                )
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _rp_tr = None
                                if task.get("roll_pipeline_inspection") or task.get(
                                    "roll_pipeline_rewind"
                                ):
                                    _rp_tr = _roll_pipeline_inspection_assign_room(
                                        task_queue,
                                        str(task.get("task_id", "") or "").strip(),
                                    )
                                _log_dispatch_trace_schedule(
                                    task.get("task_id"),
                                    "[配台トレース task=%s] タイムライン追記 chunk day=%s machine=%s "
                                    "done_units=%s already_done=%s total_u=%s rem_after=%.4f "
                                    "start=%s end=%s eff_t/unit=%.4f rp_room(当時)=%s",
                                    task.get("task_id"),
                                    current_date,
                                    eq_line,
                                    done_units,
                                    already_done,
                                    total_u,
                                    float(task.get("remaining_units") or 0)
                                    - float(done_units),
                                    best_info["start_dt"],
                                    best_info["end_dt"],
                                    float(
                                        task["base_time_per_unit"]
                                        / best_info["eff"]
                                        / _te_disp
                                        * _surplus_team_time_factor(
                                            rq_base, len(best_team), extra_max
                                        )
                                    ),
                                    _rp_tr,
                                )

                            task["remaining_units"] = max(
                                0.0,
                                float(task.get("remaining_units") or 0)
                                - float(done_units),
                            )
                            op_main = (best_info.get("op") or "").strip()
                            subs_part = ",".join(
                                s.strip() for s in sub_members if s and str(s).strip()
                            )
                            team_s = f"{op_main}, {subs_part}" if subs_part else op_main
                            need_surplus_assigned = (
                                TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
                                and extra_max > 0
                                and len(best_team) > req_num
                            )
                            _lo = (best_info.get("op") or "").strip()
                            _subs_legacy = [
                                str(s).strip()
                                for s in sub_members
                                if s and str(s).strip()
                            ]
                            _names_ord = ([] if not _lo else [_lo]) + _subs_legacy
                            _surplus_names = (
                                _names_ord[int(req_num) :]
                                if need_surplus_assigned
                                and len(_names_ord) > int(req_num)
                                else []
                            )
                            task["assigned_history"].append(
                                {
                                    "date": current_date.strftime("%m/%d"),
                                    "team": team_s,
                                    "done_m": int(done_units * task["unit_m"]),
                                    "start_dt": best_info["start_dt"],
                                    "end_dt": best_info["end_dt"],
                                    "need_surplus_assigned": need_surplus_assigned,
                                    "combo_sheet_row_id": best_c.get(
                                        "combo_sheet_row_id"
                                    ),
                                    "surplus_member_names": _surplus_names,
                                }
                            )

                            for m in best_team:
                                avail_dt[m] = best_info["end_dt"]
                            if not _gpo.get("abolish_all_scheduling_limits"):
                                machine_avail_dt[machine_occ_key] = best_info["end_dt"]
                                _bump_machine_avail_after_roll_for_calendar(
                                    current_date,
                                    machine_occ_key,
                                    machine_avail_dt,
                                    machine_calendar_plan_end=_machine_calendar_plan_end,
                                    machine_day_floor=_machine_day_start,
                                )
                            machine_handoff_legacy["last_tid"][machine_occ_key] = str(
                                task.get("task_id") or ""
                            ).strip()
                            machine_handoff_legacy["last_eq"][machine_occ_key] = eq_line
                            machine_handoff_legacy["started_today"].add(
                                machine_occ_key
                            )
                            machine_handoff_legacy["machining_today_occ"].add(
                                machine_occ_key
                            )
                            machine_handoff_legacy["last_machining_dt"][
                                machine_occ_key
                            ] = best_info["end_dt"]
                            machine_handoff_legacy["last_machining_date"][
                                machine_occ_key
                            ] = current_date
                            machine_handoff_legacy["last_lead_op"][
                                machine_occ_key
                            ] = best_info["op"]
                            machine_handoff_legacy.setdefault("last_machining_sub", {})
                            machine_handoff_legacy["last_machining_sub"][
                                machine_occ_key
                            ] = _legacy_mach_sub
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _log_dispatch_trace_schedule(
                                    task.get("task_id"),
                                    "[配台トレース task=%s] ロール確定 メイン day=%s machine=%s machine_name=%s "
                                    "start=%s end=%s 採用人数=%s req_num=%s メイン探索extra_max=%s "
                                    "余剰人数適用(メイン)=%s team=%s",
                                    task.get("task_id"),
                                    current_date,
                                    eq_line,
                                    str(machine_name or "").strip(),
                                    best_info["start_dt"],
                                    best_info["end_dt"],
                                    len(best_team),
                                    int(req_num),
                                    int(extra_max),
                                    need_surplus_assigned,
                                    team_s,
                                )
                            _sched_made_progress = True
                        else:
                            if task.get("has_done_deadline_override"):
                                logging.info(
                                    "DEBUG[完了日指定] 依頼NO=%s 日付=%s は割当試行（覝員/設備空し条件でフォーム未成立）。remaining_units=%s",
                                    task.get("task_id"),
                                    current_date,
                                    task.get("remaining_units"),
                                )
                if not STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST:
                    _legacy_pass_sec = time_module.perf_counter() - _t_legacy0
                    _dispatch_loop_profile_add("legacy_pass_wall", _legacy_pass_sec)
                    if _STAGE2_DISPATCH_LOOP_PROFILE_DAY is not None:
                        _STAGE2_DISPATCH_LOOP_PROFILE_DAY.setdefault(
                            "sched_pass_secs", []
                        ).append(_legacy_pass_sec)

                if not _sched_made_progress:
                    break

            _day_elapsed = time_module.perf_counter() - _t_day0
            _dispatch_day_timing.append(
                {
                    "date": current_date.isoformat(),
                    "elapsed_sec": round(_day_elapsed, 3),
                    "sched_passes": _day_sched_passes,
                    "sched_max_passes": _sched_max_passes,
                    "tasks_today": len(tasks_today),
                    "profile": _dispatch_loop_profile_finish_day(
                        _day_elapsed,
                        sched_passes=_day_sched_passes,
                        sched_max_passes=_sched_max_passes,
                    ),
                }
            )

            if TRACE_SCHEDULE_TASK_IDS:
                for _tt in TRACE_SCHEDULE_TASK_IDS:
                    for _t in task_queue:
                        if str(_t.get("task_id", "")).strip() != _tt:
                            continue
                        _rem_tr = float(_t.get("remaining_units") or 0)
                        if _rem_tr <= 1e-9:
                            continue
                        _log_dispatch_trace_schedule(
                            _tt,
                            "[配台トレース task=%s] 日次終了時点の残 day=%s machine=%s "
                            "machine_name=%s rem=%.4f roll_b2_follower=%s 試行順=%s",
                            _tt,
                            current_date,
                            _t.get("machine"),
                            _t.get("machine_name"),
                            _rem_tr,
                            bool(
                                _t.get("roll_pipeline_inspection")
                                or _t.get("roll_pipeline_rewind")
                            ),
                            _t.get("dispatch_trial_order"),
                        )

            if STAGE2_RETRY_SHIFT_DUE_ON_PARTIAL_REMAINING:
                missed_tids = _collect_task_ids_missed_deadline_after_day(
                    task_queue, current_date
                )
                if missed_tids:
                    blocked_tids = set()
                    shift_tid_list = []
                    for _ptid in sorted(missed_tids):
                        _do_shift, _cal_short = _partial_task_id_due_shift_outcome(
                            task_queue, _ptid, _calendar_last_plan_day
                        )
                        if _cal_short:
                            blocked_tids.add(_ptid)
                        if _do_shift:
                            shift_tid_list.append(_ptid)
                    for t in task_queue:
                        _tid = str(t.get("task_id", "") or "").strip()
                        if _tid in blocked_tids:
                            t["_partial_retry_calendar_blocked"] = True
                    if shift_tid_list:
                        allowed_shift_tids = [
                            tid
                            for tid in shift_tid_list
                            if _due_shift_retry_count_by_request.get(tid, 0)
                            < STAGE2_RETRY_SHIFT_DUE_MAX_ROUNDS
                        ]
                        for tid in shift_tid_list:
                            if tid not in allowed_shift_tids:
                                _due_shift_exhausted_requests.add(tid)
                        if allowed_shift_tids:
                            _outer_retry_round += 1
                            for tid in allowed_shift_tids:
                                _due_shift_retry_count_by_request[tid] = (
                                    _due_shift_retry_count_by_request.get(tid, 0) + 1
                                )
                            shift_set = set(allowed_shift_tids)
                            for t in task_queue:
                                if str(t.get("task_id", "") or "").strip() in shift_set:
                                    _shift_task_due_calendar_fields_one_day(t, run_date)
                            timeline_events[:] = [
                                e
                                for e in timeline_events
                                if _normalize_timeline_task_id(e) not in shift_set
                            ]
                            if _dispatch_interval_mirror is not None:
                                _dispatch_interval_mirror.rebuild_from_timeline(
                                    timeline_events
                                )
                            if (
                                _interactive_dispatch_trial_env_active()
                                and interactive_dispatch_targets
                            ):
                                _interactive_trial_meters_done.clear()
                                _interactive_trial_meters_done.update(
                                    _interactive_trial_recompute_meters_done_from_timeline(
                                        timeline_events,
                                        task_queue,
                                        interactive_dispatch_targets,
                                    )
                                )
                            for t in task_queue:
                                if str(t.get("task_id", "") or "").strip() in shift_set:
                                    t["remaining_units"] = float(
                                        t.get("initial_remaining_units") or 0
                                    )
                                    t["assigned_history"].clear()
                            _apply_dispatch_trial_order_for_generate_plan(
                                task_queue,
                                req_map,
                                need_rules,
                                need_combo_col_index,
                            )
                            _trials_detail = ",".join(
                                f"{tid}:{_due_shift_retry_count_by_request[tid]}"
                                for tid in sorted(allowed_shift_tids)
                            )
                            logging.info(
                                "紝期超靎リトライ: 計画基準+1日して当該依頼のみ再配台（検出日=%s 依頼NO=%s 当該依頼の累計試行=%s）",
                                current_date.isoformat(),
                                ",".join(sorted(allowed_shift_tids)),
                                _trials_detail,
                            )
                            _full_calendar_without_deadline_restart = False
                            break
                        else:
                            # 依頼とと上限でシフトでしないの値のときは日付ループを継続れる（break れると未処理日は残り配台試行は大量発生れる）。
                            _cap_tids = sorted(
                                tid
                                for tid in shift_tid_list
                                if tid not in allowed_shift_tids
                            )
                            _first_cap_warn = [
                                tid for tid in _cap_tids if tid not in _due_shift_cap_warned_tids
                            ]
                            for tid in _first_cap_warn:
                                _due_shift_cap_warned_tids.add(tid)
                            if _first_cap_warn:
                                logging.warning(
                                    "紝期後ゝ倒し再配台: 次の依頼NOは依頼ごとの上限（坄 %s 回）のため、この検出では +1 しません。"
                                    " カレンダーは継続しした（未完了は終了時に紝期見直し必須を付け得した）: %s",
                                    STAGE2_RETRY_SHIFT_DUE_MAX_ROUNDS,
                                    ",".join(_cap_tids),
                                )

        if _full_calendar_without_deadline_restart:
            _t_b2rew0 = time_module.perf_counter()
            _rewind_made = _run_b2_inspection_rewind_pass(
                sorted_dates,
                attendance_data,
                task_queue,
                timeline_events,
                skills_dict,
                members,
                equipment_list,
                req_map,
                need_rules,
                surplus_map,
                global_priority_override,
                macro_run_date,
                macro_now_dt,
                _need_headcount_logged_orders,
                team_combo_presets,
                dispatch_interval_mirror=_dispatch_interval_mirror,
                interactive_dispatch_targets=interactive_dispatch_targets,
                interactive_trial_pair_dates=_interactive_trial_pair_dates,
                interactive_trial_meters_done=_interactive_trial_meters_done,
            )
            _dispatch_loop_profile_add(
                "b2_inspection_rewind_pass",
                time_module.perf_counter() - _t_b2rew0,
            )
            if _rewind_made:
                logging.info(
                    "§B-2/§B-3 リワインド: EC 完走後に検査＝巻返しのみ日付先頭から再配台しました（timeline_events を占有テーブルとして利用）。"
                )
            break

    _ctx_after_dispatch = (
        "段階3配台試行"
        if _interactive_dispatch_trial_env_active()
        else "段階2"
    )
    _raise_if_remaining_tasks_exceed_attendance_calendar(
        task_queue,
        _calendar_last_plan_day,
        context_label=_ctx_after_dispatch,
    )

    if interactive_dispatch_targets is not None:
        _meta = dict(_LAST_INTERACTIVE_STAGE3_META or {})
        _meta.setdefault("mode", "single_phase")
        _LAST_INTERACTIVE_STAGE3_META = _meta
    else:
        _LAST_INTERACTIVE_STAGE3_META = {}

    if TRACE_SCHEDULE_TASK_IDS:
        for _tt in TRACE_SCHEDULE_TASK_IDS:
            for _t in task_queue:
                if str(_t.get("task_id", "")).strip() != _tt:
                    continue
                _log_dispatch_trace_schedule(
                    _tt,
                    "[配台トレース task=%s] シミュレーション終了時 machine=%s machine_name=%s "
                    "rem=%.4f initial=%.4f roll_b2_follower=%s",
                    _tt,
                    _t.get("machine"),
                    _t.get("machine_name"),
                    float(_t.get("remaining_units") or 0),
                    float(_t.get("initial_remaining_units") or 0),
                    bool(
                        _t.get("roll_pipeline_inspection")
                        or _t.get("roll_pipeline_rewind")
                    ),
                )
            _evs_tr = sorted(
                (
                    e
                    for e in timeline_events
                    if str(e.get("task_id", "")).strip() == _tt
                ),
                key=lambda e: (e.get("date"), e.get("start_dt") or datetime.min),
            )
            _last_ev_by_machine: dict = {}
            for _e in _evs_tr:
                _last_ev_by_machine[str(_e.get("machine") or "")] = _e
            for _mk, _ev in sorted(_last_ev_by_machine.items()):
                _ad = int(_ev.get("already_done_units") or 0)
                _ud = int(_ev.get("units_done") or 0)
                _log_dispatch_trace_schedule(
                    _tt,
                    "[配台トレース task=%s] タイムライン最終塊(工程列とと) machine=%s "
                    "already_done+units_done=%s+%s=%s total_units=%s end_dt=%s",
                    _tt,
                    _mk,
                    _ad,
                    _ud,
                    _ad + _ud,
                    _ev.get("total_units"),
                    _ev.get("end_dt"),
                )

    # need「配台時追加人数」: メイン割付後に、未参加×スキル適合者をサブへ追記（既定）
    if (
        not TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
        and not TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW
        and surplus_map
        and timeline_events
    ):
        _n_sur = append_surplus_staff_after_main_dispatch(
            timeline_events,
            attendance_data,
            skills_dict,
            members,
            task_queue,
            req_map,
            need_rules,
            surplus_map,
            global_priority_override,
        )
        if _n_sur:
            logging.info(
                "need余力: メイン割付完了後にサブ %s 坝を追記（未割当×スキル・時間重なりなし）",
                _n_sur,
            )

    _n_snap = _repair_timeline_daily_startup_snapped_to_first_machining(
        timeline_events,
        task_queue,
        attendance_data,
        skills_dict,
        global_priority_override,
    )
    if _n_snap:
        logging.info(
            "日次始業準備: 当日先頭の加工から離れていた帯を %s 件、先頭加工または直前休憩の直前に寄せました。",
            _n_snap,
        )

    if _dispatch_interval_mirror is not None:
        _dispatch_interval_mirror.rebuild_from_timeline(timeline_events)

    _dispatch_loop_profile_emit_run_summary()

    _t_stage2_perf = _log_stage2_phase_timing(
        "dispatch_daily_loop",
        _t_dispatch0,
        extra={
            "timeline_events": len(timeline_events),
            "task_queue_rows": len(task_queue),
            "sched_pass_total": _dispatch_sched_pass_total,
            "mh_handoff_cache_hit": _MH_HANDOFF_TIMELINE_CACHE_STATS.get("hit", 0),
            "mh_handoff_cache_miss": _MH_HANDOFF_TIMELINE_CACHE_STATS.get("miss", 0),
            "mh_handoff_cache_incremental": _MH_HANDOFF_TIMELINE_CACHE_STATS.get(
                "incremental", 0
            ),
            "dispatch_days": len(_dispatch_day_timing),
            "dispatch_day_top3": sorted(
                _dispatch_day_timing,
                key=lambda d: d.get("elapsed_sec") or 0,
                reverse=True,
            )[:3],
            "dispatch_loop_buckets_top12": _dispatch_loop_profile_top_buckets(
                _STAGE2_DISPATCH_LOOP_PROFILE_RUN.get("run_buckets") or {}, 12
            ),
        },
    )

    # タイムラインを日付別にインデックス化し、サブメンバー一覧を事剝解析（以降の出力ループを高速化）
    for e in timeline_events:
        e["subs_list"] = [s.strip() for s in e["sub"].split(",")] if e.get("sub") else []

    events_by_date = defaultdict(list)
    for e in timeline_events:
        events_by_date[e["date"]].append(e)

    # =========================================================
    # 4. Excel出力 (メイン計画)
    # =========================================================
    _stage2_out_root = stage2_output_root if stage2_output_root else output_dir
    if stage2_output_root:
        try:
            os.makedirs(stage2_output_root, exist_ok=True)
        except OSError as e:
            logging.error("段階2: 出力先ディレクトリを作成できません: %s (%s)", stage2_output_root, e)
            return
    _t_rm_prior0 = time_module.perf_counter()
    if not skip_remove_prior_stage2_workbooks:
        _remove_prior_stage2_workbooks_and_prune_empty_dirs(_stage2_out_root)
    _t_stage2_perf = _log_stage2_phase_timing(
        "remove_prior_stage2_workbooks",
        _t_rm_prior0,
        extra={"output_root": _stage2_path_debug_meta(_stage2_out_root)},
    )
    # ファイル名は短い日本語接頭辞＋時刻（本体ベース名は 20 文字以内）。同一抽出の再実行でも
    # パスがぶつからないよう実行時刻のマイクロ秒下位を含める（Excel 占有で旧ファイル削除失敗時の上書き不能を回避）。
    _stage2_run_now = datetime.now()
    _stage2_stamp = format_stage2_stamp(base_now_dt, _stage2_run_now)
    plan_xlsx_final = os.path.join(_stage2_out_root, plan_workbook_filename(_stage2_stamp))
    _t_stage2_perf = _log_stage2_phase_timing(
        "stage2_output_paths_resolved",
        _t_stage2_perf,
        extra={
            "output_root": _stage2_path_debug_meta(_stage2_out_root),
            "plan_xlsx_final": _stage2_path_debug_meta(plan_xlsx_final),
            "pm_ai_output_dir": _stage2_path_debug_meta(
                (os.environ.get("PM_AI_OUTPUT_DIR") or "").strip()
            ),
            "processing_plan_path": _stage2_path_debug_meta(
                (os.environ.get("PM_AI_PROCESSING_PLAN_PATH") or "").strip()
            ),
            "plan_input_path": _stage2_path_debug_meta(
                (os.environ.get(ENV_PLAN_INPUT_PATH) or "").strip()
            ),
            "publish_plan_xlsx": bool(_publish_plan_xlsx := _stage2_publish_excel_enabled(stage2_output_root)),
        },
    )
    if _publish_plan_xlsx:
        output_filename = plan_xlsx_final
    else:
        import tempfile

        _fd_plan_tmp, output_filename = tempfile.mkstemp(
            suffix=".xlsx", prefix="_pm_stage2_plan_", dir=_stage2_out_root
        )
        os.close(_fd_plan_tmp)
    # タスクID → 結果_設備毎の時間割で当該タスクは最初に睾れるセル（例 B12）。結果_タスク一覧のリンク用。
    _t_build0 = time_module.perf_counter()
    first_eq_schedule_cell_by_task_id: dict[str, str] = {}
    df_eq_schedule = _build_equipment_schedule_dataframe(
        sorted_dates,
        equipment_list,
        attendance_data,
        timeline_events,
        first_eq_schedule_cell_by_task_id=first_eq_schedule_cell_by_task_id,
    )
    df_equipment_by_machine_name = _build_equipment_schedule_by_machine_name_dataframe(
        sorted_dates,
        equipment_list,
        attendance_data,
        timeline_events,
    )

    # 結果_タスク一覧用: シミュレーション上の当該タスクの最早開始・最靅終了（timeline_events 集約）
    plan_window_by_task_id: dict = {}
    for _ev in timeline_events:
        tid = _ev.get("task_id")
        if tid is None:
            continue
        sd = _ev.get("start_dt")
        ed = _ev.get("end_dt")
        if sd is None or ed is None:
            continue
        if tid not in plan_window_by_task_id:
            plan_window_by_task_id[tid] = [sd, ed]
        else:
            w = plan_window_by_task_id[tid]
            if sd < w[0]:
                w[0] = sd
            if ed > w[1]:
                w[1] = ed

    # 結果_タスク一覧の工程行ごとの窓（timeline の machine＝eq_line 単位）。依頼NOのみ集約した plan_window_by_task_id とは別。
    plan_window_by_task_line: dict[tuple[str, str], list] = {}
    for _ev in timeline_events:
        if not _is_machining_timeline_event(_ev):
            continue
        _tid_ev = str(_ev.get("task_id") or "").strip()
        if not _tid_ev:
            continue
        _mch_ev = str(_ev.get("machine") or "").strip()
        if not _mch_ev:
            continue
        sd = _ev.get("start_dt")
        ed = _ev.get("end_dt")
        if sd is None or ed is None:
            continue
        _k2 = (_tid_ev, _mch_ev)
        if _k2 not in plan_window_by_task_line:
            plan_window_by_task_line[_k2] = [sd, ed]
        else:
            w2 = plan_window_by_task_line[_k2]
            if sd < w2[0]:
                w2[0] = sd
            if ed > w2[1]:
                w2[1] = ed

    # 結果_タスク一覧の「回答納期」「指定納期」は配台計画_タスク入力の当該行セルのみ。
    # 「原反投入日」は上書き列に日付があるときはその値、ないときは列「原反投入日」（計画基準納期と混同しない）
    _result_sheet_answer_spec_by_line = {}
    _result_sheet_raw_input_by_line = _build_result_sheet_effective_raw_input_date_by_line(
        tasks_df
    )
    if tasks_df is not None and not getattr(tasks_df, "empty", True):
        for _, _r in tasks_df.iterrows():
            if _plan_row_exclude_from_assignment(_r):
                continue
            _tid = str(_planning_df_cell_scalar(_r, TASK_COL_TASK_ID) or "").strip()
            _mach = str(_planning_df_cell_scalar(_r, TASK_COL_MACHINE) or "").strip()
            if not _tid or not _mach:
                continue
            _ad = parse_optional_date(_planning_df_cell_scalar(_r, TASK_COL_ANSWER_DUE))
            _sd = parse_optional_date(_planning_df_cell_scalar(_r, TASK_COL_SPECIFIED_DUE))
            _result_sheet_answer_spec_by_line[(_tid, _mach)] = (_ad, _sd)

    def _result_task_dispatch_status_memo(status: str, rem_u: float, hist: bool) -> str:
        """結果_タスク一覧「配台状況メモ」列用。配台不可系ステータスのときのみ短文で理由を返す。"""
        st = str(status or "")
        if not st.startswith("配台不可"):
            return ""
        parts: list[str] = []
        if not hist:
            parts.append("計画終了時点で割当履歴なし（設備タイムラインに載っていません）")
        try:
            if float(rem_u) > 1e-9:
                parts.append(f"残ユニット約{float(rem_u):g}")
        except Exception:
            if rem_u > 1e-9:
                parts.append("残あり")
        if "納期見直し必須" in st:
            parts.append("依頼単位で計画基準納期の再試行上限に達した可能性")
        return "／".join(parts) if parts else "割当なし"

    task_results = []
    # ステータス（配台の状態・残）：完了相当=配台済。
    # 段階2標準・段階3同一: 未割当=配台残(計画期間内未割当)、一部=配台残（配台不可は master 前提失敗時のみ・ループ前）。
    # 計画基準+1 の再試行は依頼NOごとの上限に達した依頼の未完了行には（納期見直し必須）を付与する。
    sorted_tasks_for_result = sorted(task_queue, key=_result_task_sheet_sort_key)
    _interactive_hist_override: dict[tuple[str, str], list[dict]] = {}
    if _interactive_dispatch_trial_env_active() and isinstance(interactive_result_dispatch_json_rows, list):
        _tmp_hist: dict[tuple[str, str], list[dict]] = defaultdict(list)
        for _r in interactive_result_dispatch_json_rows:
            if not isinstance(_r, dict):
                continue
            _tid = _interactive_norm_cell(_r.get(TASK_COL_TASK_ID))
            _mach = _interactive_norm_cell(_r.get(TASK_COL_MACHINE_NAME))
            _dd = _interactive_parse_dispatch_date_cell(_r.get("配台日"))
            _qc = _r.get("当日配台数量")
            try:
                _qv = (
                    float(str(_qc).replace(",", "").strip())
                    if _qc not in (None, "")
                    else 0.0
                )
            except (TypeError, ValueError):
                _qv = 0.0
            if not _tid or not _mach or _dd is None or _qv <= 1e-9:
                continue
            _tmp_hist[(_tid, _mach)].append(
                {
                    "date": _dd.strftime("%m/%d"),
                    "done_m": float(_qv),
                }
            )
        for _k, _vv in _tmp_hist.items():
            _vv.sort(key=lambda _h: str(_h.get("date") or ""))
            _interactive_hist_override[_k] = _vv
    max_history_len = max(
        [
            len(merge_assigned_history_contiguous_for_result_sheet(t.get("assigned_history")))
            for t in sorted_tasks_for_result
        ]
        + [0]
    )
    if _interactive_hist_override:
        try:
            _ov_max = max((len(v) for v in _interactive_hist_override.values()), default=0)
            max_history_len = max(int(max_history_len), int(_ov_max))
        except Exception:
            pass
    _baseline_raw_by_line = _build_result_sheet_effective_raw_input_date_by_line(
        tasks_df_raw_input_baseline
    )
    _pat_label = (result_pattern_shift_label or "").strip()
    for t in sorted_tasks_for_result:
        rem_u = float(t.get("remaining_units") or 0)
        hist = bool(t.get("assigned_history"))
        # 負の残は「配台済」に含めない（-0.5R 等は配台残）。浮動小数の負残は反長に比例したメートル幅で配台済に含める。
        _um0 = float(t.get("unit_m") or 0)
        _rem_abs_m = abs(rem_u * _um0)
        _noise_tol_m = max(3.0, min(100.0, 0.025 * abs(_um0)))
        if rem_u <= 1e-9 and (rem_u >= 0 or _rem_abs_m <= _noise_tol_m):
            status = "配台済"
        elif hist and t.get("_partial_retry_calendar_blocked"):
            if _dispatch_postpone_only_policy_active():
                status = "配台残(勤怠カレンダー不足)"
            else:
                status = "配台残(勤務カレンダー不足)"
        elif not hist and rem_u > 1e-9:
            if _dispatch_postpone_only_policy_active():
                status = "配台残(計画期間内未割当)"
            else:
                status = "配台不可"
        else:
            status = "配台残"
        _tid_res = str(t.get("task_id", "") or "").strip()
        if (
            _tid_res in _due_shift_exhausted_requests
            and rem_u > 1e-9
            and "納期見直し必須" not in status
        ):
            status = f"{status}（納期見直し必須）"

        total_r = int(t['total_qty_m'] / t['unit_m']) if t['unit_m'] else 0

        _line_key = (str(t.get("task_id", "") or "").strip(), str(t.get("machine", "") or "").strip())
        _sheet_pair = _result_sheet_answer_spec_by_line.get(_line_key)
        if _sheet_pair is not None:
            _ans_d, _spec_d = _sheet_pair
            ans_s = _ans_d.strftime("%Y/%m/%d") if _ans_d else ""
            spec_s = _spec_d.strftime("%Y/%m/%d") if _spec_d else ""
        else:
            _ans_d = t.get("answer_due_date")
            _spec_d = t.get("specified_due_date")
            ans_s = _ans_d.strftime("%Y/%m/%d") if _ans_d else ""
            spec_s = _spec_d.strftime("%Y/%m/%d") if _spec_d else ""
        _basis_for_sheet = t.get("due_basis_date_result_sheet")
        if _basis_for_sheet is None:
            _basis_for_sheet = t.get("due_basis_date")
        basis_s = (
            _basis_for_sheet.strftime("%Y/%m/%d")
            if _basis_for_sheet is not None and hasattr(_basis_for_sheet, "strftime")
            else ""
        )
        if _line_key in _result_sheet_raw_input_by_line:
            _rid_d = _result_sheet_raw_input_by_line[_line_key]
            kenhan_s = _rid_d.strftime("%Y/%m/%d") if _rid_d else ""
        else:
            _rid_t = t.get("raw_input_date")
            kenhan_s = (
                _rid_t.strftime("%Y/%m/%d")
                if _rid_t is not None and hasattr(_rid_t, "strftime")
                else ""
            )
        _pre_w = _baseline_raw_by_line.get(_line_key) if _baseline_raw_by_line else None
        pre_pattern_kenhan_s = (
            _pre_w.strftime("%Y/%m/%d") if isinstance(_pre_w, date) else ""
        )
        _cur_d = _coerce_task_raw_input_to_date(t.get("raw_input_date"))
        pattern_shift_note = ""
        if _pat_label and isinstance(_pre_w, date) and _cur_d is not None:
            _delta_days = (_pre_w - _cur_d).days
            if _delta_days > 0:
                pattern_shift_note = f"{_pat_label}・原反{_delta_days}日前"
        start_req = t["start_date_req"]
        start_req_s = start_req.strftime("%Y/%m/%d") if hasattr(start_req, "strftime") else str(start_req)
        rov = t.get("required_op")
        # 列順: A=ステータス → 配台状況メモ → タスクID/工程/機械… → 履歴1..n → しの他 → 最後に特別指定_AI
        _memo_dispatch = _result_task_dispatch_status_memo(status, rem_u, hist)
        row_status = {"ステータス": status, "配台状況メモ": _memo_dispatch}
        _dto = t.get("dispatch_trial_order")
        _spd = t.get(TASK_COL_SPEED)
        if _spd is None or (isinstance(_spd, float) and pd.isna(_spd)):
            _spd_out = ""
        elif isinstance(_spd, (int, float)) and not isinstance(_spd, bool):
            _f = float(_spd)
            _spd_out = int(_f) if _f == int(_f) else round(_f, 6)
        else:
            _spd_out = _spd
        row_core = {
            "タスクID": t['task_id'],
            "工程名": t['machine'],
            "機械名": t.get("machine_name", ""),
            TASK_COL_SPEED: _spd_out,
            "優先度": t.get("priority", 999),
            RESULT_TASK_COL_DISPATCH_TRIAL_ORDER: _dto if _dto is not None else "",
        }
        row_history = {}
        _tid_norm = _interactive_norm_cell(t.get("task_id"))
        _mach_norm = _interactive_norm_cell(t.get("machine_name"))
        _hk = (_tid_norm, _mach_norm)
        if _hk in _interactive_hist_override:
            _hist_for_sheet = _interactive_hist_override.get(_hk, [])
        else:
            _hist_for_sheet = merge_assigned_history_contiguous_for_result_sheet(
                t.get("assigned_history")
            )
        for i in range(max_history_len):
            if i < len(_hist_for_sheet):
                h = _hist_for_sheet[i]
                row_history[f"履歴{i+1}"] = _format_result_task_history_cell(t, h)
            else:
                row_history[f"履歴{i+1}"] = ""

        try:
            tot_qty = parse_float_safe(t.get("total_qty_m"), 0.0)
            done_qty = parse_float_safe(t.get("done_qty_reported"), 0.0)
            pct_macro = max(0, min(100, int(round((done_qty / tot_qty) * 100)))) if tot_qty > 0 else 0
        except Exception:
            pct_macro = 0

        _ub = t.get("unprocessed_baseline_m")
        _init_rem_u = float(t.get("initial_remaining_units") or 0)
        if _ub is not None:
            if _init_rem_u > 1e-12:
                _rem_qty_out = float(_ub) * (float(rem_u) / _init_rem_u)
            else:
                _rem_qty_out = float(_ub) if rem_u > 1e-12 else 0.0
        else:
            _rem_qty_out = max(0.0, float(rem_u) * float(t.get("unit_m") or 0))

        # 配台済_加工開始/終了: 設備時間割と同じく工程（eq_line）単位。納期判定は依頼全体の最終終了（従来の task_id 集約）。
        _eq_line_key = (
            str(t.get("equipment_line_key") or t.get("machine") or "").strip()
            or str(t.get("machine") or "").strip()
        )
        _pw_line = (
            plan_window_by_task_line.get((_tid_res, _eq_line_key))
            if _eq_line_key
            else None
        )
        _pw_agg = plan_window_by_task_id.get(t.get("task_id"))
        if _pw_agg is None and _tid_res:
            _pw_agg = plan_window_by_task_id.get(_tid_res)
        _pw_disp = _pw_line or _pw_agg
        if _pw_disp:
            _ps, _pe = _pw_disp[0], _pw_disp[1]
            plan_assign_start_s = (
                _ps.strftime("%Y/%m/%d %H:%M") if hasattr(_ps, "strftime") else ""
            )
            plan_assign_end_s = (
                _pe.strftime("%Y/%m/%d %H:%M") if hasattr(_pe, "strftime") else ""
            )
        else:
            plan_assign_start_s = ""
            plan_assign_end_s = ""

        _plan_end_ans_spec16 = _result_task_plan_end_within_answer_or_spec_16_label(
            _pw_agg, _ans_d, _spec_d, t.get("task_id")
        )

        row_tail = {
            "必須OP(上書)": rov if rov is not None else "",
            "タスク効率": parse_float_safe(t.get("task_eff_factor"), 1.0),
            "加工途中": "はい" if t.get("in_progress") else "いいえ",
            "特別指定あり": "はい" if t.get("has_special_remark") else "いいえ",
            "担当OP指定": (t.get("preferred_operator_raw") or "")[:120],
            "回答納期": ans_s,
            "指定納期": spec_s,
            "計画基準納期": basis_s,
            TASK_COL_RAW_INPUT_DATE: kenhan_s,
            RESULT_TASK_COL_RAW_INPUT_DATE_PRE_PATTERN: pre_pattern_kenhan_s,
            RESULT_TASK_COL_PATTERN_RAW_SHIFT_NOTE: pattern_shift_note,
            "紝期緊急": "はい" if t.get("due_urgent") else "いいえ",
            "加工開始日": start_req_s,
            "配台済_加工開始": plan_assign_start_s,
            "配台済_加工終了": plan_assign_end_s,
            RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16: _plan_end_ans_spec16,
            "累計加工量": f"{total_r}R ({t['total_qty_m']}m)",
            "残加工量": _rem_qty_out,
            "完了率(実行時点)": f"{pct_macro}%",
        }
        row_ai_last = {"特別指定_AI": (t.get("task_special_ai_note") or "")[:300]}
        row_data = {**row_status, **row_core, **row_history, **row_tail, **row_ai_last}
        task_results.append(row_data)

    cal_rows = []
    for d in sorted_dates:
        for m in members:
            if m in attendance_data[d]:
                data = attendance_data[d][m]
                if data['is_working']:
                    cal_end = _calendar_display_clock_out_for_calendar_sheet(data, d)
                    end_disp = cal_end if cal_end is not None else data['end_dt']
                    clock_out_s = end_disp.strftime("%H:%M")
                else:
                    clock_out_s = "休"
                cal_rows.append({
                    "日付": d,
                    "メンバー": m,
                    "出勤": data['start_dt'].strftime("%H:%M") if data['is_working'] else "休",
                    "退勤": clock_out_s,
                    "効率": data['efficiency'],
                    "備考": data['reason'],
                })

    utilization_data = []
    for d in sorted_dates:
        row_data = {"年月日": d.strftime("%Y/%m/%d (%a)")}
        # しの日のイベントからメンバー別作業分を一括集計（全メンバー×全イベントの二重ループを避ける）
        member_worked_mins = defaultdict(int)
        for ev in events_by_date[d]:
            mins = get_actual_work_minutes(ev["start_dt"], ev["end_dt"], ev["breaks"])
            member_worked_mins[ev["op"]] += mins
            for s in ev["subs_list"]:
                if s:
                    member_worked_mins[s] += mins
        for m in members:
            if m not in attendance_data[d]:
                row_data[m] = "休"
                continue
            _ud_u = attendance_data[d][m]
            _eligible_u = bool(
                _ud_u.get(
                    "eligible_for_assignment", _ud_u.get("is_working", False)
                )
            )
            if not _eligible_u:
                if _ud_u.get("is_working", False):
                    # カレンダー「-」等: 勤務だが配台母数に含めない（分子・分母とも 0 表示）
                    row_data[m] = "0.0% (0/0分)"
                else:
                    row_data[m] = "休"
                continue
            default_start = datetime.combine(d, DEFAULT_START_TIME)
            default_end = datetime.combine(d, DEFAULT_END_TIME)

            actual_start = _ud_u["start_dt"]
            actual_end = _ud_u["end_dt"]
            clip_start = max(actual_start, default_start)
            clip_end = min(actual_end, default_end)

            if clip_start >= clip_end:
                total_avail_mins = 0
            else:
                breaks_dt = _ud_u["breaks_dt"]
                total_avail_mins = get_actual_work_minutes(
                    clip_start, clip_end, breaks_dt
                )

            if total_avail_mins <= 0:
                row_data[m] = "0.0%"
                continue

            worked_mins = member_worked_mins.get(m, 0)
            ratio = (worked_mins / total_avail_mins) * 100
            row_data[m] = f"{ratio:.1f}% ({worked_mins}/{total_avail_mins}分)"
        utilization_data.append(row_data)
        
    df_utilization = pd.DataFrame(utilization_data)

    _usage_txt = build_gemini_usage_summary_text()
    if _usage_txt:
        ai_log_data["Gemini_トークン・料金サマリ"] = _usage_txt[:50000]

    _master_abs_for_result_fmt = _master_workbook_path_resolved()
    _reg_shift_start, _reg_shift_end = _read_master_main_regular_shift_times(
        _master_abs_for_result_fmt
    )
    if _reg_shift_start is not None and _reg_shift_end is not None:
        logging.info(
            "定常枠: master メイン A15/B15 → %s ～ %s（結果の定常外「日時帯」着色）",
            _reg_shift_start.strftime("%H:%M"),
            _reg_shift_end.strftime("%H:%M"),
        )

    logging.info(
        "段階2: 結果ブックを作成しした → %s",
        os.path.basename(plan_xlsx_final),
    )
    gantt_tl_label_specs: list = []
    gantt_tl_day_blocks: list = []
    gantt_detail_tl_label_specs: list = []
    gantt_detail_tl_day_blocks: list = []
    detail_timeline_events: list = []
    sorted_dates_detail = list(sorted_dates)
    chart_title_actual_detail = "湖南工場 加工実績（明細）"
    # 試行順パターン別段階2（stage2_output_root あり）は同一処理を多数回走らせるため、
    # 重い設備ガント（計画・加工実績明細）の生成と明細 DATA 読込を省略する（スコア用の結果シートは従来どおり）。
    # JavaFX から PM_AI_STAGE2_WRITE_EXCEL=0 のときも xlsx/JSON 用の一時ブックにガントを書かないよう、
    # 同様に明細 DATA 読込・実績タイムライン構築・設備ガント描画を省略する（JSON は必須シートのみ）。
    # 段階3配台試行（段階2同一パリティ）も同様に省略（結果_配台表 JSON 更新が主目的）。
    if _stage2_skip_actual_detail_gantt_prep(stage2_output_root, _publish_plan_xlsx):
        if stage2_output_root:
            logging.info(
                "段階2(試行順パターン別バッチ): 設備ガント（計画・加工実績明細）の生成を省略します。"
            )
        elif not _publish_plan_xlsx:
            logging.info(
                "段階2: Excel 出力を抑制（PM_AI_STAGE2_WRITE_EXCEL）のため、"
                "設備ガント（計画・加工実績明細）の準備・生成を省略します。"
            )
        elif _interactive_stage2_parity_active():
            logging.info(
                "段階3(配台試行): 加工実績明細 DATA 読込と実績ガント準備を省略します。"
            )
    else:
        # 実績明細ガントのメタ・時間軸: 計画ブックは「加工計画DATA_実績比較用」を優先（無ければ段階2本体と同じ基準）
        base_now_dt_act_gantt = base_now_dt
        data_extract_dt_str_act_gantt = data_extract_dt_str
        _dt_act_plan, _ = _extract_data_extraction_datetime_for_actual_related_gantt()
        if _dt_act_plan is not None:
            _orig_act = _dt_act_plan
            base_now_dt_act_gantt = _dt_act_plan
            if (
                not STAGE2_MACRO_NOW_USE_DATA_EXTRACT_CLOCK
                and isinstance(base_now_dt_act_gantt, datetime)
            ):
                base_now_dt_act_gantt = datetime.combine(
                    base_now_dt_act_gantt.date(), DEFAULT_START_TIME
                )
            data_extract_dt_str_act_gantt = _orig_act.strftime("%Y/%m/%d %H:%M:%S")

        df_actual_detail = load_machining_actual_detail_df()
        if df_actual_detail is not None and len(df_actual_detail) > 0:
            sorted_dates_detail = _sorted_dates_union_actual_bounds_df(
                sorted_dates, df_actual_detail
            )
            d_from = _parse_env_optional_date(ENV_GANTT_ACTUAL_DETAIL_DATE_FROM)
            d_to = _parse_env_optional_date(ENV_GANTT_ACTUAL_DETAIL_DATE_TO)
            if d_from is not None or d_to is not None:
                n_before = len(sorted_dates_detail)
                filtered_detail_dates = _sorted_dates_filter_inclusive_range(
                    sorted_dates_detail, d_from, d_to
                )
                if not filtered_detail_dates and sorted_dates_detail:
                    logging.warning(
                        "実績明細ガント: 日付範囲フィルタで表示日が0件になったためフィルタを無視します。"
                        "（%s=%r, %s=%r）",
                        ENV_GANTT_ACTUAL_DETAIL_DATE_FROM,
                        os.environ.get(ENV_GANTT_ACTUAL_DETAIL_DATE_FROM, ""),
                        ENV_GANTT_ACTUAL_DETAIL_DATE_TO,
                        os.environ.get(ENV_GANTT_ACTUAL_DETAIL_DATE_TO, ""),
                    )
                else:
                    sorted_dates_detail = filtered_detail_dates
                    logging.info(
                        "実績明細ガント: 表示日を %s 日 → %s 日に絞りました（FROM=%s, TO=%s）。",
                        n_before,
                        len(sorted_dates_detail),
                        d_from.isoformat() if d_from else "（指定なし）",
                        d_to.isoformat() if d_to else "（指定なし）",
                    )
                    rng_lo = d_from.isoformat() if d_from else "…"
                    rng_hi = d_to.isoformat() if d_to else "…"
                    chart_title_actual_detail = (
                        f"{chart_title_actual_detail}（表示 {rng_lo}～{rng_hi}）"
                    )

            def _first_valid_dt_from_df_col_stage2_act(_df, _col) -> datetime | None:
                try:
                    if _df is None or _col not in _df.columns:
                        return None
                    for _v in _df[_col].tolist():
                        if _v is None or (isinstance(_v, float) and pd.isna(_v)):
                            continue
                        _dt = pd.to_datetime(_v, errors="coerce")
                        if pd.isna(_dt):
                            continue
                        if isinstance(_dt, pd.Timestamp):
                            return _dt.to_pydatetime()
                        return _dt if isinstance(_dt, datetime) else None
                except Exception:
                    return None
                return None

            detail_extract_dt = _first_valid_dt_from_df_col_stage2_act(
                df_actual_detail, TASK_COL_DATA_EXTRACTION_TIME
            )
            if detail_extract_dt is not None:
                base_now_dt_act_gantt = detail_extract_dt
                data_extract_dt_str_act_gantt = detail_extract_dt.strftime(
                    "%Y/%m/%d %H:%M:%S"
                )

            detail_timeline_events = build_actual_timeline_events(
                df_actual_detail,
                equipment_list,
                sorted_dates_detail,
                log_sheet_name=_actual_detail_sheet_log_label(),
                roll_detail=True,
            )
    # 表シート: DataFrame を先に確定 → JSON 正本 → 同一ペイロードから ExcelWriter へ（ガント・2 段シート除く）
    df_tasks = pd.DataFrame(task_results)
    df_tasks, task_column_order, _, vis_map = apply_result_task_sheet_column_order(
        df_tasks, max_history_len
    )
    try:
        from planning_core.excel_trace_task import log_df_tasks as _excel_trace_df_tasks

        _excel_trace_df_tasks(
            df_tasks,
            "after_apply_result_task_column_order",
            output_basename=os.path.basename(plan_xlsx_final),
        )
    except Exception:
        pass
    if not task_column_order:
        task_column_order, vis_map = _result_task_column_config_fallback_from_existing(
            df_tasks, max_history_len
        )
    seen_tc: set[str] = set()
    task_column_order_dedup: list = []
    vis_list_dedup: list = []
    for c in task_column_order:
        if c in seen_tc:
            continue
        seen_tc.add(c)
        task_column_order_dedup.append(c)
        vis_list_dedup.append(bool(vis_map.get(c, True)))
    df_column_config = pd.DataFrame(
        {
            "列名": task_column_order_dedup,
            "表示": vis_list_dedup,
        }
    )
    _t_ldf0 = time_module.perf_counter()
    try:
        df_src_for_dispatch = load_tasks_df()
    except Exception as e:
        logging.warning("結果_配台表: 加工計画DATA 読込に失敗したため空欄補完をスキップ: %s", e)
        df_src_for_dispatch = None
    _t_stage2_perf = _log_stage2_phase_timing(
        "load_tasks_df",
        _t_ldf0,
        extra={
            "ok": df_src_for_dispatch is not None,
            "processing_plan_path": _stage2_path_debug_meta(
                (os.environ.get("PM_AI_PROCESSING_PLAN_PATH") or "").strip()
            ),
            "task_input_source_dir": _stage2_path_debug_meta(
                (os.environ.get("PM_AI_TASK_INPUT_SOURCE_DIR") or "").strip()
            ),
        },
    )
    if _interactive_dispatch_trial_env_active():
        # 試行中の meters_done は _cap_key 解決で「JSON の配台日」と異なるキーへ載ることがある。
        # 未達一覧はタイムライン上の暦日×依頼×工程×機械で再集計し、短い検証（dispatch_qty_shortfall）と整合させる。
        if interactive_dispatch_targets:
            try:
                _reco_md = _stage35_merge_floor_into_meters_done(
                    _interactive_trial_meters_done_by_timeline_calendar_date(
                        timeline_events,
                        task_queue,
                    )
                )
                _interactive_trial_meters_done.clear()
                _interactive_trial_meters_done.update(_reco_md)
            except Exception as _e_md_reco:
                logging.warning(
                    "インタラクティブ試行: meters_done のタイムライン再集計に失敗しました（累積キーのままスナップショット）: %s",
                    _e_md_reco,
                )
        _LAST_INTERACTIVE_TRIAL_METERS_DONE_SNAPSHOT.clear()
        _LAST_INTERACTIVE_TRIAL_METERS_DONE_SNAPSHOT.update(_interactive_trial_meters_done)
    df_dispatch = build_result_dispatch_table_dataframe(
        timeline_events,
        sorted_tasks_for_result,
        tasks_df,
        df_src_for_dispatch,
    )
    _in_progress_dispatch_plan_day = _stage2_dialog_target_plan_day(
        run_date,
        working_days,
        skip_today=_stage2_truthy_env("PM_AI_STAGE2_SKIP_TODAY_DISPATCH"),
    )
    df_dispatch = append_in_progress_next_day_dialog_rows_to_dispatch_table(
        df_dispatch,
        tasks_df,
        df_src_for_dispatch,
        run_date,
        working_days,
        calendar_today=calendar_today,
        timeline_events=timeline_events,
        sorted_tasks_for_result=sorted_tasks_for_result,
    )
    df_dispatch = append_plan_input_rows_missing_from_dispatch_table(
        df_dispatch, tasks_df, df_src_for_dispatch
    )
    df_dispatch = _apply_dispatch_table_timeline_meta_from_events(
        df_dispatch, timeline_events, sorted_tasks_for_result
    )
    if (
        _interactive_dispatch_trial_env_active()
        and interactive_dispatch_targets
        and not interactive_result_dispatch_json_rows
    ):
        df_dispatch = _stage3_attach_actual_dispatch_qty_from_timeline_plan_rows(
            df_dispatch
        )
    df_dispatch = _interactive_dispatch_trial_use_editor_rows_for_result_table(
        df_dispatch,
        interactive_result_dispatch_json_rows,
        interactive_result_dispatch_json_columns,
        interactive_dispatch_targets=interactive_dispatch_targets,
        timeline_events=timeline_events,
        task_queue=task_queue,
        working_days=working_days,
    )
    if interactive_dispatch_targets:
        _LAST_INTERACTIVE_TRIAL_PLAN_TARGETS_SNAPSHOT.clear()
        _LAST_INTERACTIVE_TRIAL_PLAN_TARGETS_SNAPSHOT.update(
            _interactive_aggregate_plan_dispatch_targets_from_df(df_dispatch)
        )
        _interactive_validate_dispatch_quantities(
            df_dispatch, interactive_dispatch_targets
        )
        _LAST_INTERACTIVE_TRIAL_META_MISS_SHORTFALL[:] = (
            _interactive_dispatch_timeline_meta_miss_shortfalls(
                df_dispatch,
                meters_done=_LAST_INTERACTIVE_TRIAL_METERS_DONE_SNAPSHOT or None,
            )
        )
        for _mm in _LAST_INTERACTIVE_TRIAL_META_MISS_SHORTFALL:
            _tid = _interactive_norm_cell(_mm.get("task_id"))
            _proc = _interactive_norm_cell(_mm.get("process"))
            _mach = _interactive_norm_cell(_mm.get("machine_name"))
            _dd_raw = _mm.get("dispatch_date")
            _dd_parsed = _interactive_parse_dispatch_date_cell(_dd_raw)
            if not _tid or not _mach or _dd_parsed is None:
                continue
            _mk = (_tid, _proc, _mach, _dd_parsed)
            _LAST_INTERACTIVE_TRIAL_METERS_DONE_SNAPSHOT[_mk] = 0.0
            if interactive_dispatch_targets is not None:
                try:
                    _interactive_trial_meters_done[_mk] = 0.0
                except Exception:
                    pass
    if interactive_relax_intraday or interactive_dispatch_targets is not None:
        _interactive_validate_timeline_midnight_if_interactive(timeline_events)
    _interactive_append_machining_end_after_member_shift_shortages(
        timeline_events, attendance_data
    )
    df_ai_log = pd.DataFrame(list(ai_log_data.items()), columns=["項目", "内容"])
    _t_stage2_perf = _log_stage2_phase_timing(
        "build_result_sheet_dataframes",
        _t_build0,
        extra={
            "dispatch_rows": len(df_dispatch) if df_dispatch is not None else 0,
            "task_rows": len(df_tasks) if df_tasks is not None else 0,
        },
    )

    from planning_core.workbook_payload import (
        build_workbook_payload_from_dataframes,
        write_tabular_source_json_file,
        write_tabular_sheets_from_payload_to_excel_writer,
    )

    _stage2_tabular_sheet_order = [
        RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME,
        RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME,
        "結果_カレンダー(出勤簿)",
        RESULT_MEMBER_WORK_UTIL_SHEET_NAME,
        COLUMN_CONFIG_SHEET_NAME,
        RESULT_TASK_SHEET_NAME,
        RESULT_DISPATCH_TABLE_SHEET_NAME,
        "結果_AIログ",
    ]
    _stage2_tabular_dfs = {
        RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME: df_eq_schedule,
        RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME: df_equipment_by_machine_name,
        "結果_カレンダー(出勤簿)": pd.DataFrame(cal_rows),
        RESULT_MEMBER_WORK_UTIL_SHEET_NAME: df_utilization,
        COLUMN_CONFIG_SHEET_NAME: df_column_config,
        RESULT_TASK_SHEET_NAME: df_tasks,
        RESULT_DISPATCH_TABLE_SHEET_NAME: df_dispatch,
        "結果_AIログ": df_ai_log,
    }
    _stage2_tabular_payload = build_workbook_payload_from_dataframes(
        _stage2_tabular_dfs,
        source_xlsx_basename=os.path.basename(plan_xlsx_final),
        metadata_extra={
            "schema": "stage2_tabular_source_v1",
            "excel_tabular_sheets_rendered_from_this_payload": True,
        },
    )
    _t_tab_json0 = time_module.perf_counter()
    try:
        _tabular_json_path, _tabular_json_strat = write_tabular_source_json_file(
            plan_xlsx_final, _stage2_tabular_payload
        )
        if _tabular_json_path:
            logging.info(
                "段階2: 表シート正本 JSON（Excel より先）を '%s' に出力しました（%s）。",
                _tabular_json_path,
                _tabular_json_strat,
            )
    except Exception as e:
        logging.warning("段階2: 表シート正本 JSON の出力に失敗しました: %s", e)
    _t_stage2_perf = _log_stage2_phase_timing(
        "write_tabular_source_json",
        _t_tab_json0,
        extra={"plan_xlsx_final": _stage2_path_debug_meta(plan_xlsx_final)},
    )

    _t_xlw0 = time_module.perf_counter()
    try:
        with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
            write_tabular_sheets_from_payload_to_excel_writer(
                writer,
                _stage2_tabular_payload,
                sheet_order=_stage2_tabular_sheet_order,
            )

            from planning_core.gantt_render_contract import (
                make_gantt_render_contract,
                write_gantt_contract_json,
            )

            _write_excel_gantt_sheets = _stage2_write_excel_gantt_sheets_enabled()

            if not stage2_output_root and _publish_plan_xlsx:
                if _write_excel_gantt_sheets:
                    logging.info(
                        "段階2: 設備ガントチャートを生成（データ量により数分かかることがあります）"
                    )
                else:
                    logging.info(
                        "段階2: 計画ブックへの設備ガント Excel シートは省略します（"
                        "設備ガント(グラフィック)用の契約 JSON のみ出力。ブックへ含める場合は "
                        "PM_AI_STAGE2_WRITE_EXCEL_GANTT_SHEETS=1）。"
                    )
                try:
                    from planning_core.excel_trace_task import (
                        log_gantt_label_specs as _excel_trace_gantt_specs,
                        log_timeline_events as _excel_trace_timeline,
                    )

                    _excel_trace_timeline(
                        timeline_events, "before_write_results_equipment_gantt_sheet"
                    )
                except Exception:
                    pass

                if not timeline_events:
                    logging.warning(
                        "段階2: 設備ガント契約 JSON に書き込む timeline_events が空です。"
                        " Java の設備ガント（グラフィック）や契約からの再描画ではタイムラインに文言が出ません。"
                        " 配台ループでイベントが追記されているか、タスクキュー・稼働メンバー・計画日を確認してください。"
                    )

                _equipment_gantt_contract = make_gantt_render_contract(
                    timeline_events=timeline_events,
                    equipment_list=equipment_list,
                    sorted_dates=sorted_dates,
                    attendance_data=attendance_data,
                    data_extract_dt_str=data_extract_dt_str,
                    base_now_dt=base_now_dt,
                    regular_shift_times=(_reg_shift_start, _reg_shift_end),
                    plan_rows=True,
                    kind="equipment_gantt",
                )
                try:
                    _eg_path, _eg_strat = write_gantt_contract_json(
                        plan_xlsx_final, "equipment", _equipment_gantt_contract
                    )
                    if _eg_path:
                        logging.info(
                            "段階2: 設備ガント描画契約 JSON（再描画用）を '%s' に出力しました（%s）。",
                            _eg_path,
                            _eg_strat,
                        )
                except Exception as _e_egc:
                    logging.warning("段階2: 設備ガント契約 JSON 出力をスキップ: %s", _e_egc)
                if _write_excel_gantt_sheets:
                    from planning_core.gantt_render_contract import (
                        render_gantt_sheet_from_contract,
                    )

                    gantt_tl_label_specs, gantt_tl_day_blocks = (
                        render_gantt_sheet_from_contract(writer, _equipment_gantt_contract)
                    )
                    try:
                        from planning_core.excel_trace_task import (
                            log_gantt_label_specs as _excel_trace_gantt_specs,
                        )

                        _excel_trace_gantt_specs(
                            gantt_tl_label_specs,
                            "after_write_results_equipment_gantt_sheet",
                        )
                    except Exception:
                        pass
            else:
                try:
                    from planning_core.excel_trace_task import append as _excel_trace_append

                    _excel_trace_append(
                        {
                            "stage": "equipment_gantt_branch",
                            "hypothesisId": "EX3",
                            "message": "設備ガントシート生成をスキップした分岐",
                            "stage2_output_root": bool(stage2_output_root),
                            "publish_plan_xlsx": bool(_publish_plan_xlsx),
                        }
                    )
                except Exception:
                    pass

            if detail_timeline_events:
                if _write_excel_gantt_sheets:
                    logging.info(
                        "段階2: 設備ガントチャート（加工実績明細）を生成します（データ量により時間がかかることがあります）"
                    )
                else:
                    logging.info(
                        "段階2: 計画ブックへの実績明細ガント Excel シートは省略します（契約 JSON のみ）。"
                    )

                _actual_detail_gantt_contract = make_gantt_render_contract(
                    timeline_events=[],
                    equipment_list=equipment_list,
                    sorted_dates=sorted_dates_detail,
                    attendance_data=attendance_data,
                    data_extract_dt_str=data_extract_dt_str_act_gantt,
                    base_now_dt=base_now_dt_act_gantt,
                    actual_timeline_events=detail_timeline_events,
                    regular_shift_times=(_reg_shift_start, _reg_shift_end),
                    plan_rows=False,
                    chart_title=chart_title_actual_detail,
                    sheet_name_override=RESULT_SHEET_GANTT_ACTUAL_DETAIL_NAME,
                    kind="actual_detail_gantt",
                )
                try:
                    _ad_path, _ad_strat = write_gantt_contract_json(
                        plan_xlsx_final, "actual_detail", _actual_detail_gantt_contract
                    )
                    if _ad_path:
                        logging.info(
                            "段階2: 実績明細ガント描画契約 JSON（再描画用）を '%s' に出力しました（%s）。",
                            _ad_path,
                            _ad_strat,
                        )
                except Exception as _e_adgc:
                    logging.warning("段階2: 実績明細ガント契約 JSON 出力をスキップ: %s", _e_adgc)
                if _write_excel_gantt_sheets:
                    from planning_core.gantt_render_contract import (
                        render_gantt_sheet_from_contract,
                    )

                    (
                        gantt_detail_tl_label_specs,
                        gantt_detail_tl_day_blocks,
                    ) = render_gantt_sheet_from_contract(
                        writer, _actual_detail_gantt_contract
                    )

            for sheet_name, ws_out in writer.sheets.items():
                if sheet_name in (
                    RESULT_SHEET_GANTT_NAME,
                    RESULT_SHEET_GANTT_ACTUAL_DETAIL_NAME,
                ):
                    continue
                _apply_output_font_to_result_sheet(ws_out)

            # 段階2 結果ブックの Excel 表現は以下を維持すること（簡略化・削除しない）。
            # - 「結果_設備ガント」系シートは既定では含めない（PM_AI_STAGE2_WRITE_EXCEL_GANTT_SHEETS で任意復活）。
            #   表示は設備ガント(グラフィック)＋描画契約 JSON を正とする。
            # - 条件付き書式に相当する着色（設備時間割の定常外・準備・機械別の強調、タスク一覧の
            #   未配台行・履歴 need/surplus・セル不一致・納期判定ハイライトなど）
            # - ハイパーリンク（結果_タスク一覧の依頼NO → 結果_設備毎の時間割）
            # - 列幅調整（結果_設備毎の時間割はヘッダー長基準のオートフィット相当）
            # - Excel テーブル化（配台表シート）
            # 共通の見出し背景・罫線・固定窓枠は末尾の _apply_stage2_production_plan_workbook_polish。

            if _reg_shift_start is not None and _reg_shift_end is not None:
                for _eq_sched_sheet in (
                    RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME,
                    RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME,
                ):
                    if _eq_sched_sheet in writer.sheets:
                        _apply_equipment_schedule_outside_regular_fill(
                            writer.sheets[_eq_sched_sheet],
                            _reg_shift_start,
                            _reg_shift_end,
                        )

            if RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME in writer.sheets:
                _apply_equipment_by_machine_dispatched_request_fill(
                    writer.sheets[RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME]
                )

            if RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME in writer.sheets:
                _apply_equipment_schedule_prep_cleanup_fill(
                    writer.sheets[RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME]
                )

            if RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME in writer.sheets:
                _apply_equipment_schedule_machine_calendar_fill(
                    writer.sheets[RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME],
                    equipment_list,
                    _MACHINE_CALENDAR_BLOCKS_BY_DATE,
                )

            ws_cfg = writer.sheets[COLUMN_CONFIG_SHEET_NAME]
            _add_column_config_sheet_helpers(ws_cfg, len(task_column_order_dedup))

            worksheet_tasks = writer.sheets[RESULT_TASK_SHEET_NAME]
            max_col = worksheet_tasks.max_column
            for row in worksheet_tasks.iter_rows(min_row=1, max_row=worksheet_tasks.max_row, max_col=max_col):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False, vertical="top")

            _apply_result_task_sheet_column_visibility(
                worksheet_tasks, list(df_tasks.columns), vis_map
            )

            _apply_result_task_history_rich_text(worksheet_tasks, list(df_tasks.columns))
            _apply_result_task_date_columns_blue_font(worksheet_tasks, list(df_tasks.columns))

            # 未スケジュール行（配台不可・配台残）を目立たせる
            status_col_idx = None
            for col_idx, col_name in enumerate(df_tasks.columns, 1):
                if str(col_name) == "ステータス":
                    status_col_idx = col_idx
                    break
            if status_col_idx is not None:
                unscheduled_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                for r in range(2, worksheet_tasks.max_row + 1):
                    st_val = worksheet_tasks.cell(row=r, column=status_col_idx).value
                    st = str(st_val).strip() if st_val is not None else ""
                    if st in ("配台不可", "配台残", "配台試行"):
                        for c in range(1, max_col + 1):
                            worksheet_tasks.cell(row=r, column=c).fill = unscheduled_fill

            _apply_result_task_history_need_surplus_highlight(
                worksheet_tasks, list(df_tasks.columns), sorted_tasks_for_result
            )

            _apply_result_task_task_id_content_mismatch_highlight(
                worksheet_tasks, list(df_tasks.columns), sorted_tasks_for_result
            )
            _apply_result_task_plan_end_answer_spec_16_no_highlight(
                worksheet_tasks, list(df_tasks.columns)
            )
            _apply_result_task_id_hyperlinks_to_equipment_schedule(
                worksheet_tasks,
                list(df_tasks.columns),
                sorted_tasks_for_result,
                first_eq_schedule_cell_by_task_id,
                RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME,
            )

            if RESULT_DISPATCH_TABLE_SHEET_NAME in writer.sheets and len(df_dispatch) > 0:
                _apply_result_dispatch_table_excel_table(
                    writer.sheets[RESULT_DISPATCH_TABLE_SHEET_NAME],
                    table_display_name=RESULT_DISPATCH_TABLE_EXCEL_TABLE_NAME,
                )

            if RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME in writer.sheets:
                _apply_equipment_schedule_day_banner_row_style(
                    writer.sheets[RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME]
                )
                _apply_equipment_schedule_auto_column_widths(
                    writer.sheets[RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME]
                )

            _apply_stage2_production_plan_workbook_polish(
                writer.sheets,
                member_priority_second_header_row=None,
            )

            if RESULT_TASK_SHEET_NAME in writer.sheets:
                _apply_result_task_sheet_layout_polish(
                    writer.sheets[RESULT_TASK_SHEET_NAME],
                    list(df_tasks.columns),
                )

            if RESULT_DISPATCH_TABLE_SHEET_NAME in writer.sheets:
                _apply_result_dispatch_table_sheet_layout_polish(
                    writer.sheets[RESULT_DISPATCH_TABLE_SHEET_NAME]
                )

    except OSError as e:
        logging.error(
            "段階2: 結果ブックの作成・保存に失敗しました: %s（%s）。"
            "output 内の 計画*.xlsx を Excel で開いでいないか確認してください。",
            output_filename,
            e,
        )
        raise
    _t_stage2_perf = _log_stage2_phase_timing(
        "write_plan_workbook_excel",
        _t_xlw0,
        extra={
            "output_filename": _stage2_path_debug_meta(output_filename),
            "write_excel_gantt_sheets": _stage2_write_excel_gantt_sheets_enabled(),
        },
    )

    try:
        _sj = write_result_task_json_sidecar(
            plan_xlsx_final, df_tasks, sheet_name=RESULT_TASK_SHEET_NAME
        )
        try:
            from planning_core.excel_trace_task import (
                log_sidecar_result_task_row as _excel_trace_sidecar,
            )

            _excel_trace_sidecar(_sj)
        except Exception:
            pass
        if _sj:
            logging.info(
                "段階2: 結果_タスク一覧 JSON サイドカーを '%s' に出力しました。",
                _sj,
            )
    except Exception as e:
        logging.warning("段階2: 結果_タスク一覧 JSON サイドカー出力をスキップ: %s", e)

    try:
        _apply_excel_date_columns_date_only_display(
            output_filename, "結果_カレンダー(出勤簿)", frozenset({"日付"})
        )
    except Exception as e:
        logging.warning(f"結果_カレンダー(出勤簿)の日付列表示整形: {e}")

    _stage2_try_copy_column_config_shapes_from_input(
        output_filename,
        _excel_plan_input_wb(),
    )

    _stage2_try_add_gantt_timeline_shape_labels(
        output_filename, gantt_tl_label_specs, gantt_tl_day_blocks
    )
    if gantt_detail_tl_label_specs:
        _stage2_try_add_gantt_timeline_shape_labels(
            output_filename,
            gantt_detail_tl_label_specs,
            gantt_detail_tl_day_blocks,
            sheet_name=RESULT_SHEET_GANTT_ACTUAL_DETAIL_NAME,
        )

    # 計画ブック全シート JSON: セル値のみ。抽出ロジックは workbook_payload（設備ガント列見出しの reheader 込み）。
    _plan_wb_json = None
    try:
        _meta_wb = (
            {"source_xlsx": os.path.basename(plan_xlsx_final)}
            if not _publish_plan_xlsx
            else None
        )
        _plan_wb_json = write_production_plan_workbook_json(
            output_filename,
            json_out_path=(
                normalized_workbook_json_path(plan_xlsx_final)
                if not _publish_plan_xlsx
                else None
            ),
            metadata_extra=_meta_wb,
        )
        if _plan_wb_json:
            logging.info(
                "段階2: 計画ブック（全シート）の JSON を '%s' に出力しました。",
                _plan_wb_json,
            )
    except Exception as e:
        logging.warning("段階2: 計画ブック JSON（全シート）出力をスキップ: %s", e)

    try:
        from .logical_workbook_view import logical_view_json_path

        if _stage2_skip_heavy_workbook_sidecars():
            logging.info(
                "段階3(配台試行): 計画ブック 論理ビュー JSON の出力を省略します。"
            )
        else:
            _logical_view_out = (
                logical_view_json_path(plan_xlsx_final) if not _publish_plan_xlsx else None
            )
            _plan_lv_json = write_production_plan_logical_view_json(
                output_filename, json_out_path=_logical_view_out
            )
            if _plan_lv_json:
                logging.info(
                    "段階2: 計画ブック 論理ビュー JSON（結合展開）を '%s' に出力しました。",
                    _plan_lv_json,
                )
    except Exception as e:
        logging.warning("段階2: 論理ビュー JSON 出力をスキップ: %s", e)

    if not _publish_plan_xlsx:
        try:
            os.remove(output_filename)
        except OSError as _rm_err:
            logging.warning(
                "段階2: 一時計画ブック xlsx の削除に失敗しました: %s (%s)",
                output_filename,
                _rm_err,
            )

    if _publish_plan_xlsx:
        logging.info(f"完了: '{plan_xlsx_final}' を生成しました。")
    else:
        logging.info(
            "段階2: 計画ブック xlsx は出力しませんでした（PM_AI_STAGE2_WRITE_EXCEL）。"
            " JSON のみ: %s",
            normalized_workbook_json_path(plan_xlsx_final),
        )

    # ---------------------------------------------------------
    # 追加出力: Power Query 用「結果_配台表.xlsx」＋同一データの JSON（既定は repo の code/output など）
    # ---------------------------------------------------------
    _t_dispatch_out0 = time_module.perf_counter()
    _dispatch_table_out_dir: str | None = None
    try:
        _wb_path = _excel_plan_input_wb()
        _out_dir = resolve_result_dispatch_table_output_dir(_wb_path)
        _dispatch_table_out_dir = _out_dir
        if not _out_dir:
            _out_dir = _stage2_out_root or output_dir
            _dispatch_table_out_dir = _out_dir
            logging.info(
                "結果_配台表: 専用出力先が解決できなかったため、段階2成果物フォルダへ出します → %s",
                _out_dir,
            )
        _wrote = None
        if _publish_plan_xlsx:
            _wrote = _write_dispatch_table_standalone_xlsx(df_dispatch, _out_dir)
        if _wrote:
            logging.info("段階2: PowerQuery 用に '%s' を出力しました。", _wrote)
        _jwrote = _write_dispatch_table_standalone_json_to_resolved_dirs(
            df_dispatch, _out_dir, _wb_path
        )
        if _jwrote:
            logging.info("段階2: 結果_配台表 JSON を '%s' に出力しました。", _jwrote)
    except Exception as e:
        logging.warning("段階2: 結果_配台表.xlsx / .json の出力をスキップしました: %s", e)
    _t_stage2_perf = _log_stage2_phase_timing(
        "write_result_dispatch_table_outputs",
        _t_dispatch_out0,
        extra={
            "out_dir": _stage2_path_debug_meta(_dispatch_table_out_dir),
        },
    )

    # =========================================================
    # 5. ★追加: メンバー毎の行動スケジュール (別ファイル) 出力
    # =========================================================
    member_xlsx_final = os.path.join(
        _stage2_out_root, member_workbook_filename(_stage2_stamp)
    )
    member_schedule_json_path = None
    if _stage2_skip_member_schedule_output():
        logging.info(
            "段階3(配台試行): メンバー別スケジュール出力を省略します。"
        )
    else:
        if _publish_plan_xlsx:
            member_output_filename = member_xlsx_final
        else:
            import tempfile

            _fd_mem_tmp, member_output_filename = tempfile.mkstemp(
                suffix=".xlsx", prefix="_pm_stage2_member_", dir=_stage2_out_root
            )
            os.close(_fd_mem_tmp)

        # 時間帯は全メンバー共通で1回の値生成（メンバー数分の重複計算を避ける）
        time_labels = []
        time_grids = []
        curr_dt = datetime.combine(run_date, DEFAULT_START_TIME)
        end_dt_grid = datetime.combine(run_date, DEFAULT_END_TIME)
        while curr_dt < end_dt_grid:
            next_dt = curr_dt + timedelta(minutes=10)
            if next_dt > end_dt_grid:
                next_dt = end_dt_grid
            time_labels.append(f"{curr_dt.strftime('%H:%M')}-{next_dt.strftime('%H:%M')}")
            time_grids.append((curr_dt.time(), next_dt.time()))
            curr_dt = next_dt

        logging.info(
            "段階2: メンバー別スケジュールを作成しした → %s",
            os.path.basename(member_xlsx_final),
        )
        _t_member_xlw0 = time_module.perf_counter()
        try:
            with pd.ExcelWriter(member_output_filename, engine="openpyxl") as member_writer:
                for m in members:
                    # 坄行の辞書を初期化
                    m_schedule = {t_label: {"時間帯": t_label} for t_label in time_labels}

                    # 坄日付のスケジュールを列として埋ゝでいし
                    for d in sorted_dates:
                        d_str = d.strftime("%m/%d (%a)")

                        # 全日非勤務: 年休（カレンダー *）は『年休」」工場休日などは『休」
                        if m not in attendance_data[d] or not attendance_data[d][m]['is_working']:
                            off_label = _member_schedule_full_day_off_label(
                                attendance_data[d].get(m) if m in attendance_data[d] else None
                            )
                            for t_label in time_labels:
                                m_schedule[t_label][d_str] = off_label
                            continue

                        daily_info = attendance_data[d][m]
                        d_start_dt = daily_info['start_dt']
                        d_end_dt = daily_info['end_dt']
                        breaks_dt = daily_info['breaks_dt']

                        events_today = events_by_date[d]

                        for i, (t_start, t_end) in enumerate(time_grids):
                            t_label = time_labels[i]

                            # 判定用の中間時刻を計算
                            grid_start_dt = datetime.combine(d, t_start)
                            grid_end_dt = datetime.combine(d, t_end)
                            grid_mid_dt = grid_start_dt + (grid_end_dt - grid_start_dt) / 2

                            text = ""
                            if grid_mid_dt < d_start_dt or grid_mid_dt >= d_end_dt:
                                text = _member_schedule_off_shift_label(
                                    d, grid_mid_dt, d_start_dt, d_end_dt, daily_info.get("reason")
                                )
                            else:
                                br_txt = _member_schedule_break_cell_label(
                                    grid_mid_dt, breaks_dt, d_end_dt, daily_info.get("reason")
                                )
                                if br_txt is not None:
                                    text = br_txt
                            if text == "":
                                # 該当れるタスクを探れ（subs_list は事剝解析済み）
                                active_ev = next((e for e in events_today if e['start_dt'] <= grid_mid_dt < e['end_dt'] and (e['op'] == m or m in e.get('subs_list', []))), None)
                                if active_ev:
                                    role = "主" if active_ev['op'] == m else "補"
                                    text = f"[{active_ev['task_id']}] {active_ev['machine']}({role})"
                                else:
                                    text = "" # 何も割り当でられでいない空し時間

                            m_schedule[t_label][d_str] = text

                    # データフレーム化してシートに書き込み
                    df_m = pd.DataFrame(list(m_schedule.values()))
                    cols = ["時間帯"] + [d.strftime("%m/%d (%a)") for d in sorted_dates]
                    df_m = df_m[[c for c in cols if c in df_m.columns]]
                    df_m.to_excel(member_writer, sheet_name=m, index=False)

                    # --- 既定フォント・罫線・見出し背景（列幅は VBA 取り込み時の AutoFit） ---
                    worksheet = member_writer.sheets[m]
                    _apply_output_font_to_result_sheet(worksheet)
                    header_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
                    for cell in worksheet[1]:
                        cell.fill = header_fill

                    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                    for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, max_col=worksheet.max_column):
                        for cell in row:
                            cell.border = thin_border

        except OSError as e:
            logging.error(
                "段階2: メンバー別スケジュールの保存に失敗しました: %s（%s）。"
                "人員*.xlsx を Excel で開いでいないか確認してください。",
                member_output_filename,
                e,
            )
            raise
        _t_stage2_perf = _log_stage2_phase_timing(
            "write_member_schedule_excel",
            _t_member_xlw0,
            extra={
                "member_output": _stage2_path_debug_meta(member_output_filename),
                "member_count": len(members),
            },
        )

        if _publish_plan_xlsx:
            logging.info(
                f"完了: 個人別スケジュールを '{member_output_filename}' に出力しました。"
            )
        else:
            logging.info(
                "段階2: メンバー別スケジュール xlsx は出力しませんでした（PM_AI_STAGE2_WRITE_EXCEL）。"
            )

        try:
            _meta_ms = (
                {"source_xlsx": os.path.basename(member_xlsx_final)}
                if not _publish_plan_xlsx
                else None
            )
            member_schedule_json_path = write_member_schedule_workbook_json(
                member_output_filename,
                json_out_path=(
                    normalized_workbook_json_path(member_xlsx_final)
                    if not _publish_plan_xlsx
                    else None
                ),
                metadata_extra=_meta_ms,
            )
            if member_schedule_json_path:
                logging.info(
                    "段階2: メンバー別スケジュール JSON を '%s' に出力しました。",
                    member_schedule_json_path,
                )
        except Exception as e:
            logging.warning("段階2: メンバー別スケジュール JSON 出力をスキップ: %s", e)

        if not _publish_plan_xlsx:
            try:
                os.remove(member_output_filename)
            except OSError as _rm_mem_err:
                logging.warning(
                    "段階2: 一時メンバー別スケジュール xlsx の削除に失敗しました: %s (%s)",
                    member_output_filename,
                    _rm_mem_err,
                )

    if return_output_paths:
        _pp_json = normalized_workbook_json_path(plan_xlsx_final)
        _ms_json = normalized_workbook_json_path(member_xlsx_final)
        out_paths = {
            "production_plan": os.path.abspath(
                plan_xlsx_final if _publish_plan_xlsx else (_plan_wb_json or _pp_json)
            ),
            "member_schedule": os.path.abspath(
                member_xlsx_final
                if _publish_plan_xlsx
                else (member_schedule_json_path or _ms_json)
            ),
        }
        if member_schedule_json_path:
            out_paths["member_schedule_json"] = os.path.abspath(member_schedule_json_path)
        _PLAN_IMPL_INTERACTIVE_DISPATCH_TARGETS = None
        _PLAN_IMPL_INTERACTIVE_TRIAL_METERS_DONE = None
        return out_paths
    _PLAN_IMPL_INTERACTIVE_DISPATCH_TARGETS = None
    _PLAN_IMPL_INTERACTIVE_TRIAL_METERS_DONE = None
    return None

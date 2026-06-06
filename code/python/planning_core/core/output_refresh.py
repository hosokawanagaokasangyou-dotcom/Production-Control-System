# -*- coding: utf-8 -*-
# planning_core.core.output_refresh — body only (loaded via _core exec chain)
def refresh_plan_input_dispatch_trial_order_via_openpyxl(
    workbook_path: str | None = None,
    *,
    apply_post_load_mutations: bool = True,
) -> bool:
    """
    マクロブックの「配台計画_タスク入力」について、
    段階2 と同じ ``fill_plan_dispatch_trial_order_column_stage1`` で「配台試行順番」を
    再付与し、段階1 出力直前と同じ手順で行を並べ替える（ディスク上のブックを openpyxl で読み書き）。

    事前処理は ``_apply_planning_sheet_post_load_mutations``（設定シートの行同期・分割行の自動配台不要）。
    **「設定_配台不要工程」の C/E による計画シートへの配台不要の上書きは行わない**（段階1のみ。
    段階2の ``load_planning_tasks_df`` も同様に再適用しない）。シート上で消した「配台不要」は本経路では復活しない。
    """
    path = (workbook_path or "").strip() or _excel_plan_input_wb().strip()
    if not path:
        logging.error("配台試行順番更新: ブックパスは空です。")
        return False
    if _workbook_should_skip_openpyxl_io(path):
        logging.error(
            "配台試行順番更新: ブックに「%s」があるため openpyxl で編集できません。",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return False
    keep_vba = str(path).lower().endswith(".xlsm")
    wb = None
    try:
        wb = load_workbook(path, keep_vba=keep_vba, read_only=False, data_only=False)
    except Exception as e:
        logging.error("配台試行順番更新: ブックを開けません: %s", e)
        return False
    try:
        try:
            ws = wb[PLAN_INPUT_SHEET_NAME]
        except KeyError as e:
            logging.error("配台試行順番更新: シート接続に失敗: %s", e)
            return False

        mat = _openpyxl_sheet_to_matrix(ws)
        df = _matrix_to_dataframe_header_first(mat)
        if df is None or df.empty:
            logging.warning("配台試行順番更新: データ行はありません。")
            return False

        df = df.copy()
        df.columns = df.columns.str.strip()
        df = _align_dataframe_headers_to_canonical(df, plan_input_sheet_column_order())
        for c in plan_input_sheet_column_order():
            if c not in df.columns:
                df[c] = ""

        df.insert(0, _PLAN_INPUT_XLWINGS_ORIG_ROW, range(len(df)))

        _apply_planning_sheet_post_load_mutations(
            df,
            path,
            "配台試行順番更新",
            apply_exclude_rules_from_config=False,
            compile_exclude_rules_d_to_e_with_ai=False,
        )

        dto_col = RESULT_TASK_COL_DISPATCH_TRIAL_ORDER
        if dto_col not in df.columns:
            logging.error("配台試行順番更新: 列「%s」はありません。", dto_col)
            return False

        _dto_loc = df.columns.get_loc(dto_col)
        if isinstance(_dto_loc, slice):
            logging.error("配台試行順番更新: 列「%s」は複数ありした。", dto_col)
            return False
        if pd.api.types.is_numeric_dtype(df[dto_col]):
            df[dto_col] = float("nan")
        else:
            df[dto_col] = ""

        data_extract_dt, _ = _extract_data_extraction_datetime()
        base_now_dt = data_extract_dt if data_extract_dt is not None else datetime.now()
        run_date = base_now_dt.date()

        try:
            (
                _sd,
                _mem,
                equipment_list,
                req_map,
                need_rules,
                _sm,
                need_combo_col_index,
            ) = load_skills_and_needs()
        except Exception as e:
            logging.exception("配台試行順番更新: master 読込に失敗: %s", e)
            return False

        try:
            fill_plan_dispatch_trial_order_column_stage1(
                df,
                run_date,
                req_map,
                need_rules,
                need_combo_col_index,
                equipment_list,
                members_for_gpo=_mem,
            )
        except Exception as e:
            logging.exception("配台試行順番更新: 試行順計算に失敗: %s", e)
            return False

        df_sorted = _sort_stage1_plan_df_by_dispatch_trial_order_asc(df)
        orig_list = [int(x) for x in df_sorted[_PLAN_INPUT_XLWINGS_ORIG_ROW].tolist()]
        df_sorted = df_sorted.drop(columns=[_PLAN_INPUT_XLWINGS_ORIG_ROW])

        header_row = mat[0] if mat else []
        n_hdr = len(header_row)
        if n_hdr == 0:
            return False

        def _pad_row(r, n):
            r = list(r) if r is not None else []
            if len(r) < n:
                r = r + [None] * (n - len(r))
            return r

        new_mat = [_pad_row(header_row, n_hdr)]
        for i in range(len(df_sorted)):
            orig = orig_list[i]
            src_row = mat[orig + 1] if orig + 1 < len(mat) else []
            src_row = _pad_row(src_row, n_hdr)
            out_row = []
            for j in range(n_hdr):
                h_cell = header_row[j]
                if h_cell is None or (isinstance(h_cell, float) and pd.isna(h_cell)):
                    hname = ""
                else:
                    hname = str(h_cell).strip()
                if hname and hname in df_sorted.columns:
                    v = df_sorted.iat[i, df_sorted.columns.get_loc(hname)]
                    if pd.isna(v):
                        out_row.append(None)
                    else:
                        out_row.append(v)
                else:
                    out_row.append(src_row[j])
            new_mat.append(out_row)

        n_r = len(new_mat)
        for r in range(1, n_r + 1):
            for c in range(1, n_hdr + 1):
                ws.cell(row=r, column=c).value = new_mat[r - 1][c - 1]

        try:
            wb.save(path)
        except Exception as e:
            logging.warning("配台試行順番更新: Save 警告: %s", e)

        logging.info(
            "配台試行順番更新: 「%s」を %s 行で更新しました。",
            PLAN_INPUT_SHEET_NAME,
            len(df_sorted),
        )
        return True
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
def refresh_plan_input_dispatch_trial_order_via_xlwings(
    workbook_path: str | None = None,
    *,
    apply_post_load_mutations: bool = True,
) -> bool:
    """互換名（関数名は後方互換）。"""
    return refresh_plan_input_dispatch_trial_order_via_openpyxl(
        workbook_path, apply_post_load_mutations=apply_post_load_mutations
    )
def refresh_plan_input_dispatch_trial_order_only() -> bool:
    """TASK_INPUT_WORKBOOK に対れる配台試行順番再計算（VBA / cmd 経由のエントリ）。
    環境変数 PLAN_INPUT_DISPATCH_TRIAL_ORDER_LOCAL_ONLY=1 等でシート上のセル値のみを入力とれる。
    """
    p = _excel_plan_input_wb()
    local = _plan_input_dispatch_trial_order_local_only_from_env()
    return refresh_plan_input_dispatch_trial_order_via_xlwings(
        p, apply_post_load_mutations=not local
    )
def refresh_dispatch_trial_pattern_list_sheet_only() -> bool:
    """TASK_INPUT_WORKBOOK に対する配台試行順パターン一覧シート作成（VBA / cmd 用）。"""
    p = _excel_plan_input_wb()
    local = _plan_input_dispatch_trial_order_local_only_from_env()
    return write_dispatch_trial_pattern_list_via_xlwings(
        p, apply_post_load_mutations=not local
    )
def refresh_dispatch_trial_pattern_stage2_batch_only() -> bool:
    """各パターンで段階2を実行しサマリをマクロブックに書く（VBA / cmd 用）。"""
    p = _excel_plan_input_wb()
    local = _plan_input_dispatch_trial_order_local_only_from_env()
    return run_dispatch_trial_pattern_stage2_batch_via_xlwings(
        p, apply_post_load_mutations=not local
    )
def refresh_dispatch_pattern_stage2_selection_to_plan_only() -> bool:
    """サマリで選んだパターンの試行順を配台計画シートへ反映（VBA / cmd 用）。"""
    p = _excel_plan_input_wb()
    local = _plan_input_dispatch_trial_order_local_only_from_env()
    return apply_dispatch_pattern_stage2_selection_to_plan_via_xlwings(
        p, apply_post_load_mutations=not local
    )
def generate_plan():
    """
    段階2のメイン処理。戻り値なし（ログ・Excel 出力で完絝）。

    オーケストレーションの正本は ``planning_core.stage2_identical_dispatch_runner``。

    前提: 環境変数 TASK_INPUT_WORKBOOK、カレントディレクトリがスクリプトフォルダ。
    出力: ``output_dir`` 直下の ``計画*.xlsx`` / ``人員*.xlsx``（実行直前に同名パターンを削除しようとする。ファイル名は短い日本語接頭辞＋時刻で実行ごとに一意）、および log/execution_log.txt。
    """
    from planning_core.stage2_identical_dispatch_runner import run_stage2_generate_plan

    run_stage2_generate_plan()
def refresh_equipment_gantt_actual_detail_only() -> str:
    """
    段階2全体を実行せず、「結果_設備ガント_実績明細」相当のシートだけを
    ``output_dir`` 直下の ``ACTUAL_DETAIL_GANTT_REFRESH_FILENAME`` に出力する。

    マクロブックの勤怠・実績明細DATA・master（工場枠・定常枠・機械カレンダー等）を
    段階2と同様に読み、実績タイムラインのみ描画する。

    既存の出力ファイルがあり、メタ行の「データ抽出」表示が今回採用した
    ``データ抽出時間``（加工実績明細DATA 優先、無ければ計画ブックの
    「加工計画DATA_実績比較用」→「加工計画DATA」）の表示と
    一致する場合は、再生成をスキップしてそのファイルパスを返す。

    Returns:
        生成した（またはスキップ時は既存の）xlsx の絶対パス。

    Raises:
        PlanningValidationError: メンバー0人・表示対象日なし・実績イベント空など。
    """
    master_abs = _master_workbook_path_resolved()
    with _override_default_factory_hours_from_master(master_abs):
        global _MACHINE_CALENDAR_BLOCKS_BY_DATE
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
        (
            _skills_dict,
            members,
            equipment_list,
            _req_map,
            _need_rules,
            _surplus_map,
            _need_combo_col_index,
        ) = load_skills_and_needs()
        if not members:
            raise PlanningValidationError(
                "実績明細ガントのみ更新を中断しました: メンバーが0人です（マスタ skills を確認してください）。"
            )
        try:
            _MACHINE_CALENDAR_BLOCKS_BY_DATE = load_machine_calendar_occupancy_blocks(
                master_abs,
                equipment_list,
            )
        except Exception as e:
            logging.warning(
                "機械カレンダー: 読込例外のため占有なしとして続行します (%s)", e
            )
            _MACHINE_CALENDAR_BLOCKS_BY_DATE = {}
        try:
            (
                _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE,
                _STAGE2_MACHINE_DAILY_STARTUP_REQ_BY_MACHINE,
            ) = load_machine_daily_startup_settings(master_abs)
        except Exception as e:
            logging.warning(
                "機械日次始業準備設定: 読込例外のため無視します (%s)", e
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
            ) = load_request_switch_prep_settings(master_abs)
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
        if any(int(v or 0) > 0 for v in _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE.values()):
            _a12r, _a12r2 = _read_master_main_factory_operating_times(master_abs)
            if _a12r is None or _a12r2 is None:
                raise PlanningValidationError(
                    "日次始業準備に有効な準備時間（分）が登録されているため、"
                    "master.xlsm メインの A12・B12（工場稼働開始・終了）を正しく設定してください。"
                )
        try:
            _rs_a15, _ = _read_master_main_regular_shift_times(master_abs)
            _STAGE2_REGULAR_SHIFT_START = _rs_a15
        except Exception as e:
            logging.warning(
                "定常開始(A15) 読込失敗: 結果の定常外着色等で参照しません (%s)", e
            )
            _STAGE2_REGULAR_SHIFT_START = None

        # 実績明細ガントの「データ抽出」は、加工実績明細DATAのデータ抽出時間を最優先する。
        # 無いときの計画ブック側は「加工計画DATA_実績比較用」を優先し、得られなければ加工計画DATA。
        # 勤怠の当日判定などの実行基準もこの抽出時刻と揃える。
        data_extract_dt, plan_base_dt_column = (
            _extract_data_extraction_datetime_for_actual_related_gantt()
        )
        _STAGE2_DATA_EXTRACTION_DATETIME = data_extract_dt
        base_now_dt = data_extract_dt if data_extract_dt is not None else datetime.now()
        run_date = base_now_dt.date()
        data_extract_dt_str = (
            base_now_dt.strftime("%Y/%m/%d %H:%M:%S")
            if data_extract_dt is not None
            else "—"
        )

        attendance_data, ai_log_data = load_attendance_and_analyze(members)
        global_priority_raw = load_main_sheet_global_priority_override_text()
        global_priority_override = analyze_global_priority_override_comment(
            global_priority_raw, members, run_date.year, ai_sheet_sink=ai_log_data
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

        sorted_dates = sorted(list(attendance_data.keys()))
        sorted_dates = [d for d in sorted_dates if d >= run_date]
        if not sorted_dates:
            raise PlanningValidationError(
                "実績明細ガントのみ更新を中断しました: 当日以降の処理対象日付がありません。"
            )

        _reg_shift_start, _reg_shift_end = _read_master_main_regular_shift_times(
            master_abs
        )

        df_actual_detail = load_machining_actual_detail_df()

        def _first_valid_dt_from_df_col(_df, _col) -> datetime | None:
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

        # 加工実績明細DATA の「データ抽出時間」を最優先（無い/空なら計画ブックは実績比較用→加工計画DATA）
        detail_extract_dt = _first_valid_dt_from_df_col(
            df_actual_detail, TASK_COL_DATA_EXTRACTION_TIME
        )
        if detail_extract_dt is not None:
            base_now_dt = detail_extract_dt
            run_date = base_now_dt.date()
            data_extract_dt_str = base_now_dt.strftime("%Y/%m/%d %H:%M:%S")
            plan_base_dt_column = (
                f"{_actual_detail_sheet_log_label()}:{TASK_COL_DATA_EXTRACTION_TIME}"
            )

        logging.info(
            "実績明細ガントのみ: 抽出基準日時 %s（%s）",
            base_now_dt.strftime("%Y/%m/%d %H:%M:%S"),
            plan_base_dt_column if data_extract_dt is not None else "現在時刻フォールバック",
        )

        out_path = os.path.join(output_dir, ACTUAL_DETAIL_GANTT_REFRESH_FILENAME)
        prev_extract_display = _read_existing_equipment_gantt_data_extract_display(
            out_path, RESULT_SHEET_GANTT_ACTUAL_DETAIL_NAME
        )
        cur_extract_display = (data_extract_dt_str or "").strip()
        input_wb_mtime = None
        out_mtime = None
        try:
            _pip_wb = _excel_plan_input_wb()
            if _pip_wb and os.path.exists(_pip_wb):
                input_wb_mtime = os.path.getmtime(_pip_wb)
        except Exception:
            input_wb_mtime = None
        try:
            if out_path and os.path.exists(out_path):
                out_mtime = os.path.getmtime(out_path)
        except Exception:
            out_mtime = None
        mtime_ok = (
            input_wb_mtime is None
            or out_mtime is None
            or (out_mtime >= input_wb_mtime - 1e-6)
        )
        if (
            cur_extract_display
            and cur_extract_display not in ("—", "-")
            and prev_extract_display is not None
            and prev_extract_display.strip() == cur_extract_display
            and mtime_ok
        ):
            logging.info(
                "実績明細ガントのみ: データ抽出時間が前回出力と同一のためファイル更新をスキップしました（%s）。",
                cur_extract_display,
            )
            return os.path.abspath(out_path)

        detail_timeline_events: list = []
        sorted_dates_detail = list(sorted_dates)
        chart_title_actual_detail = "湖南工場 加工実績（明細）"
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
            detail_timeline_events = build_actual_timeline_events(
                df_actual_detail,
                equipment_list,
                sorted_dates_detail,
                log_sheet_name=_actual_detail_sheet_log_label(),
                roll_detail=True,
            )

        if not detail_timeline_events:
            raise PlanningValidationError(
                "実績明細ガントを生成できるイベントがありません。"
                "「加工実績明細DATA」の有無・日付・必須列を確認してください。"
            )

        _try_remove_path_with_retries(out_path)

        gantt_detail_tl_label_specs: list = []
        gantt_detail_tl_day_blocks: list = []
        with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
            gantt_detail_tl_label_specs, gantt_detail_tl_day_blocks = (
                _write_results_equipment_gantt_sheet(
                    writer,
                    [],
                    equipment_list,
                    sorted_dates_detail,
                    attendance_data,
                    data_extract_dt_str,
                    base_now_dt,
                    actual_timeline_events=detail_timeline_events,
                    regular_shift_times=(_reg_shift_start, _reg_shift_end),
                    plan_rows=False,
                    chart_title=chart_title_actual_detail,
                    sheet_name_override=RESULT_SHEET_GANTT_ACTUAL_DETAIL_NAME,
                )
            )
            wb = writer.book
            for _sn in list(wb.sheetnames):
                if _sn != RESULT_SHEET_GANTT_ACTUAL_DETAIL_NAME:
                    wb.remove(wb[_sn])

        _stage2_try_add_gantt_timeline_shape_labels(
            out_path,
            gantt_detail_tl_label_specs,
            gantt_detail_tl_day_blocks,
            sheet_name=RESULT_SHEET_GANTT_ACTUAL_DETAIL_NAME,
        )
        logging.info(
            "実績明細ガントのみ: %s を出力しました。",
            os.path.basename(out_path),
        )
        return os.path.abspath(out_path)
def _compare_csv_pick_column(df, candidates: tuple[str, ...]) -> str | None:
    """結果_タスク一覧.csv の列名揺れに耐える。"""
    if df is None or not len(getattr(df, "columns", [])):
        return None
    headers = {str(c).strip(): c for c in df.columns}
    for cand in candidates:
        if cand in headers:
            return str(headers[cand])
    low = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        k = cand.strip().lower()
        if k in low:
            return str(low[k])
    return None
def _timeline_events_force_machine_display_name(events: list | None) -> None:
    """イベント machine を機械名のみに正規化（ガント行キーと一致させる）。"""
    if not events:
        return
    for ev in events:
        mk = str(ev.get("machine") or "").strip()
        if not mk:
            continue
        _, mn = _split_equipment_line_process_machine(mk)
        if mn:
            ev["machine"] = mn
        else:
            ev["machine"] = mk
def _compare_gantt_unique_machine_row_order(
    full_equipment_list: list, machine_names: set[str]
) -> list[str]:
    """マスタ設備列の出現順を優先し、残りをソートして機械名の一覧を返す。"""
    seen: set[str] = set()
    out: list[str] = []
    for eq in full_equipment_list or []:
        _, mn = _split_equipment_line_process_machine(str(eq))
        k = (mn or str(eq).strip()).strip()
        if not k or k in seen or k not in machine_names:
            continue
        seen.add(k)
        out.append(k)
    for k in sorted(machine_names - seen):
        if k:
            out.append(k)
    return out
def _compare_gantt_find_first_plan_overlap_message(
    events: list | None, log_label: str
) -> str | None:
    """同一機械・同一暦日で半開区間が重なる最初の組のメッセージを返す。なければ None。"""
    if not events:
        return None
    buck: dict[tuple[str, date], list[tuple[datetime, datetime, str]]] = defaultdict(list)
    for ev in events:
        if not _is_machining_timeline_event(ev):
            continue
        m = str(ev.get("machine") or "").strip()
        d0 = ev.get("date")
        st = ev.get("start_dt")
        ed = ev.get("end_dt")
        if not m or not isinstance(d0, date):
            continue
        if not isinstance(st, datetime) or not isinstance(ed, datetime) or st >= ed:
            continue
        tid = str(ev.get("task_id") or "").strip()
        buck[(m, d0)].append((st, ed, tid))
    for (m, d0), segs in buck.items():
        if len(segs) < 2:
            continue
        segs.sort(key=lambda x: (x[0], x[1], x[2]))
        for i in range(len(segs)):
            a0, a1, ta = segs[i]
            for j in range(i + 1, len(segs)):
                b0, b1, tb = segs[j]
                if max(a0, b0) < min(a1, b1):
                    return (
                        f"{log_label}: 同一設備のタイムライン区間が重なります。"
                        f" 機械={m} 日={d0.isoformat()} 依頼NO={ta!r} vs {tb!r} "
                        f"[{a0}..{a1}) [{b0}..{b1})"
                    )
    return None
def _compare_gantt_assert_no_overlap(events: list | None, log_label: str) -> None:
    """同一機械・同一暦日で半開区間が重なる組があるとき例外（配台不整合の検知）。

    環境変数 ``COMPARE_GANTT_ALLOW_PLAN_OVERLAP`` が 1/true/yes/on のときは
    ``PlanningValidationError`` にせず ``logging.warning`` のみ（比較ガントは続行）。
    """
    msg = _compare_gantt_find_first_plan_overlap_message(events, log_label)
    if not msg:
        return
    allow = os.environ.get(ENV_COMPARE_GANTT_ALLOW_PLAN_OVERLAP, "").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )
    if allow:
        logging.warning("%s が有効のため続行します: %s", ENV_COMPARE_GANTT_ALLOW_PLAN_OVERLAP, msg)
        print(msg, file=sys.stderr)
        return
    raise PlanningValidationError(msg)
def _build_plan_timeline_events_from_snapshot_result_task_csv(csv_path: str) -> list:
    """スナップショットの結果_タスク一覧.csv から計画タイムラインイベントを構築する。"""
    if not csv_path or not os.path.isfile(csv_path):
        raise PlanningValidationError(
            f"計画実績比較ガント: CSV が見つかりません: {csv_path!r}"
        )
    df = None
    for enc in ("utf-8-sig", "utf-8", "cp932"):
        try:
            df = pd.read_csv(csv_path, encoding=enc)
            break
        except Exception:
            df = None
    if df is None or len(df) == 0:
        raise PlanningValidationError(
            f"計画実績比較ガント: CSV を読み込めないか空です: {csv_path!r}"
        )
    col_tid = _compare_csv_pick_column(
        df, ("タスクID", "依頼NO", TASK_COL_TASK_ID)
    )
    col_mach = _compare_csv_pick_column(df, ("機械名", TASK_COL_MACHINE_NAME))
    col_proc = _compare_csv_pick_column(df, ("工程名", TASK_COL_MACHINE))
    col_s = _compare_csv_pick_column(
        df,
        (
            "配台済_加工開始",
            "配完_加工開始",
        ),
    )
    col_e = _compare_csv_pick_column(
        df,
        (
            "配台済_加工終了",
            "配完_加工終了",
        ),
    )
    col_qty_len = _compare_csv_pick_column(
        df,
        (
            "残加工量",
            "残り加工量",
            TASK_COL_QTY,
            "換算数量",
        ),
    )
    if not col_tid or not col_mach or not col_s or not col_e:
        raise PlanningValidationError(
            "計画実績比較ガント: 結果_タスク一覧.csv に必要列"
            "（タスクID/依頼NO・機械名・配台済_加工開始・配台済_加工終了 等）が不足しています。"
        )
    events: list = []
    for _, row in df.iterrows():
        tid = row.get(col_tid)
        if tid is None or (isinstance(tid, float) and pd.isna(tid)):
            continue
        tid_s = str(tid).strip()
        if not tid_s:
            continue
        mv = row.get(col_mach)
        if mv is None or (isinstance(mv, float) and pd.isna(mv)):
            continue
        mach_s = str(mv).strip()
        if not mach_s:
            continue
        pv = row.get(col_proc) if col_proc else None
        proc_s = (
            ""
            if pv is None or (isinstance(pv, float) and pd.isna(pv))
            else str(pv).strip()
        )
        s_raw = row.get(col_s)
        e_raw = row.get(col_e)
        if s_raw is None or e_raw is None:
            continue
        s_dt = pd.to_datetime(s_raw, errors="coerce")
        e_dt = pd.to_datetime(e_raw, errors="coerce")
        if pd.isna(s_dt) or pd.isna(e_dt):
            continue
        if isinstance(s_dt, pd.Timestamp):
            s_dt = s_dt.to_pydatetime()
        if isinstance(e_dt, pd.Timestamp):
            e_dt = e_dt.to_pydatetime()
        if not isinstance(s_dt, datetime) or not isinstance(e_dt, datetime):
            continue
        if s_dt >= e_dt:
            continue
        d0 = s_dt.date()
        day_start = datetime.combine(d0, DEFAULT_START_TIME)
        day_end = datetime.combine(d0, DEFAULT_END_TIME)
        if e_dt <= day_start or s_dt >= day_end:
            continue
        s_clip = max(s_dt, day_start)
        e_clip = min(e_dt, day_end)
        if s_clip >= e_clip:
            continue
        _lm_snap = None
        if col_qty_len:
            try:
                _qv = row.get(col_qty_len)
                if _qv is not None and not (isinstance(_qv, float) and pd.isna(_qv)):
                    _lm_snap = float(_qv)
                    if (
                        math.isnan(_lm_snap)
                        or math.isinf(_lm_snap)
                        or abs(_lm_snap) < 1e-12
                    ):
                        _lm_snap = None
            except (TypeError, ValueError):
                _lm_snap = None
        events.append(
            {
                "date": d0,
                "task_id": tid_s,
                "machine": mach_s,
                "op": "",
                "sub": proc_s,
                "start_dt": s_clip,
                "end_dt": e_clip,
                "breaks": [],
                "units_done": 0,
                "already_done_units": 0,
                "total_units": 0,
                "eff_time_per_unit": 0.0,
                "unit_m": 0.0,
                # スナップショット CSV に長さ(m)列があればシェイプラベル「依頼NO ○○m」に反映する
                "label_len_m": _lm_snap,
                "event_kind": TIMELINE_EVENT_MACHINING,
            }
        )
    if not events:
        raise PlanningValidationError(
            "計画実績比較ガント: CSV から有効な計画区間を1件も構築できませんでした。"
        )
    return events
_COMPARE_GANTT_ALADDIN_QTY_COL_RE = re.compile(
    r"^\s*(\d{4})[./-](\d{1,2})[./-](\d{1,2})(?:_加工数量)?(?:__\d+)?\s*$"
)
def _try_read_plan_tasks_sheet_for_compare_aladdin():
    """PM_AI_PLAN_INPUT_PATH の計画タスクシートを読む（アラジン日次数量列用）。失敗時 None。

    ``ENV_COMPARE_GANTT_PLAN_TASKS_SHEET`` が空でないときはそのシート名、空のときは
    ``TASKS_SHEET_NAME``（加工計画DATA）。配台基準日時の取得は別経路のまま。
    """
    wb = _excel_plan_input_wb()
    if not wb or not os.path.isfile(wb):
        logging.warning(
            "計画実績比較ガント: PM_AI_PLAN_INPUT_PATH がブック実ファイルでないため、アラジン入力数量の行は空表示になります。"
        )
        return None
    sheet_name = (os.environ.get(ENV_COMPARE_GANTT_PLAN_TASKS_SHEET, "") or "").strip()
    if not sheet_name:
        sheet_name = TASKS_SHEET_NAME
    try:
        df = pd.read_excel(wb, sheet_name=sheet_name)
        df.columns = df.columns.astype(str).str.strip()
        df = _align_dataframe_headers_to_canonical(df, list(SOURCE_BASE_COLUMNS))
        logging.info(
            "計画実績比較ガント: アラジン参照用に '%s' の『%s』を読み込みました。",
            os.path.basename(wb),
            sheet_name,
        )
        return df
    except Exception as e:
        logging.warning(
            "計画実績比較ガント: シート『%s』の読込に失敗したためアラジン行は空です（%s）。",
            sheet_name,
            e,
        )
        return None
def _parse_optional_float_non_nan(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None
def _format_qty_short(q: float) -> str:
    if abs(q - round(q)) < 1e-9:
        return str(int(round(q)))
    s = f"{q:.4f}".rstrip("0").rstrip(".")
    return s if s else "0"
_COMPARE_GANTT_TRACE_BASE_TIDS = frozenset({"W3-12", "A4-3", "W4-12"})
def _compare_gantt_trace_should_log_btid(btid: str) -> bool:
    n = planning_task_id_str_from_scalar(btid) or (btid or "").strip()
    if not n:
        return False
    return n.upper() in {x.upper() for x in _COMPARE_GANTT_TRACE_BASE_TIDS}
def _aggregate_actual_qty_for_aladdin_compare_from_detail_df(
    df: pd.DataFrame | None,
    equipment_list,
    sorted_dates: list,
) -> dict[tuple[str, date], dict[str, float]]:
    """
    アラジン「日付_加工数量」と突き合わせる実績数量（m）を、加工実績明細DFから集計する。
    タイムラインイベントを足し上げない（同一内容の複数行で二重計上しないよう、
    依頼NO基底・工程・開始終了（分単位に丸め）・実加工/累積が一致する行は先に1行分として扱う）。
    量の按分は build_actual_timeline_events と同様、実加工数を優先し無いとき累積を時間比按分する。
    機械キーは ``_timeline_events_force_machine_display_name`` と同様に
    設備列の「+」より右（機械名のみ）へ寄せ、アラジン側の TASK_COL_MACHINE_NAME キーと一致させる。
    同一機械×日×依頼NOに **多数のほぼ同一按分** が付くエクスポート二重（件数と按分のばらつ閾値）のときは
    1 件分に畳む（ロール等で少数行のみ同額のときは閾値で誤畳みを避ける）。
    """
    if df is None or len(df) == 0:
        return {}

    equip_lookup = _equipment_lookup_normalized_to_canonical(equipment_list)
    date_ok = set(sorted_dates)
    # (機械キー, 日, 依頼NO) → 行ごとの按分値のリスト（後で同一値大量時に畳む）
    _per_mdt: dict[tuple[str, date, str], list[float]] = defaultdict(list)
    _trace_per_mdt: dict[tuple[str, date, str], list[float]] = defaultdict(list)
    seen_sig: set[tuple] = set()

    def _scalar_for_dedupe(v) -> float | str:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "__na__"
        try:
            return round(float(v), 4)
        except (TypeError, ValueError):
            return "__na__"

    # 同一機械×日×依頼NOに「ほぼ同じ按分」が複数行付く重複エクスポート対策。
    # 同一按分が 3 件のみの日もあり、閾値を下げないと sum のまま過大になることがある。
    # 閾値 3 は「同一日・同依頼で等分割された正規の 3 セグメント」が偶然同じ m に揃うと誤畳みのリスクあり（稀）。
    _ALADDIN_DUP_COLLAPSE_MIN_SAME = 3
    _ALADDIN_DUP_SPREAD_TOL_M = 0.05  # 同一按分とみなす m 幅（表示単位の揺れ吸収）

    def _sig_time_minute(dt: datetime) -> str:
        """重複エクスポートで秒以下だけ異なる行を同一シグネチャに寄せる。"""
        if not isinstance(dt, datetime):
            return ""
        return dt.replace(second=0, microsecond=0).strftime("%Y-%m-%d %H:%M")

    for _, row in df.iterrows():
        tid = row.get(ACT_COL_TASK_ID)
        if tid is None or pd.isna(tid):
            continue
        tid_s = str(tid).strip()
        if not tid_s:
            continue
        btid = planning_task_id_str_from_scalar(row.get(ACT_COL_TASK_ID)) or ""
        if not btid:
            continue
        proc = row.get(ACT_COL_PROCESS)
        if proc is None or pd.isna(proc):
            continue
        proc_key = _normalize_equipment_match_key(proc)
        mach = equip_lookup.get(proc_key)
        if not mach:
            continue
        mach_raw = str(mach or "").strip()
        _, mn_part = _split_equipment_line_process_machine(mach_raw)
        mach_display = (mn_part or mach_raw).strip()
        mk = _normalize_equipment_match_key(mach_display)
        if not mk:
            continue
        start_dt, end_dt = _actual_row_time_bounds(row)
        if not start_dt or not end_dt or start_dt >= end_dt:
            continue
        try:
            actual_done_m = parse_float_safe(row.get(ACT_COL_ACTUAL_QTY), None)
        except Exception:
            actual_done_m = None
        try:
            cumulative_actual_m = parse_float_safe(
                row.get(ACT_COL_CUMULATIVE_ACTUAL_QTY), None
            )
        except Exception:
            cumulative_actual_m = None
        try:
            total_seconds = max(0.0, float((end_dt - start_dt).total_seconds()))
        except Exception:
            total_seconds = None
        sig = (
            btid,
            proc_key,
            _sig_time_minute(start_dt),
            _sig_time_minute(end_dt),
            _scalar_for_dedupe(actual_done_m),
            _scalar_for_dedupe(cumulative_actual_m),
        )
        if sig in seen_sig:
            continue
        seen_sig.add(sig)

        for d in sorted_dates:
            if d not in date_ok:
                continue
            day_start = datetime.combine(d, DEFAULT_START_TIME)
            day_end = datetime.combine(d, DEFAULT_END_TIME)
            if end_dt <= day_start or start_dt >= day_end:
                continue
            s_clip = max(start_dt, day_start)
            e_clip = min(end_dt, day_end)
            if s_clip >= e_clip:
                continue
            seg_seconds = float((e_clip - s_clip).total_seconds())
            qty = None
            try:
                if (
                    actual_done_m is not None
                    and isinstance(actual_done_m, (int, float))
                    and not (isinstance(actual_done_m, float) and pd.isna(actual_done_m))
                    and float(actual_done_m) > 1e-12
                    and total_seconds
                    and float(total_seconds) > 1e-9
                    and seg_seconds > 1e-9
                ):
                    qty = float(actual_done_m) * (
                        seg_seconds / float(total_seconds)
                    )
                elif (
                    cumulative_actual_m is not None
                    and isinstance(cumulative_actual_m, (int, float))
                    and not (
                        isinstance(cumulative_actual_m, float)
                        and pd.isna(cumulative_actual_m)
                    )
                    and float(cumulative_actual_m) > 1e-12
                    and total_seconds
                    and float(total_seconds) > 1e-9
                    and seg_seconds > 1e-9
                ):
                    qty = float(cumulative_actual_m) * (
                        seg_seconds / float(total_seconds)
                    )
            except Exception:
                qty = None
            if qty is not None and qty > 1e-12:
                _per_mdt[(mk, d, btid)].append(float(qty))
                if _compare_gantt_trace_should_log_btid(btid):
                    _trace_per_mdt[(mk, d, btid)].append(float(qty))

    tmp: dict[tuple[str, date], dict[str, float]] = defaultdict(
        lambda: defaultdict(float)
    )
    for (_mk, _d, _tid), _vals in _per_mdt.items():
        if not _vals:
            continue
        _vals_f = [float(x) for x in _vals]
        _vals_r = [round(x, 4) for x in _vals_f]
        _spread_raw = max(_vals_f) - min(_vals_f)
        _spread_r = (max(_vals_r) - min(_vals_r)) if _vals_r else 0.0
        # 按分の浮動小数ノイズで spread が閾値をわずかに超えて畳み込み失敗するのを防ぐ（dedupe と同じ 4 桁）
        _spread = float(_spread_r)
        if (
            len(_vals) >= _ALADDIN_DUP_COLLAPSE_MIN_SAME
            and _spread <= _ALADDIN_DUP_SPREAD_TOL_M
        ):
            merged = float(max(_vals_r))
        else:
            merged = float(sum(_vals_f))
        tmp[_mk, _d][_tid] = merged

    for _tk, _tvals in sorted(
        _trace_per_mdt.items(),
        key=lambda kv: (kv[0][1], kv[0][0], kv[0][2]),
    ):
        _tmk, _td, _ttid = _tk
        _merged = tmp.get((_tmk, _td), {}).get(_ttid)
        _n = len(_tvals)
        _tf = [float(x) for x in _tvals]
        _tr = [round(x, 4) for x in _tf]
        _mm_r = (max(_tr) - min(_tr)) if _tr else 0.0
        _collapsed = (
            _n >= _ALADDIN_DUP_COLLAPSE_MIN_SAME and _mm_r <= _ALADDIN_DUP_SPREAD_TOL_M
        )
        _preview = _tvals[:15]
        _suffix = " …" if _n > 15 else ""
        logging.info(
            "計画実績比較ガント[トレース] 明細→日次按分 依頼NO=%s 日=%s 機械キー=%s "
            "按分値件数=%s 畳み込み適用=%s minmax差(m,4桁丸め)=%s 閾値(件>=%s 幅<=%s) 按分値先頭=%s%s 最終実績(m)=%s",
            _ttid,
            _td.isoformat() if isinstance(_td, date) else str(_td),
            _log_plain_label(_tmk),
            _n,
            _collapsed,
            _mm_r,
            _ALADDIN_DUP_COLLAPSE_MIN_SAME,
            _ALADDIN_DUP_SPREAD_TOL_M,
            _preview,
            _suffix,
            _merged,
        )

    return {k: dict(v) for k, v in tmp.items()}
def _compare_aladdin_plan_buckets_vs_actual(
    buckets: dict[tuple[str, date], list[tuple[str, float]]],
    actual_agg: dict[tuple[str, date], dict[str, float]],
) -> dict[tuple[str, date], str]:
    """
    アラジン計画（バケツ内の依頼NOごとの数量）と実績集計を比較し、
    機械×日ごとの不一致説明文（空なら一致）を返す。
    """
    notes: dict[tuple[str, date], str] = {}
    for key, parts in buckets.items():
        plan_by_tid: dict[str, float] = defaultdict(float)
        for t, q in parts:
            tid = planning_task_id_str_from_scalar(t) or (
                str(t).strip() if t else ""
            )
            tid = tid or "—"
            plan_by_tid[tid] += float(q)
        act_map = actual_agg.get(key) or {}
        pieces: list[str] = []
        for tid in sorted(plan_by_tid.keys()):
            pq = float(plan_by_tid[tid])
            aq = float(act_map.get(tid, 0.0))
            if _compare_gantt_trace_should_log_btid(tid):
                _k_mk, _k_dt = key
                logging.info(
                    "計画実績比較ガント[トレース] アラジン比較 機械キー=%s 日=%s 依頼NO=%s "
                    "計画(m)=%s 実績(m)=%s isclose=%s",
                    _log_plain_label(_k_mk),
                    _k_dt.isoformat() if isinstance(_k_dt, date) else str(_k_dt),
                    tid,
                    _format_qty_short(pq),
                    _format_qty_short(aq),
                    math.isclose(pq, aq, rel_tol=1e-9, abs_tol=1e-2),
                )
            if not math.isclose(pq, aq, rel_tol=1e-9, abs_tol=1e-2):
                pieces.append(
                    f"{tid} 計画{_format_qty_short(pq)}≠実績{_format_qty_short(aq)}"
                )
        if pieces:
            notes[key] = "【実績不一致】" + " ".join(pieces)
    nm = len(notes)
    if nm:
        logging.info(
            "計画実績比較ガント: アラジン計画と実績数量が異なる機械×日が %s 件あります。",
            nm,
        )
    return notes
def _merge_aladdin_lookup_with_actual_mismatch_annotations(
    lookup: dict[tuple[str, date], tuple[str, str]],
    buckets: dict[tuple[str, date], list[tuple[str, float]]],
    actual_agg: dict[tuple[str, date], dict[str, float]],
) -> dict[tuple[str, date], tuple[str, str, str]]:
    """表示用 lookup に第3要素（不一致時の注記、一致時は空文字）を付与する。"""
    mismatch_by_key = _compare_aladdin_plan_buckets_vs_actual(buckets, actual_agg)
    out: dict[tuple[str, date], tuple[str, str, str]] = {}
    for k, pair in lookup.items():
        note = (mismatch_by_key.get(k) or "").strip()
        out[k] = (pair[0], pair[1], note)
    return out
def _build_compare_gantt_aladdin_qty_lookup(
    df: pd.DataFrame | None, dates_set: set
) -> tuple[
    dict[tuple[str, date], tuple[str, str]],
    dict[tuple[str, date], list[tuple[str, float]]],
]:
    """
    加工計画DATA の日付付き数量列から、(機械名キー, 日付) ごとの
    （タスク概覝＝依頼NOのみ, タイムライン中央＝「依頼NO(数量)」をスペース区切り）を構築する。
    列名は ``YYYY/MM/DD_加工数量`` または工程別成形後の ``YYYY/MM/DD``（同一セグメントが加工数量）。
    併せて同一キーの生バケツ（依頼NO別数量の比較用）も返す。
    """
    out: dict[tuple[str, date], tuple[str, str]] = {}
    if df is None or len(df) == 0:
        return out, {}
    if TASK_COL_MACHINE_NAME not in df.columns:
        logging.warning(
            "計画実績比較ガント: 列『%s』が無いためアラジン数量は結合できません。",
            TASK_COL_MACHINE_NAME,
        )
        return out, {}

    buckets: dict[tuple[str, date], list[tuple[str, float]]] = defaultdict(list)
    date_cols: list[tuple[str, date]] = []
    for col in df.columns:
        # Excel 見出しの全角「＿」・数字などが混ざると素の str(col) では正規表現に届かない。
        col_key = _nfkc_column_aliases(col)
        m = _COMPARE_GANTT_ALADDIN_QTY_COL_RE.match(col_key)
        if not m:
            continue
        try:
            y, mo, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
            dk = date(y, mo, dd)
        except ValueError:
            continue
        date_cols.append((str(col).strip(), dk))

    if not date_cols:
        logging.info(
            "計画実績比較ガント: 「YYYY/MM/DD」または「YYYY/MM/DD_加工数量」形式の数量列がありません（アラジン行は空または予約のみ）。"
        )

    for col_name, dk in date_cols:
        if dk not in dates_set:
            continue
        if col_name not in df.columns:
            continue
        for _, row in df.iterrows():
            mach_k = _normalize_equipment_match_key(row.get(TASK_COL_MACHINE_NAME))
            if not mach_k:
                continue
            qty = _parse_optional_float_non_nan(row.get(col_name))
            if qty is None or abs(qty) < 1e-12:
                continue
            tid = planning_task_id_str_from_scalar(row.get(TASK_COL_TASK_ID))
            buckets[mach_k, dk].append((tid or "—", qty))

    for key, parts in buckets.items():
        parts_show = parts[:15]
        _seen_tid: set[str] = set()
        _uids: list[str] = []
        for t, _q in parts:
            tt = (t or "—").strip() or "—"
            if tt not in _seen_tid:
                _seen_tid.add(tt)
                _uids.append(tt)
        detail = " ".join(_uids[:15])
        if len(parts) > 15 or len(_uids) > 15:
            detail += " …"
        center = " ".join(
            f"{(t or '—')}({_format_qty_short(q)})" for t, q in parts_show
        )
        if len(parts) > 15:
            center += " …"
        out[key] = (detail, center)

    return out, dict(buckets)
def write_plan_actual_compare_gantt_from_snapshot_dir(snapshot_dir: str) -> str:
    """
    過去スナップショット（``結果_タスク一覧.csv``）の計画と、現在マスタの加工実績明細を
    同一シートで比較する設備ガント（計画／実績／アラジン入力数量の3段）を
    ``output_dir`` / ``COMPARE_GANTT_OUTPUT_FILENAME`` に出力する。

    Args:
        snapshot_dir: 例 ``.../pdf/20260422_153045``（直下に ``結果_タスク一覧.csv`` があること）。

    Returns:
        出力 xlsx の絶対パス。
    """
    snap = os.path.abspath(os.path.normpath((snapshot_dir or "").strip()))
    if not snap or not os.path.isdir(snap):
        raise PlanningValidationError(
            f"計画実績比較ガント: スナップショットフォルダが無効です: {snapshot_dir!r}"
        )
    csv_path = os.path.join(snap, f"{RESULT_TASK_SHEET_NAME}.csv")
    plan_events = _build_plan_timeline_events_from_snapshot_result_task_csv(csv_path)
    _compare_gantt_assert_no_overlap(plan_events, "計画(CSV)")

    master_abs = _master_workbook_path_resolved()
    with _override_default_factory_hours_from_master(master_abs):
        global _MACHINE_CALENDAR_BLOCKS_BY_DATE
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
        (
            _skills_dict,
            members,
            equipment_list,
            _req_map,
            _need_rules,
            _surplus_map,
            _need_combo_col_index,
        ) = load_skills_and_needs()
        if not members:
            raise PlanningValidationError(
                "計画実績比較ガント: メンバーが0人です（マスタ skills を確認してください）。"
            )
        try:
            _MACHINE_CALENDAR_BLOCKS_BY_DATE = load_machine_calendar_occupancy_blocks(
                master_abs,
                equipment_list,
            )
        except Exception as e:
            logging.warning(
                "計画実績比較ガント: 機械カレンダー読込例外のため占有なし (%s)", e
            )
            _MACHINE_CALENDAR_BLOCKS_BY_DATE = {}
        try:
            (
                _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE,
                _STAGE2_MACHINE_DAILY_STARTUP_REQ_BY_MACHINE,
            ) = load_machine_daily_startup_settings(master_abs)
        except Exception as e:
            logging.warning(
                "計画実績比較ガント: 機械日次始業準備設定読込例外 (%s)", e
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
            ) = load_request_switch_prep_settings(master_abs)
        except Exception as e:
            logging.warning(
                "計画実績比較ガント: 依頼切替準備設定読込例外 (%s)", e
            )
            _STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE = {}
            _STAGE2_REQUEST_SWITCH_PREP_BY_MACHINE = {}
            _STAGE2_BREAK_RESUME_PREP_BY_PROC_MACHINE = {}
            _STAGE2_BREAK_RESUME_PREP_BY_MACHINE = {}
            _STAGE2_POST_MACHINING_CLEANUP_BY_PROC_MACHINE = {}
            _STAGE2_POST_MACHINING_CLEANUP_BY_MACHINE = {}
            _STAGE2_REQUEST_INTERVAL_BUFFER_BY_PROC_MACHINE = {}
            _STAGE2_REQUEST_INTERVAL_BUFFER_BY_MACHINE = {}
        if any(
            int(v or 0) > 0
            for v in _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE.values()
        ):
            _a12r, _a12r2 = _read_master_main_factory_operating_times(master_abs)
            if _a12r is None or _a12r2 is None:
                raise PlanningValidationError(
                    "計画実績比較ガント: 日次始業準備が有効なため master メイン A12・B12 を設定してください。"
                )
        try:
            _rs_a15, _ = _read_master_main_regular_shift_times(master_abs)
            _STAGE2_REGULAR_SHIFT_START = _rs_a15
        except Exception as e:
            logging.warning(
                "計画実績比較ガント: 定常開始(A15) 読込失敗 (%s)", e
            )
            _STAGE2_REGULAR_SHIFT_START = None

        data_extract_dt, _plan_base_dt_column = _extract_data_extraction_datetime()
        _STAGE2_DATA_EXTRACTION_DATETIME = data_extract_dt
        base_now_dt = data_extract_dt if data_extract_dt is not None else datetime.now()
        run_date = base_now_dt.date()
        data_extract_dt_str = (
            base_now_dt.strftime("%Y/%m/%d %H:%M:%S")
            if data_extract_dt is not None
            else "—"
        )

        attendance_data, ai_log_data = load_attendance_and_analyze(members)
        global_priority_raw = load_main_sheet_global_priority_override_text()
        global_priority_override = analyze_global_priority_override_comment(
            global_priority_raw, members, run_date.year, ai_sheet_sink=ai_log_data
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

        sorted_dates = sorted(list(attendance_data.keys()))
        sorted_dates = [d for d in sorted_dates if d >= run_date]
        if not sorted_dates:
            raise PlanningValidationError(
                "計画実績比較ガント: 当日以降の勤怠日がありません。"
            )

        _reg_shift_start, _reg_shift_end = _read_master_main_regular_shift_times(
            master_abs
        )

        df_actual_detail = load_machining_actual_detail_df()

        def _first_valid_dt_from_df_col2(_df, _col) -> datetime | None:
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

        detail_extract_dt = _first_valid_dt_from_df_col2(
            df_actual_detail, TASK_COL_DATA_EXTRACTION_TIME
        )
        if detail_extract_dt is not None:
            base_now_dt = detail_extract_dt
            run_date = base_now_dt.date()
            data_extract_dt_str = base_now_dt.strftime("%Y/%m/%d %H:%M:%S")

        sorted_dates_detail = list(sorted_dates)
        if df_actual_detail is not None and len(df_actual_detail) > 0:
            sorted_dates_detail = _sorted_dates_union_actual_bounds_df(
                sorted_dates, df_actual_detail
            )
            d_from = _parse_env_optional_date(ENV_GANTT_ACTUAL_DETAIL_DATE_FROM)
            d_to = _parse_env_optional_date(ENV_GANTT_ACTUAL_DETAIL_DATE_TO)
            if d_from is not None or d_to is not None:
                filtered_detail_dates = _sorted_dates_filter_inclusive_range(
                    sorted_dates_detail, d_from, d_to
                )
                if filtered_detail_dates:
                    sorted_dates_detail = filtered_detail_dates

        detail_timeline_events: list = []
        if df_actual_detail is not None and len(df_actual_detail) > 0:
            detail_timeline_events = build_actual_timeline_events(
                df_actual_detail,
                equipment_list,
                sorted_dates_detail,
                log_sheet_name=_actual_detail_sheet_log_label(),
                roll_detail=True,
            )
        if not detail_timeline_events:
            raise PlanningValidationError(
                "計画実績比較ガント: 実績明細からタイムラインイベントを生成できません。"
            )
        _compare_gantt_assert_no_overlap(detail_timeline_events, "実績(マスタ)")
        _timeline_events_force_machine_display_name(detail_timeline_events)

        mnames: set[str] = set()
        for ev in plan_events:
            m0 = str(ev.get("machine") or "").strip()
            if m0:
                mnames.add(m0)
        for ev in detail_timeline_events:
            m0 = str(ev.get("machine") or "").strip()
            if m0:
                mnames.add(m0)
        equipment_list_compare = _compare_gantt_unique_machine_row_order(
            equipment_list, mnames
        )
        if not equipment_list_compare:
            raise PlanningValidationError(
                "計画実績比較ガント: 表示対象の機械名が0件です。"
            )

        plan_dates = {
            ev["date"] for ev in plan_events if isinstance(ev.get("date"), date)
        }
        act_dates = {
            ev["date"]
            for ev in detail_timeline_events
            if isinstance(ev.get("date"), date)
        }
        sorted_dates_show = sorted(
            d for d in (plan_dates | act_dates) if d in attendance_data
        )
        if not sorted_dates_show:
            raise PlanningValidationError(
                "計画実績比較ガント: 勤怠と交差する表示日がありません。"
            )

        _df_tasks_aladdin = _try_read_plan_tasks_sheet_for_compare_aladdin()
        _aladdin_lookup, _aladdin_buckets = _build_compare_gantt_aladdin_qty_lookup(
            _df_tasks_aladdin,
            set(sorted_dates_show),
        )
        _actual_tid_qty_agg = _aggregate_actual_qty_for_aladdin_compare_from_detail_df(
            df_actual_detail,
            equipment_list,
            sorted_dates_detail,
        )
        _aladdin_qty_lookup = _merge_aladdin_lookup_with_actual_mismatch_annotations(
            _aladdin_lookup,
            _aladdin_buckets,
            _actual_tid_qty_agg,
        )

        out_path = os.path.join(output_dir, COMPARE_GANTT_OUTPUT_FILENAME)
        _try_remove_path_with_retries(out_path)
        chart_title = (
            f"計画vs実績比較（スナップショット {os.path.basename(snap)}）"
        )
        gantt_tl_label_specs: list = []
        gantt_tl_day_blocks: list = []
        with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
            gantt_tl_label_specs, gantt_tl_day_blocks = _write_results_equipment_gantt_sheet(
                writer,
                plan_events,
                equipment_list_compare,
                sorted_dates_show,
                attendance_data,
                data_extract_dt_str,
                base_now_dt,
                actual_timeline_events=detail_timeline_events,
                regular_shift_times=(_reg_shift_start, _reg_shift_end),
                plan_rows=True,
                chart_title=chart_title,
                sheet_name_override=RESULT_SHEET_GANTT_COMPARE_NAME,
                gantt_compare_shape_styling=True,
                compare_aladdin_qty_by_machine_date=_aladdin_qty_lookup,
            )
            wb = writer.book
            for _sn in list(wb.sheetnames):
                if _sn != RESULT_SHEET_GANTT_COMPARE_NAME:
                    wb.remove(wb[_sn])

        _stage2_try_add_gantt_timeline_shape_labels(
            out_path,
            gantt_tl_label_specs,
            gantt_tl_day_blocks,
            sheet_name=RESULT_SHEET_GANTT_COMPARE_NAME,
        )
        _cmp_plan_sheet = (os.environ.get(ENV_COMPARE_GANTT_PLAN_TASKS_SHEET, "") or "").strip()
        if not _cmp_plan_sheet:
            _cmp_plan_sheet = TASKS_SHEET_NAME
        logging.info(
            "計画実績比較ガント: %s を出力しました（アラジン入力数量行の参照元＝%s の日付_加工数量列）。",
            os.path.basename(out_path),
            _cmp_plan_sheet,
        )
        return os.path.abspath(out_path)
def _stage2_publish_excel_enabled(stage2_output_root) -> bool:
    """
    既定では計画ブック・メンバー別スケジュールの xlsx を出力する。
    JavaFX から ``PM_AI_STAGE2_WRITE_EXCEL=0`` のときは成果物フォルダに xlsx を残さず JSON のみとする。
    試行順プローブ等（stage2_output_root あり）は従来どおり xlsx パスが必要なため常に出力する。
    """
    if stage2_output_root:
        return True
    v = (os.environ.get("PM_AI_STAGE2_WRITE_EXCEL") or "1").strip().lower()
    return v not in ("0", "false", "no", "off", "none")
def _stage2_write_excel_gantt_sheets_enabled() -> bool:
    """計画ブックに「結果_設備ガント」「結果_設備ガント_実績明細」を Excel シートとして含めるか。

    既定は False（設備ガント(グラフィック)が描画契約 JSON を参照するため Excel への重複出力を省略）。
    従来どおりブックへ書くときは ``PM_AI_STAGE2_WRITE_EXCEL_GANTT_SHEETS=1``。
    """
    v = (os.environ.get("PM_AI_STAGE2_WRITE_EXCEL_GANTT_SHEETS") or "0").strip().lower()
    return v in ("1", "true", "yes", "on")

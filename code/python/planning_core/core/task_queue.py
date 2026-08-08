# -*- coding: utf-8 -*-
# planning_core.core.task_queue — body only (loaded via _core exec chain)
def _limited_operator_error_context(excel_row_number, task_id) -> str:
    try:
        row_num = int(excel_row_number)
    except (TypeError, ValueError):
        row_num = 0
    tid = str(task_id or "").strip() or "（空欄）"
    return f"Excel行{row_num}・依頼NO={tid}・列「{PLAN_COL_LIMITED_OP}」"


def _parse_limited_operator_json_cell(
    raw_value, excel_row_number, task_id
) -> tuple[str, ...]:
    """「担当OP_限定」を厳格な JSON 文字列配列として読む。"""
    context = _limited_operator_error_context(excel_row_number, task_id)
    if raw_value is None or (isinstance(raw_value, float) and pd.isna(raw_value)):
        return ()
    if not isinstance(raw_value, str):
        raise PlanningValidationError(
            f"{context}: JSON文字列ではありません（実値の型={type(raw_value).__name__}）。"
        )
    raw = raw_value.strip()
    if not raw:
        return ()
    try:
        parsed = json.loads(raw)
    except (json.JSONDecodeError, TypeError) as exc:
        raise PlanningValidationError(
            f"{context}: 不正JSONです（{exc}）。"
        ) from exc
    if not isinstance(parsed, list):
        raise PlanningValidationError(f"{context}: JSONの最上位は配列でなければなりません。")
    if not parsed:
        raise PlanningValidationError(
            f"{context}: JSON配列には選択名を1名以上指定してください。"
        )

    names: list[str] = []
    normalized_seen: set[str] = set()
    for index, value in enumerate(parsed, start=1):
        if not isinstance(value, str):
            raise PlanningValidationError(
                f"{context}: 配列要素{index}は文字列でなければなりません。"
            )
        name = value.strip()
        if not name:
            raise PlanningValidationError(f"{context}: 配列要素{index}が空名です。")
        normalized = _normalize_person_name_for_match(name)
        if not normalized:
            raise PlanningValidationError(f"{context}: 配列要素{index}が空名です。")
        if normalized in normalized_seen:
            raise PlanningValidationError(
                f"{context}: 重複名があります（{name!r}）。"
            )
        normalized_seen.add(normalized)
        names.append(name)
    return tuple(names)


def _stage2_in_progress_next_day_dispatch_key(
    task_id: str, machine: str, machine_name: str
) -> str:
    tid = planning_task_id_str_from_scalar(task_id)
    return f"{tid}\x1e{str(machine or '').strip()}\x1e{str(machine_name or '').strip()}"
def build_task_queue_from_planning_df(
    tasks_df,
    run_date,
    req_map,
    ai_by_tid=None,
    global_priority_override=None,
    equipment_list=None,
):
    """
    ``generate_plan`` 内で呼みれる。完了済み・配台試行行を除し」残りを task_queue に穝む。

    「配台不要」列に「配台計画除外」が含まれる行は段階2の配台キューへ入れない（``PLAN_COL_STAGE2_DISPATCH_PLAN_EXCLUDE_MARKER``）。
    ai_by_tid は None のときの値内部で analyze_task_special_remarks を実行れる。
    """
    if ai_by_tid is None:
        ai_by_tid = analyze_task_special_remarks(tasks_df, reference_year=run_date.year)
    gpo = global_priority_override or {}
    in_progress_next_day_m = _load_stage2_in_progress_next_day_dispatch_overrides()
    aladdin_exclude_next_day_m = _load_stage2_aladdin_today_exclude_next_day_overrides()
    task_queue = []
    n_exclude_plan = 0
    seq_by_tid = _collect_process_content_order_by_task_id(tasks_df)
    same_tid_line_seq = defaultdict(int)
    # 依頼NO直列配台の順庝用: iterrows の読み込み順（0 始まり）。task_queue.sort 後も試行。
    planning_sheet_row_seq = 0
    _has_unprocessed_col = TASK_COL_UNPROCESSED in tasks_df.columns
    _plan_df_reset_effective_roll_unit_ilocs(tasks_df)

    for planning_df_iloc, (row_idx, row) in enumerate(tasks_df.iterrows()):
        task_id = planning_task_id_str_from_plan_row(row)
        planning_excel_row = planning_df_iloc + 2
        # 枝番タスク（入力3表）の親。列「元依頼NO」が空なら自身を親とする。
        # 配台 task_id は枝番依頼NO（例 Y3-24-01）のままで、特別ルール・WIP・工程依存は rule_task_id=親で集計する。
        parent_task_id = str(
            _planning_df_cell_scalar(row, PLAN_COL_PARENT_TASK_ID) or ""
        ).strip()
        if not parent_task_id:
            parent_task_id = task_id
        if row_has_completion_keyword(row):
            continue
        if _plan_row_exclude_as_completed_mikan_unprocessed_zero_actual_done_rule(row):
            continue
        if _plan_row_exclude_from_assignment(row):
            n_exclude_plan += 1
            continue
        if _plan_row_stage2_dispatch_plan_excluded(row):
            n_exclude_plan += 1
            continue

        machine = str(row.get(TASK_COL_MACHINE, "")).strip()
        machine_name = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
        _metrics_rem, done_qty, qty_total, _from_unprocessed_qty = (
            _plan_row_dispatch_qty_metrics(row)
        )
        dispatch_m, dispatch_rolls = _plan_row_stage2_dispatch_qty_and_rolls(row)
        # 加工速度: ②列「加工速度」（master.xlsm speed で基本速度×実稼働比率を反映）→
        # speed_ov は列「加工速度_上書き」のみ（①があれば上書き）。
        speed_raw = row.get(TASK_COL_SPEED, 1)
        product_name = row.get(TASK_COL_PRODUCT, None)
        answer_due = parse_optional_date(_planning_df_cell_scalar(row, TASK_COL_ANSWER_DUE))
        specified_due = parse_optional_date(_planning_df_cell_scalar(row, TASK_COL_SPECIFIED_DUE))
        specified_due_ov = None
        # 紝期基準: ①回答納期（空でなけれみ）②列「指定納期」（列「指定納期_上書き」は廃止済み）
        specified_basis = specified_due
        due_basis = None
        due_source = "none"
        due_source_rank = 9
        raw_input_sheet = parse_optional_date(
            _planning_df_cell_scalar(row, TASK_COL_RAW_INPUT_DATE)
        )
        raw_input_date = raw_input_sheet

        in_progress = done_qty > 0.0
        if in_progress and _stage2_truthy_env("PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH"):
            continue

        qty = max(0.0, float(dispatch_m))
        qty_from_in_progress_next_day_dialog = False
        aladdin_today_exclude_next_day_m = 0.0
        aladdin_today_exclude_next_day_dialog = False
        if in_progress and in_progress_next_day_m:
            ov_key = _stage2_in_progress_next_day_dispatch_key(
                task_id, machine, machine_name
            )
            if ov_key in in_progress_next_day_m:
                qty = _sanitize_dispatch_qty_m(float(in_progress_next_day_m[ov_key]))
                qty_from_in_progress_next_day_dialog = True
                logging.info(
                    "段階2: 加工途中の翌日配台量を適用 依頼NO=%s 工程=%s 機械名=%s → %s m（シート残量 %s m、1ロール固定）",
                    task_id,
                    _log_plain_label(machine),
                    _log_plain_label(machine_name),
                    qty,
                    dispatch_m,
                )

        if (
            not in_progress
            and aladdin_exclude_next_day_m
            and done_qty <= 1e-12
        ):
            ov_key = _stage2_in_progress_next_day_dispatch_key(
                task_id, machine, machine_name
            )
            if ov_key in aladdin_exclude_next_day_m:
                aladdin_today_exclude_next_day_m = _sanitize_dispatch_qty_m(
                    float(aladdin_exclude_next_day_m[ov_key])
                )
                aladdin_today_exclude_next_day_dialog = (
                    aladdin_today_exclude_next_day_m > 1e-12
                )
                if aladdin_today_exclude_next_day_dialog:
                    logging.info(
                        "段階2: アラジン当日・翌日除外を登録 依頼NO=%s 工程=%s 機械名=%s → %s m",
                        task_id,
                        _log_plain_label(machine),
                        _log_plain_label(machine_name),
                        aladdin_today_exclude_next_day_m,
                    )

        speed = parse_float_safe(speed_raw, 1.0)
        if speed <= 0:
            speed = 1.0

        if qty <= 0 or not machine or not task_id:
            continue

        _line_seq = same_tid_line_seq[task_id]
        same_tid_line_seq[task_id] += 1

        remark_raw = str(row.get(PLAN_COL_SPECIAL_REMARK, "") or "").strip()
        has_special_remark = bool(remark_raw) and remark_raw.lower() not in ("nan", "none")
        remark_implies_due_dispatch_priority = (
            _special_remark_implies_due_related_dispatch_priority(remark_raw)
        )

        ai_one = _ai_task_special_entry_for_row(ai_by_tid, row)
        allow_from_ai_dispatch_signals = (
            has_special_remark
            and _ai_task_special_entry_has_dispatch_priority_signals(ai_one)
        )
        allow_ai_dispatch_priority = (
            remark_implies_due_dispatch_priority or allow_from_ai_dispatch_signals
        )
        req_op, speed_ov, task_eff_factor, priority, start_date_ov, start_time_ov, ai_used = _merge_task_row_with_ai(
            row,
            ai_one,
            allow_ai_dispatch_priority_from_remark=allow_ai_dispatch_priority,
        )
        limited_operator_names = _parse_limited_operator_json_cell(
            _planning_df_cell_scalar(row, PLAN_COL_LIMITED_OP),
            planning_excel_row,
            task_id,
        )
        if answer_due is not None:
            due_basis = answer_due
            due_source = "answer_due"
            due_source_rank = 0
        elif specified_basis is not None:
            due_basis = specified_basis
            due_source = "specified_due"
            due_source_rank = 1
        has_done_deadline_override = False

        if speed_ov is not None:
            speed = speed_ov
        if speed <= 0:
            speed = 1.0

        gsm = _global_speed_multiplier_for_row(
            machine, machine_name, gpo.get("global_speed_rules") or []
        )
        if abs(gsm - 1.0) > 1e-12:
            speed_before_g = speed
            speed = speed * gsm
            if speed <= 0:
                speed = 1.0
            logging.info(
                "メイングローバル: 依頼NO=%s 工程=%s 機械名=%s に speed_multiplier 累穝=%s を適用（速度 %s → %s）",
                task_id,
                _log_plain_label(machine),
                _log_plain_label(machine_name),
                gsm,
                speed_before_g,
                speed,
            )

        # 特別ルール列挙.md（L4/L5/L6/L8）: 列・global_speed で確定した速度のあとに上書き
        from planning_core.dispatch_rules.hook_adapter import apply_speed_special_rules

        speed = apply_speed_special_rules(
            row, task_id, machine, machine_name, speed, _apply_dispatch_speed_special_rules_enumerated_md
        )

        _prod_w = _planning_df_cell_scalar(row, PLAN_COL_PRODUCT_WIDTH)
        try:
            _prod_w_i = int(float(_prod_w)) if _prod_w is not None else None
        except (TypeError, ValueError):
            _prod_w_i = None

        if qty_from_in_progress_next_day_dialog and qty > 1e-12:
            # 翌日配台量は目標総 m。1 ロール=全量だと 1 日の割当プローブが通らず配台 0 になりうるため、
            # シートのロール単位（実効化）でロール数に分割する（段階2・段階3共通）。
            _dlg_sheet_unit = _dispatch_simulator_unit_m_from_plan_row(
                row, fallback_m=qty_total if qty_total > 0 else qty
            )
            if _dlg_sheet_unit > 1e-12 and qty > _dlg_sheet_unit + 1e-6:
                unit = _effective_roll_unit_m_for_dispatch_task_simulator(
                    qty, float(_dlg_sheet_unit)
                )
                _init_rem = max(1.0, float(math.ceil(qty / unit)))
            else:
                unit = float(qty)
                _init_rem = 1.0
        elif dispatch_rolls > 1e-12 and qty > 1e-12:
            unit = float(qty) / float(dispatch_rolls)
            _init_rem = float(dispatch_rolls)
        else:
            unit = _dispatch_simulator_unit_m_from_plan_row(
                row, fallback_m=qty_total if qty_total > 0 else qty
            )
            _sheet_roll_unit_before_sim_adjust = float(unit)
            if qty > 1e-12 and unit > 1e-12:
                unit = _effective_roll_unit_m_for_dispatch_task_simulator(qty, unit)
                if abs(unit - _sheet_roll_unit_before_sim_adjust) > 1e-6:
                    logging.info(
                        "配台シミュレータ: 原反ロール単位を実効化 依頼NO=%s qty_m=%s シート原反ロール=%s → 実効=%s",
                        task_id,
                        qty,
                        _sheet_roll_unit_before_sim_adjust,
                        unit,
                    )
                    if PLAN_COL_RAW_ROLL_UNIT_LENGTH in tasks_df.columns:
                        try:
                            tasks_df.at[row_idx, PLAN_COL_RAW_ROLL_UNIT_LENGTH] = (
                                float(unit)
                            )
                        except Exception as e:
                            logging.warning(
                                "列「%s」の明示更新に失敗（行=%s 依頼NO=%s）: %s",
                                PLAN_COL_RAW_ROLL_UNIT_LENGTH,
                                row_idx,
                                task_id,
                                e,
                            )
                        else:
                            _plan_df_note_effective_roll_unit_iloc(
                                tasks_df, planning_df_iloc
                            )
            _init_rem = (
                float(math.ceil(qty / unit))
                if qty > 1e-12 and unit > 1e-12
                else 0.0
            )

        # 配台使用残数量・配台ロール数が揃っているときは unit_m=残量/本数（実効化しない）。

        # 納期は優先順位・緊急度には使うが、開始日の下限には使わない（余力があれば前倒し開始するため）。
        if due_basis is None:
            due_urgent = False
        else:
            due_urgent = due_basis <= run_date

        # 開始日ルール:
        # 1) 原反投入日があるときは「原反投入日 12:45 以降」を開始可能日時の下限にする。
        #    （日付下限: max(run_date, raw_input_date)」同日時間下限: DISPATCHABLE_FROM_TIME=12:45。
        #    湖南工場かつ在庫場所「湖南」は DISPATCHABLE_FROM_TIME_KONAN_STOCK=9:30、
        #    dispatchable_from_time_for() 参照）
        # 2) 特別指定（セル/AI）の開始日があっても原反投入日より前倒しにはしない（date 下限を維持）
        # 3) 原反投入日が無いときは run_date
        if raw_input_date:
            effective_start_date = max(run_date, raw_input_date)
        else:
            effective_start_date = run_date
        if start_date_ov is not None:
            effective_start_date = (
                max(start_date_ov, raw_input_date)
                if raw_input_date
                else start_date_ov
            )
            if raw_input_date and start_date_ov < raw_input_date:
                logging.info(
                    "開始日上書きは原反投入日より前倒し不可: 依頼NO=%s 指定開始日=%s 原反投入日=%s 採用開始日=%s",
                    task_id,
                    start_date_ov,
                    raw_input_date,
                    effective_start_date,
                )

        _stock_location_for_dispatch = _planning_df_cell_scalar(row, TASK_COL_STOCK_LOCATION)
        same_day_raw_start_limit = (
            dispatchable_from_time_for(_stock_location_for_dispatch)
            if (raw_input_date and effective_start_date == raw_input_date)
            else None
        )

        # 段階2.0: 配台開始の下限は「配台可能日時」列（上書き列 → 算出列 → 原反投入日+12:45）を正とする。
        # 列・上書きが指定されていれば原反投入日由来の下限より優先し、開始日（暦日）と同日開始時刻を上書きする。
        # 原反投入日が無く列も空の行は dispatchable_dt=None となり、従来どおり run_date 起点のまま。
        dispatchable_dt = resolve_dispatchable_datetime_from_plan_row(
            row,
            run_date=run_date,
        )
        if dispatchable_dt is not None:
            disp_date = dispatchable_dt.date()
            effective_start_date = max(run_date, disp_date)
            if start_date_ov is not None and start_date_ov > effective_start_date:
                effective_start_date = start_date_ov
            same_day_raw_start_limit = (
                dispatchable_dt.time() if effective_start_date == disp_date else None
            )

        calc_time_val = qty * speed
        ai_note = ""
        if ai_used:
            try:
                ai_note = json.dumps(ai_used, ensure_ascii=False)[:500]
            except Exception:
                ai_note = str(ai_used)[:500]

        _order_list = seq_by_tid.get(task_id) or []
        _p_rank = _process_sequence_rank_for_machine(machine, _order_list)
        _process_content_mismatch = bool(_order_list) and not _process_name_matches_kakou_content_tokens(
            machine, _order_list
        )

        _dto_from_sheet = None
        if RESULT_TASK_COL_DISPATCH_TRIAL_ORDER in tasks_df.columns:
            from planning_core.core.plan_input import dispatch_trial_order_positive_finite

            _dto_from_sheet = dispatch_trial_order_positive_finite(
                _planning_df_cell_scalar(row, RESULT_TASK_COL_DISPATCH_TRIAL_ORDER)
            )

        _unp_base = _plan_cell_dispatch_remaining_m(row)

        task_queue.append(
            {
                "task_id": task_id,
                "machine": machine,
                "machine_name": machine_name,
                TASK_COL_PRODUCT: product_name,
                PLAN_COL_PRODUCT_WIDTH: _prod_w_i,
                PLAN_COL_PRODUCT_THICKNESS: _planning_df_cell_scalar(
                    row, PLAN_COL_PRODUCT_THICKNESS
                ),
                "process_content_tokens": list(_order_list) if _order_list else [],
                "equipment_line_key": _resolve_equipment_line_key_for_task(
                    {"machine": machine, "machine_name": machine_name},
                    equipment_list,
                ),
                "start_date_req": effective_start_date,
                "answer_due_date": answer_due,
                "specified_due_date": specified_due,
                "specified_due_override": specified_due_ov,
                "due_basis_date": due_basis,
                # 紝期後ゝ倒し再試行で due_basis_date を内部 +1 しても」結果_タスク一覧の当列（列名は互換で「計画基準納期」）はこの値のまま
                "due_basis_date_result_sheet": due_basis,
                "due_source": due_source,
                "due_source_rank": due_source_rank,
                "due_urgent": due_urgent,
                "raw_input_date": raw_input_date,
                "stock_location": _stock_location_for_dispatch,
                "dispatchable_datetime": dispatchable_dt,
                "parent_task_id": parent_task_id,
                "rule_task_id": parent_task_id,
                "same_day_raw_start_limit": same_day_raw_start_limit,
                "total_qty_m": int(qty_total),
                "unit_m": float(unit),
                "remaining_units": _init_rem,
                "base_time_per_unit": (qty / speed) / (qty / unit)
                if unit and speed and qty
                else 0,
                "assigned_history": [],
                "calc_time_value": calc_time_val,
                # 列「加工速度_上書き」「加工速度」・global_speed_rules・特別ルール列挙.md（L4/L5/L6/L8）適用後の m/分
                TASK_COL_SPEED: float(speed),
                "required_op": req_op,
                "task_eff_factor": task_eff_factor,
                "priority": priority,
                "earliest_start_time": start_time_ov,
                "limited_operator_names": limited_operator_names,
                "planning_excel_row": planning_excel_row,
                "task_special_ai_note": ai_note,
                "in_progress": in_progress,
                "qty_from_in_progress_next_day_dialog": qty_from_in_progress_next_day_dialog,
                "aladdin_today_exclude_next_day_m": aladdin_today_exclude_next_day_m,
                "aladdin_today_exclude_next_day_dialog": aladdin_today_exclude_next_day_dialog,
                "aladdin_next_day_exclude_remaining_m": aladdin_today_exclude_next_day_m,
                "has_special_remark": has_special_remark,
                "has_done_deadline_override": has_done_deadline_override,
                "done_qty_reported": done_qty,
                "process_sequence_rank": _p_rank,
                "same_request_line_seq": _line_seq,
                "initial_remaining_units": _init_rem,
                "roll_pipeline_ec": _row_matches_roll_pipeline_ec(machine, machine_name),
                "roll_pipeline_inspection": _row_matches_roll_pipeline_inspection(
                    machine, machine_name
                ),
                "roll_pipeline_rewind": _row_matches_roll_pipeline_rewind(
                    machine, machine_name
                ),
                "process_content_mismatch": _process_content_mismatch,
                "planning_sheet_row_seq": planning_sheet_row_seq,
                "planning_df_iloc": planning_df_iloc,
                "dispatch_trial_order_from_sheet": _dto_from_sheet,
                "unprocessed_baseline_m": _unp_base,
            }
        )
        planning_sheet_row_seq += 1

    _sync_roll_pipeline_start_date_req_min_for_same_request(task_queue)

    logging.info(
        "task_queue 構築完了: total=%s（配台試行によりスキップ %s 行）",
        len(task_queue),
        n_exclude_plan,
    )
    return task_queue
def _sync_roll_pipeline_start_date_req_min_for_same_request(task_queue: list) -> None:
    """
    §B-2/§B-3: 同一依頼NO内で EC/検査/巻返し行の start_date_req が先行工程より遅いと、
    先行完了後も start_date_req 未到達のままカレンダー終端まで進まない（例: W6-4）。
    ロールパイプライン行は同一依頼NOの start_date_req 最小値に揃える。
    """
    by_tid: dict[str, list[tuple[dict, date]]] = {}
    for t in task_queue or []:
        # 枝番タスクは元依頼NO（rule_task_id）でまとめ、同一親の EC/検査/巻返しを同期する。
        tid = _rule_task_id(t)
        if not tid:
            continue
        sdr = t.get("start_date_req")
        if not isinstance(sdr, date):
            continue
        by_tid.setdefault(tid, []).append((t, sdr))
    for tid, items in by_tid.items():
        if not any(
            t.get("roll_pipeline_ec")
            or t.get("roll_pipeline_inspection")
            or t.get("roll_pipeline_rewind")
            for t, _ in items
        ):
            continue
        min_sdr = min(sdr for _, sdr in items)
        for t, sdr in items:
            if not (
                t.get("roll_pipeline_ec")
                or t.get("roll_pipeline_inspection")
                or t.get("roll_pipeline_rewind")
            ):
                continue
            if sdr <= min_sdr:
                continue
            logging.info(
                "roll_pipeline start_date_req 同期: 依頼NO=%s 工程=%s %s→%s",
                tid,
                t.get("machine_name") or t.get("machine"),
                sdr.isoformat(),
                min_sdr.isoformat(),
            )
            t["start_date_req"] = min_sdr
def _task_id_priority_key(task_id):
    """
    依頼NOの同一条件タイブレーク用キー。
    例: Y3-24, Y3-34 のよごな場合はポイフン後坊の数値は尝さい方を優先。
    """
    s = str(task_id or "").strip()
    if not s:
        return ("", 10**9, "")
    parts = s.rsplit("-", 1)
    if len(parts) == 2:
        head = parts[0].strip()
        tail = parts[1].strip()
        if re.match(r"^\d+$", tail):
            return (head, int(tail), s)
    return (s, 10**9, s)
def _wip_l11_bucket_key_for_task_id(task_id: str) -> str:
    """L11 の task_id_head 集計キー（_task_id_priority_key の頭部と同一規則）。"""
    _k = _task_id_priority_key(str(task_id or "").strip())
    return str(_k[0] or _k[2] or "").strip()
def _wip_ec_before_insp_roll_count(
    task_queue: list,
    *,
    task_id_exact: str | None = None,
    task_id_head: str | None = None,
) -> float:
    """
    EC 完了ロール − 後続（検査＋巻返し）完了ロール（負は 0 にクリップ）。
    - 両方 None: 全行（global）
    - task_id_exact: 当該依頼NO行のみ（同一文字列の task_id）
    - task_id_head: 接頭辞グループ（_wip_l11_bucket_key_for_task_id と一致する行）
    """
    if task_id_exact is not None and task_id_head is not None:
        raise ValueError("task_id_exact と task_id_head は同時に指定しない")
    ec_done_total = 0.0
    follower_done_total = 0.0
    for _t in task_queue:
        _tid = str(_t.get("task_id") or "").strip()
        if task_id_exact is not None:
            if _tid != task_id_exact:
                continue
        elif task_id_head is not None:
            if _wip_l11_bucket_key_for_task_id(_tid) != task_id_head:
                continue
        init = float(_t.get("initial_remaining_units") or 0)
        rem = float(_t.get("remaining_units") or 0)
        done = max(0.0, init - rem)
        if done <= 1e-12:
            continue
        if _t.get("roll_pipeline_ec"):
            ec_done_total += done
        elif _t.get("roll_pipeline_inspection") or _t.get("roll_pipeline_rewind"):
            follower_done_total += done
    return max(0.0, ec_done_total - follower_done_total)
def _serial_dispatch_order_task_ids(task_queue) -> list:
    """
    依頼NO直列配台の処理順。坄依頼NOについで **配台試行順番の最尝値** は尝さい依頼を先に完走させる
    （同一依頼内の複数行は最尝幅の試行順で代表）。タイブレークは計画シート上の先行行
    （planning_sheet_row_seq）と依頼NOキー。
    """
    min_dto_by_tid: dict = {}
    first_seq_by_tid: dict = {}
    for t in task_queue:
        tid = str(t.get("task_id", "") or "").strip()
        if not tid:
            continue
        try:
            from planning_core.core.plan_input import dispatch_trial_order_key_from_task

            dto = dispatch_trial_order_key_from_task(t)
        except (TypeError, ValueError):
            dto = 10**9
        prev_d = min_dto_by_tid.get(tid)
        if prev_d is None or dto < prev_d:
            min_dto_by_tid[tid] = dto
        seq = t.get("planning_sheet_row_seq")
        seq = int(seq) if seq is not None else 10**9
        prev = first_seq_by_tid.get(tid)
        if prev is None or seq < prev:
            first_seq_by_tid[tid] = seq
    return sorted(
        min_dto_by_tid.keys(),
        key=lambda tid: (
            min_dto_by_tid[tid],
            first_seq_by_tid.get(tid, 10**9),
            _task_id_priority_key(tid),
        ),
    )
def _excel_scalar_to_plan_string_cell(v):
    """
    既存シート（read_excel）由来のスカラーを」配台計画 DataFrame の文字列列（StringDtype）へ
    代入でしる str に正規化する。Excel は数値として保挝した優先度 1 → \"1\" など。
    """
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    if isinstance(v, bool):
        return str(v).lower()
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if isinstance(v, float) and math.isfinite(v) and float(int(v)) == v:
            return str(int(v))
        if isinstance(v, float) and math.isfinite(v):
            s = str(v)
            if "." in s:
                s = s.rstrip("0").rstrip(".")
            return s
        return str(int(v))
    if isinstance(v, pd.Timestamp):
        if pd.isna(v):
            return ""
        if v.hour == 0 and v.minute == 0 and v.second == 0 and v.microsecond == 0:
            return v.strftime("%Y/%m/%d")
        return v.strftime("%Y/%m/%d %H:%M")
    if isinstance(v, datetime):
        # pd.NaT は Timestamp ではないが datetime 互換で strftime 不可
        if pd.isna(v):
            return ""
        if v.hour == 0 and v.minute == 0 and v.second == 0 and v.microsecond == 0:
            return v.date().strftime("%Y/%m/%d")
        return v.strftime("%Y/%m/%d %H:%M")
    if isinstance(v, date):
        return v.strftime("%Y/%m/%d")
    if isinstance(v, str):
        return v.strip()
    return str(v).strip()
def _merge_plan_sheet_user_overrides(out_df):
    """
    ブック内の「配台計画_タスク入力」にユーザーは入力した上書き列を」
    段階1の抽出結果へ (依頼NO, 工程名) 短縮で引き継ぎ。
    空のセルはマージしない（新規抽出坴の空のまま）。

    「ロール単位長さ」は、製品名が「製品名,ロール単位の長さ.txt」に登録されている行では
    シート上の旧値で上書きしない（テーブル＋段階1の再計算を優先し、誤った過去値の混入を防ぐ）。
    「(原反)ロール単位長さ」は旧シートからマージしない（使用原反列と当該テーブル由来のみ）。

    工程「分割」で同一依頼NO内に同一機械名の複数行がないとき、過去に誤って付いた
    「配台不要」=オン相当の値はマージしない（``apply_exclude_rules_config_to_plan_df`` の
    分割ガードと同じ重複条件）。手入力で分割かつ配台不要オンを単独行に付けたい場合は稀とし、
    シート上で再入力すれば引き継がれる。
    """
    if out_df is None or out_df.empty:
        return out_df
    _wbm = _excel_plan_input_wb()
    if not _wbm or not os.path.exists(_wbm):
        return out_df
    try:
        df_old = pd.read_excel(_wbm, sheet_name=PLAN_INPUT_SHEET_NAME)
    except Exception as e:
        logging.info("段階1: 既存の配台シートを読めないため上書き継承をスキップ (%s)", e)
        return out_df
    df_old.columns = df_old.columns.str.strip()
    df_old = _migrate_deprecated_plan_override_columns(df_old)
    df_old = _align_dataframe_headers_to_canonical(
        df_old,
        plan_input_sheet_column_order(),
    )
    df_old = _coalesce_plan_plain_remark_into_special(df_old)
    if TASK_COL_TASK_ID not in df_old.columns or TASK_COL_MACHINE not in df_old.columns:
        return out_df

    lookup = {}
    for _, r in df_old.iterrows():
        tid = planning_task_id_str_from_plan_row(r)
        mach = str(r.get(TASK_COL_MACHINE, "") or "").strip()
        if not tid or not mach:
            continue
        key = (tid, mach)
        bucket = lookup.setdefault(key, {})
        for c in (*PLAN_STAGE1_MERGE_COLUMNS, *PLAN_STAGE1_MERGE_EXTRA_COLUMNS):
            if c not in df_old.columns or c not in out_df.columns:
                continue
            v = r.get(c)
            # float 以外の欠損（例: pd.NaT）も空セル扱いでマージしない
            if v is None or pd.isna(v):
                continue
            if isinstance(v, str):
                s = v.strip()
                if not s or s.lower() in ("nan", "none"):
                    continue
            bucket[c] = v

    if not lookup:
        return out_df

    by_tid_idx: dict[str, list] = defaultdict(list)
    if TASK_COL_TASK_ID in out_df.columns:
        for j in out_df.index:
            tj = _normalize_task_id_for_dup_grouping(out_df.at[j, TASK_COL_TASK_ID])
            if tj:
                by_tid_idx[tj].append(j)

    merged_rows = 0
    for i, row in out_df.iterrows():
        tid = planning_task_id_str_from_plan_row(row)
        mach = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        bucket = lookup.get((tid, mach))
        if not bucket:
            continue
        merged_rows += 1
        for c, v in bucket.items():
            if (
                c == PLAN_COL_ROLL_UNIT_LENGTH
                and TASK_COL_PRODUCT in out_df.columns
            ):
                _pn_m = out_df.at[i, TASK_COL_PRODUCT]
                if _lookup_roll_unit_length_m_from_table(_pn_m) is not None:
                    continue
            if c == PLAN_COL_EXCLUDE_FROM_ASSIGNMENT:
                tp_m = str(row.get(TASK_COL_MACHINE, "") or "").strip()
                if _process_name_is_bunkatsu_for_auto_exclude(tp_m):
                    tid_m = _normalize_task_id_for_dup_grouping(
                        out_df.at[i, TASK_COL_TASK_ID]
                    )
                    dup_m = False
                    if tid_m:
                        dup_m, _ = _same_tid_nonempty_machine_dup_ge2(
                            out_df, by_tid_idx.get(tid_m, [])
                        )
                    if not dup_m:
                        coerced_ex = _coerce_plan_exclude_column_value_for_storage(v)
                        if _plan_row_exclude_from_assignment(
                            {PLAN_COL_EXCLUDE_FROM_ASSIGNMENT: coerced_ex}
                        ):
                            continue
                v = _coerce_plan_exclude_column_value_for_storage(v)
            elif c in out_df.columns and pd.api.types.is_string_dtype(out_df[c].dtype):
                v = _excel_scalar_to_plan_string_cell(v)
            out_df.at[i, c] = v

    if merged_rows:
        logging.info(
            "段階1: 既存シートから上書き列を %s 行へ引し継ねました（キー: 依頼NO+工程名）。",
            merged_rows,
        )
    return out_df
def _auto_exclude_cell_empty_for_autofill(v) -> bool:
    """配台試行セルは未入力のときの値自動で yes を書き込む。"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return True
    if isinstance(v, str):
        s = str(v).strip()
        return not s or s.lower() in ("nan", "none")
    return False
def _normalize_task_id_for_dup_grouping(raw) -> str:
    """同一依頼NOのグルーピング用（表記ゆれ・英字の大尝を寄せる）。"""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    if isinstance(raw, float) and raw == int(raw):
        s = str(int(raw))
    else:
        s = unicodedata.normalize("NFKC", str(raw).strip())
    s = s.strip()
    if not s or s.lower() == "nan":
        return ""
    return s.upper()
def _process_name_is_bunkatsu_for_auto_exclude(raw) -> bool:
    """工程名は「分割」（空白除去・NFKC 後）。"""
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    t = re.sub(r"[\s　]+", "", t)
    return t == "分割"
def _same_tid_nonempty_machine_dup_ge2(
    df: pd.DataFrame, idx_list: list
) -> tuple[bool, dict[str, int]]:
    """
    ``_apply_auto_exclude_bunkatsu_duplicate_machine`` と同一の重複判定。
    idx_list 内で、正規化後の非空「機械名」が同一の行が2件以上あるとき True。
    """
    if len(idx_list) < 2:
        return False, {}
    counts: dict[str, int] = defaultdict(int)
    for i in idx_list:
        mn_key = _normalize_equipment_match_key(df.at[i, TASK_COL_MACHINE_NAME])
        if not mn_key:
            continue
        counts[mn_key] += 1
    dup_ge2 = any(c >= 2 for c in counts.values())
    return dup_ge2, dict(counts)
def _apply_auto_exclude_bunkatsu_duplicate_machine(
    df: pd.DataFrame, log_prefix: str = "段階1"
) -> pd.DataFrame:
    """
    同一依頼NOは2行以上あり」かつ空でない同一機械名は2行以上あるグループでは」
    工程名は「分割」の行の「配台試行」に yes を入れる（セルは空のときのみ）。
    機械名は _normalize_equipment_match_key で重複判定。
    """
    if df is None or df.empty:
        return df
    need_cols = (TASK_COL_TASK_ID, TASK_COL_MACHINE, TASK_COL_MACHINE_NAME)
    for c in need_cols:
        if c not in df.columns:
            return df
    if PLAN_COL_EXCLUDE_FROM_ASSIGNMENT not in df.columns:
        df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = ""
    # read_excel 等で StringDtype になると数値・真偽の .at 代入で TypeError になるため、 object に寄せる
    df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT].astype(object)

    by_tid = defaultdict(list)
    for i in df.index:
        tid = _normalize_task_id_for_dup_grouping(df.at[i, TASK_COL_TASK_ID])
        if not tid:
            continue
        by_tid[tid].append(i)

    n_set = 0
    for _tid_key, idx_list in by_tid.items():
        dup_ge2, _counts = _same_tid_nonempty_machine_dup_ge2(df, idx_list)
        if not dup_ge2:
            continue
        for i in idx_list:
            if not _process_name_is_bunkatsu_for_auto_exclude(df.at[i, TASK_COL_MACHINE]):
                continue
            if not _auto_exclude_cell_empty_for_autofill(
                df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT]
            ):
                continue
            # 列は StringDtype のとき int 代入で TypeError になるため、文字列にれる（_plan_row_exclude_from_assignment は yes を真とみなす）
            df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
            n_set += 1

    if n_set:
        logging.info(
            "%s: 同一依頼NOかつ同一機械名は複数行あるグループで」工程名「分割」の行 %s 件に「配台試行」=yes を自動設定しました。",
            log_prefix,
            n_set,
        )
    return df
def _is_in_house_self_processing_task_id(task_id) -> bool:
    """依頼NO先頭が「2」の自社加工品（配台対象外）。"""
    tid = planning_task_id_str_from_scalar(task_id)
    return bool(tid) and tid.lstrip().startswith("2")
def _apply_auto_exclude_in_house_self_processing(
    df: pd.DataFrame, log_prefix: str = "段階1"
) -> pd.DataFrame:
    """
    依頼NOが「2」から始まる自社加工品の「配台不要」に yes を入れる（セルは空のときのみ）。
    手入力の配台不要は上書きしない。
    """
    if df is None or df.empty:
        return df
    if TASK_COL_TASK_ID not in df.columns:
        return df
    if PLAN_COL_EXCLUDE_FROM_ASSIGNMENT not in df.columns:
        df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = ""
    df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT].astype(
        object
    )
    n_set = 0
    for i in df.index:
        if not _is_in_house_self_processing_task_id(df.at[i, TASK_COL_TASK_ID]):
            continue
        if not _auto_exclude_cell_empty_for_autofill(
            df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT]
        ):
            continue
        df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
        n_set += 1
    if n_set:
        logging.info(
            "%s: 自社加工品（依頼NO先頭が2）%s 行の「配台不要」=yes を自動設定しました。",
            log_prefix,
            n_set,
        )
    return df
def _normalize_process_name_for_rule_match(raw) -> str:
    """工程名のルール照合（NFKC・空白除去）。"""
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    t = re.sub(r"[\s　]+", "", t)
    return t
def _master_speed_sheet_apply_enabled() -> bool:
    """環境変数 MASTER_USE_SPEED_SHEET で master.xlsm の speed 由来の加工速度上書きを無効化できる。"""
    raw = os.environ.get("MASTER_USE_SPEED_SHEET", "1")
    v = str(raw).strip().lower()
    if not v:
        return True
    return v not in ("0", "false", "no", "off")
def _master_speed_first_excel_col_1based() -> int:
    """speed シートで設備列が始まる Excel 列番号（既定 4 = D 列）。環境変数 MASTER_SPEED_FIRST_EXCEL_COL。"""
    raw = os.environ.get("MASTER_SPEED_FIRST_EXCEL_COL", "").strip()
    if not raw:
        return 4
    try:
        n = int(raw)
    except ValueError:
        return 4
    return n if n >= 1 else 4
def _load_master_speed_lookup_from_master_workbook() -> dict[tuple[str, str], float]:
    """
    master.xlsm の speed シートから (工程名, 機械名) 正規化キー → 加工速度 (m/分)。
    速度は Excel 4 行目×5 行目（基本速度×実稼働比率）。同一キーが複数列で数値が食い違うときは先頭列を採用。
    """
    out: dict[tuple[str, str], float] = {}
    if not _master_speed_sheet_apply_enabled():
        logging.info("master.xlsm speed シートによる加工速度の上書きは無効です（MASTER_USE_SPEED_SHEET）。")
        return out
    path = _require_master_workbook_path_exists()
    sheet = MASTER_SHEET_SPEED
    try:
        raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    except Exception as e:
        logging.info(
            "master.xlsm speed: シート %r を読めません（%s）。加工速度は従来どおりです。",
            sheet,
            e,
        )
        return out
    if raw is None or raw.empty or raw.shape[0] < 5:
        logging.info("master.xlsm speed: 行が不足しています（5行目まで必須）。")
        return out
    first_col = _master_speed_first_excel_col_1based()
    c0 = first_col - 1
    if raw.shape[1] <= c0:
        logging.info(
            "master.xlsm speed: 列が足りません（データ開始列=%s）。",
            first_col,
        )
        return out
    conflicts: list[tuple[tuple[str, str], float, float]] = []
    for j in range(c0, raw.shape[1]):
        p_raw = raw.iat[0, j]
        m_raw = raw.iat[1, j]
        if p_raw is None or (isinstance(p_raw, float) and pd.isna(p_raw)):
            p_str = ""
        else:
            p_str = str(p_raw).strip()
        if m_raw is None or (isinstance(m_raw, float) and pd.isna(m_raw)):
            m_str = ""
        else:
            m_str = str(m_raw).strip()
        if not p_str and not m_str:
            continue
        p_norm = _normalize_process_name_for_rule_match(p_str)
        m_norm = _normalize_equipment_match_key(m_str)
        if not p_norm or not m_norm:
            continue
        bs_raw = raw.iat[3, j]
        rr_raw = raw.iat[4, j]
        base = parse_float_safe(bs_raw, 0.0)
        ratio = parse_float_safe(rr_raw, 0.0)
        if base <= 0 or ratio <= 0:
            continue
        spd = float(base * ratio)
        if spd <= 0:
            continue
        key = (p_norm, m_norm)
        if key in out:
            if abs(out[key] - spd) > 1e-6:
                conflicts.append((key, out[key], spd))
            continue
        out[key] = spd
    if conflicts:
        logging.warning(
            "master.xlsm speed: 同一工程+機械キーに複数の速度列があり数値が異なります（先頭列を採用）。例: %s",
            conflicts[:5],
        )
    if out:
        logging.info(
            "master.xlsm speed: シート %r から %s 件の (工程名, 機械名) 速度を読み込みました。",
            sheet,
            len(out),
        )
    else:
        logging.info(
            "master.xlsm speed: シート %r に有効な速度列がありませんでした。",
            sheet,
        )
    return out
def _apply_master_speed_sheet_to_plan_df(
    df: "pd.DataFrame",
    *,
    log_prefix: str,
) -> None:
    """配台計画 DataFrame の「加工速度」を master.xlsm speed に一致する行だけ上書きする。"""
    if df is None or df.empty:
        return
    if TASK_COL_SPEED not in df.columns:
        return
    if TASK_COL_MACHINE not in df.columns or TASK_COL_MACHINE_NAME not in df.columns:
        return
    lu = _load_master_speed_lookup_from_master_workbook()
    if not lu:
        return
    n_hit = 0
    n_miss = 0
    for i, row in df.iterrows():
        key = (
            _normalize_process_name_for_rule_match(row.get(TASK_COL_MACHINE)),
            _normalize_equipment_match_key(row.get(TASK_COL_MACHINE_NAME)),
        )
        spd = lu.get(key)
        if spd is not None and spd > 0:
            df.at[i, TASK_COL_SPEED] = spd
            n_hit += 1
        else:
            n_miss += 1
    logging.info(
        "%s: master.xlsm「%s」の速度を %s 行に適用（マスタ未該当 %s 行は加工速度セルを変更しませんでした）。",
        log_prefix,
        MASTER_SHEET_SPEED,
        n_hit,
        n_miss,
    )
def _exclude_rules_sheet_header_map(ws) -> dict:
    """1行目見出し → 列番坷(1始まり)。
    openpyxl は新規シート直後に max_column は 0 のままのことはあり」見出しは読ゝう保存剝に return してしまご。
    しのため、最低 A～E 列は必う走査れる。
    """
    def _norm_hdr(x) -> str:
        if x is None:
            return ""
        s = unicodedata.normalize("NFKC", str(x))
        s = s.replace("\u3000", " ").replace("\xa0", " ").replace("\t", " ")
        s = " ".join(s.split())
        return s.strip()

    h = {}
    last_col = max(5, int(ws.max_column or 0))
    for col in range(1, last_col + 1):
        v = ws.cell(1, col).value
        k = _norm_hdr(v)
        if k:
            h[k] = col
    return h
def _ensure_exclude_rules_sheet_headers_and_columns(ws, log_prefix: str) -> tuple[int, int, int, int, int]:
    """
    1行目に標準見出し（工程名・機械名・配台不要・配台不要ロジック・ロジック式）はあることを保証れる。
    手動で空シートの値追加した場合は A1:E1 は空のため、ここで書き込んで列番坷を返す。
    """
    headers = (
        EXCLUDE_RULE_COL_PROCESS,
        EXCLUDE_RULE_COL_MACHINE,
        EXCLUDE_RULE_COL_FLAG,
        EXCLUDE_RULE_COL_LOGIC_JA,
        EXCLUDE_RULE_COL_LOGIC_JSON,
    )
    hm = _exclude_rules_sheet_header_map(ws)
    want = tuple(
        unicodedata.normalize("NFKC", str(x)).replace("\u3000", " ").strip()
        for x in headers
    )
    if all(hm.get(x) for x in want):
        return tuple(hm[x] for x in want)

    # 見出し文字列が文字化け・表記ゆれ等で一致しない場合でも、
    # 設定_配台不要工程は A～E 列固定で扱うため、A1:E1 のいずれかが非空なら列位置は確定とみなし、
    # 毎回の HEADER_FIX（=ブック変更→保存→段階1遅延）を避ける。
    try:
        a1e1 = [ws.cell(1, c).value for c in range(1, 6)]
        if any(v is not None and str(v).strip() for v in a1e1):
            logging.info(
                "%s: 「%s」の見出しは標準と一致しませんが、A1:E1 が非空のため列位置(1..5)をそのまま使用します。",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
            )
            return (1, 2, 3, 4, 5)
    except Exception:
        pass
    for i, name in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=name)
    logging.info(
        "%s: 「%s」の見出しは無い＝列名は一致しないため、標準の1行目（A1:E1）を設定しました。",
        log_prefix,
        EXCLUDE_RULES_SHEET_NAME,
    )
    return (1, 2, 3, 4, 5)
def _compact_exclude_rules_data_rows(
    ws,
    c_proc: int,
    c_mach: int,
    c_flag: int,
    c_d: int,
    c_e: int,
    log_prefix: str,
) -> tuple[int, int]:
    """
    2 行目以降から「空行」を除いで上に詰ゝる（元の並よは維挝」ソートしない）。
    空行: 工程名は空」または A～E 相当の5セルはまとめて空白相当。
    Returns (残したデータ行数, 削除した行数).
    """
    max_r0 = int(ws.max_row or 1)
    max_r = max_r0
    if max_r < 2:
        return 0, 0

    # openpyxl の max_row は「書式だけ残っている末尾空行」を含み得る。
    # 末尾の空行を毎回 delete_rows で詰め直すと、段階1が遅くなるため無視する。
    try:
        last_data_r = 1
        for r in range(max_r0, 1, -1):
            pv = ws.cell(row=r, column=c_proc).value
            p = (
                ""
                if pv is None or (isinstance(pv, float) and pd.isna(pv))
                else str(pv).strip()
            )
            if p:
                last_data_r = r
                break
        max_r = max(2, int(last_data_r))
    except Exception:
        max_r = max_r0

    old_body = max_r - 1
    cols = (c_proc, c_mach, c_flag, c_d, c_e)
    rows: list[tuple[str, str, object, object, object]] = []
    for r in range(2, max_r + 1):
        pv = ws.cell(row=r, column=c_proc).value
        mv = ws.cell(row=r, column=c_mach).value
        cv = ws.cell(row=r, column=c_flag).value
        dv = ws.cell(row=r, column=c_d).value
        ev = ws.cell(row=r, column=c_e).value
        all_blank = all(
            _cell_is_blank_for_rule(ws.cell(row=r, column=c).value) for c in cols
        )
        p = str(pv).strip() if pv is not None and not (isinstance(pv, float) and pd.isna(pv)) else ""
        m = str(mv).strip() if mv is not None and not (isinstance(mv, float) and pd.isna(mv)) else ""
        if all_blank or not p:
            continue
        rows.append((p, m, cv, dv, ev))

    n_skip = old_body - len(rows)

    if not rows:
        # 末尾の空行のみは無視済み。ここで delete_rows すると毎回変更扱いになるため、空行削除はしない。
        if old_body > 0:
            logging.info(
                "%s: 「%s」は有効なデータ行はなかったため、データ行 %s 行を削除しました。",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
                old_body,
            )
        return 0, 0

    if n_skip <= 0:
        return len(rows), 0

    ws.delete_rows(2, old_body)
    for i, (p, m, cv, dv, ev) in enumerate(rows, start=2):
        ws.cell(row=i, column=c_proc, value=p)
        ws.cell(row=i, column=c_mach, value=m)
        ws.cell(row=i, column=c_flag, value=cv)
        ws.cell(row=i, column=c_d, value=dv)
        ws.cell(row=i, column=c_e, value=ev)

    if n_skip:
        logging.info(
            "%s: 「%s」から空行を %s 件削除し、%s 行に詰ゝました（並よ順は維挝）。",
            log_prefix,
            EXCLUDE_RULES_SHEET_NAME,
            n_skip,
            len(rows),
        )
    return len(rows), n_skip
def _cell_is_blank_for_rule(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    s = str(v).strip()
    return not s or s.lower() in ("nan", "none", "null")
def _exclude_rule_c_column_is_yes(v) -> bool:
    """C列「配台不要」はオン（この工程+機械は常に配台対象外）。"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return int(v) == 1
        except (TypeError, ValueError):
            pass
    s = unicodedata.normalize("NFKC", str(v).strip()).lower()
    return s in ("yes", "true", "1", "y", "はい", "○", "〇", "◝")
def _task_row_matches_exclude_rule_target(
    task_proc: str, task_mach: str, rule_proc: str, rule_mach: str
) -> bool:
    if _normalize_process_name_for_rule_match(task_proc) != _normalize_process_name_for_rule_match(
        rule_proc
    ):
        return False
    rm = str(rule_mach or "").strip()
    if not rm:
        # 機械名が空のルールは「当該工程の全機械」を意味するが、工程「分割」は同一依頼で別行の
        # スリット等と同じ機械名が重なるケースが多く、ワイルドカード一致だと実設備向け分割行まで
        # 一律配台不要になる。分割を除外する場合は設定シートで機械名を明示する。
        if _process_name_is_bunkatsu_for_auto_exclude(task_proc):
            return False
        return True
    return _normalize_equipment_match_key(task_mach) == _normalize_equipment_match_key(rm)
def _collect_process_machine_pairs_for_exclude_rules(df_src: pd.DataFrame) -> list[tuple[str, str]]:
    """加工計画DATA から」段階1とともに抽出条件で (工程名, 機械名) の一覧（重複除し・順庝維挝）。"""
    out: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for _, row in df_src.iterrows():
        if row_has_completion_keyword(row):
            continue
        if _plan_row_exclude_as_completed_mikan_unprocessed_zero_actual_done_rule(row):
            continue
        task_id = planning_task_id_str_from_scalar(row.get(TASK_COL_TASK_ID))
        machine = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        machine_name = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
        qty, _done_qty, _qty_total, _from_unp = _plan_row_dispatch_qty_metrics(row)
        if qty <= 0 or not machine or not task_id:
            continue
        key = (
            _normalize_process_name_for_rule_match(machine),
            _normalize_equipment_match_key(machine_name),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append((machine, machine_name))
    return out
def _collect_plan_input_process_machine_pairs_for_exclude_rules_sync(
    df: "pd.DataFrame",
) -> list[tuple[str, str]]:
    """配台計画_タスク入力相当の DataFrame から (工程名, 機械名) を重複除しで列挙。

    ``run_exclude_rules_sheet_maintenance`` の行同期と同一の正規化キーでまとめる。
    """
    out: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    if df is None or getattr(df, "empty", True):
        return out
    if TASK_COL_MACHINE not in df.columns:
        return out
    _has_mach_name = TASK_COL_MACHINE_NAME in df.columns
    for _, row in df.iterrows():
        p = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        m = (
            str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
            if _has_mach_name
            else ""
        )
        if not p:
            continue
        key = (
            _normalize_process_name_for_rule_match(p),
            _normalize_equipment_match_key(m),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append((p, m))
    return out
def _parse_exclude_rule_json_cell(raw) -> dict | None:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    s = str(raw).strip()
    if not s:
        return None
    fence = re.search(
        r"```(?:json)?\s*(\{.*\})\s*```",
        s,
        re.DOTALL | re.IGNORECASE,
    )
    if fence:
        s = fence.group(1).strip()
    try:
        obj = json.loads(s)
    except json.JSONDecodeError:
        return None
    return obj if isinstance(obj, dict) else None
def _validate_exclude_rule_parsed_dict(o: object) -> dict | None:
    """Gemini＝E列から得た dict は配台試行ルールとして有効か。"""
    if not isinstance(o, dict):
        return None
    if int(o.get("version") or 0) != 1:
        return None
    mode = str(o.get("mode") or "").strip().lower()
    if mode not in ("always_exclude", "conditions"):
        return None
    return o
def _exclude_rule_de_cache_key(stripped_blob: str) -> str:
    """「配台不要ロジック」文言（正規化済み）に対れる ai_remarks_cache 用キー。"""
    h = hashlib.sha256(stripped_blob.encode("utf-8")).hexdigest()
    return f"{AI_CACHE_KEY_PREFIX_EXCLUDE_RULE_DE}:{h}"
def _cache_get_exclude_rule_de_parsed(cache_obj: dict, blob: str) -> dict | None:
    s = str(blob or "").strip()
    if not s:
        return None
    data = get_cached_ai_result(
        cache_obj, _exclude_rule_de_cache_key(s), content_key=s
    )
    if not isinstance(data, dict):
        return None
    return _validate_exclude_rule_parsed_dict(data)
def _cache_put_exclude_rule_de_parsed(
    cache_obj: dict, blob: str, parsed: dict | None
) -> None:
    if parsed is None:
        return
    s = str(blob or "").strip()
    if not s:
        return
    put_cached_ai_result(
        cache_obj, _exclude_rule_de_cache_key(s), parsed, content_key=s
    )
def _exclude_rule_logic_gemini_schema_instructions() -> str:
    allowed = ", ".join(sorted(EXCLUDE_RULE_ALLOWED_COLUMNS))
    return (
        "」スキーマ version は必う 1】\n"
        "1) 常に配台試行（説明は条件なしで外れ愝味）のとき:\n"
        '{"version":1,"mode":"always_exclude"}\n\n'
        "2) 列の条件で配台試行とれるとし:\n"
        '{"version":1,"mode":"conditions","require_all": true または false,"conditions":[ ... ]}\n\n'
        "conditions の坄覝素:\n"
        "- {\"column\":\"列名\",\"op\":\"empty\"} … セルは空\n"
        "- {\"column\":\"列名\",\"op\":\"not_empty\"}\n"
        "- {\"column\":\"列名\",\"op\":\"eq\",\"value\":\"文字列\"} / ne / contains / not_contains / regex（正覝表睾）\n"
        "- {\"column\":\"列名\",\"op\":\"gt\"|\"gte\"|\"lt\"|\"lte\",\"value\":数値} … 数値比較（列は数として解釈）\n\n"
        f"」使用可能な列名のみ】（これ以外は使えない）:\n{allowed}\n"
    )
def _parse_exclude_rule_json_array_response(text: str) -> list | None:
    """モデル応答から JSON 配列を取り出す（```json フェンス付し坯）。"""
    s = (text or "").strip()
    if not s:
        return None
    fence = re.search(
        r"```(?:json)?\s*(\[.*\])\s*```",
        s,
        re.DOTALL | re.IGNORECASE,
    )
    if fence:
        s = fence.group(1).strip()
    try:
        obj = json.loads(s)
    except json.JSONDecodeError:
        return None
    return obj if isinstance(obj, list) else None
def _row_scalar_for_exclude_rule(row, col_name: str):
    try:
        return row.get(col_name)
    except Exception:
        return None
def _exclude_rule_string_eq_allow_csv_tokens(val_s: str, pat_raw) -> bool:
    """配台不要条件の eq 用。完全一致に加え、セルが「A,B」「A、B」のように複数トークンを含むときはいずれかが pat と一致すれば真。"""
    pat = "" if pat_raw is None else str(pat_raw).strip()
    if val_s == pat:
        return True
    if "," in val_s or "\u3001" in val_s:
        parts = [p.strip() for p in re.split(r"[,、]", val_s) if p.strip()]
        return bool(pat) and pat in parts
    return False
def _evaluate_exclude_rule_one_condition(cond: dict, row) -> bool:
    if not isinstance(cond, dict):
        return False
    col = cond.get("column")
    if col not in EXCLUDE_RULE_ALLOWED_COLUMNS:
        logging.warning("配台試行ルール: 未対応の列名をスキップしました: %s", col)
        return False
    op = str(cond.get("op") or "").strip().lower()
    val = _row_scalar_for_exclude_rule(row, col)
    val_s = "" if val is None or (isinstance(val, float) and pd.isna(val)) else str(val).strip()
    val_s_lower = val_s.lower()

    if op == "empty":
        return val_s == ""
    if op == "not_empty":
        return val_s != ""

    if op in ("contains", "not_contains", "regex", "eq", "ne"):
        rhs = cond.get("value", "")
        pat = "" if rhs is None else str(rhs)
        if op == "contains":
            return pat in val_s
        if op == "not_contains":
            return pat not in val_s
        if op == "regex":
            try:
                return re.search(pat, val_s) is not None
            except re.error:
                return False
        if op == "eq":
            return _exclude_rule_string_eq_allow_csv_tokens(val_s, pat)
        if op == "ne":
            return not _exclude_rule_string_eq_allow_csv_tokens(val_s, pat)

    def _num(x):
        try:
            return float(x)
        except (TypeError, ValueError):
            return None

    nv = _num(val)
    cv = _num(cond.get("value"))
    if nv is None or cv is None:
        return False
    if op == "gt":
        return nv > cv
    if op == "gte":
        return nv >= cv
    if op == "lt":
        return nv < cv
    if op == "lte":
        return nv <= cv
    return False
def evaluate_exclude_rule_json_for_row(rule: dict, row) -> bool:
    """
    E列の JSON（version=1）を評価し、当該タスク行を配台試行とれきしなら True。
    mode: always_exclude | conditions
    """
    if not isinstance(rule, dict) or int(rule.get("version") or 0) != 1:
        return False
    mode = str(rule.get("mode") or "").strip().lower()
    if mode == "always_exclude":
        return True
    if mode != "conditions":
        return False
    conds = rule.get("conditions")
    if not isinstance(conds, list) or not conds:
        return False
    require_all = bool(rule.get("require_all", True))
    checks = []
    for c in conds:
        if isinstance(c, dict) and c.get("column") in EXCLUDE_RULE_ALLOWED_COLUMNS:
            checks.append(_evaluate_exclude_rule_one_condition(c, row))
    if not checks:
        return False
    return all(checks) if require_all else any(checks)
def _ai_compile_exclude_rule_logic_to_json(natural_language: str) -> dict | None:
    """
    D列の自然言語を Gemini で JSON ルールに変杛。失敗時 None。
    json/ai_remarks_cache.json に TTL 付しでキャッシュ（同一文言なら API を呼みない）。
    """
    blob = str(natural_language or "").strip()
    if not blob:
        return None
    ai_cache = load_ai_cache()
    hit = _cache_get_exclude_rule_de_parsed(ai_cache, blob)
    if hit is not None:
        logging.info("配台不要ルール: AIキャッシュヒット（配台不要ロジック→JSON）")
        return hit
    if not API_KEY:
        return None
    schema = _exclude_rule_logic_gemini_schema_instructions()
    prompt = (
        "あなたは工場の配台システム用です。次の「配台試行の説明」を」タスク1行を判定れる機械坯読ルールに変杛してください。\n\n"
        "」出力】先頭は { で終ゝりは } の JSON オブジェクト1つのみ（説明・マークダウン禁止）。\n\n"
        f"{schema}\n"
        f"」説明文】\n{blob}\n"
    )
    try:
        ppath = os.path.join(log_dir, "ai_exclude_rule_logic_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("配台試行ルール: プロンプト → %s", ppath)
    except OSError as ex:
        logging.warning("配台試行ルール: プロンプト保存失敗: %s", ex)
    try:
        client = _gemini_client(API_KEY)
        res, gem_model_used = _gemini_generate_content_with_retry(
            client, contents=prompt, log_label="配台不要ルールD→E"
        )
        record_gemini_response_usage(res, gem_model_used)
        raw = (_gemini_result_text(res) or "").strip()
        rpath = os.path.join(log_dir, "ai_exclude_rule_logic_last_response.txt")
        try:
            with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                rf.write(raw)
        except OSError:
            pass
        parsed = _validate_exclude_rule_parsed_dict(_parse_exclude_rule_json_cell(raw))
        if parsed:
            _cache_put_exclude_rule_de_parsed(ai_cache, blob, parsed)
            save_ai_cache(ai_cache)
        return parsed
    except Exception as e:
        logging.warning("配台試行ルール: Gemini 変杛失敗: %s", e)
        return None
def _ai_compile_exclude_rule_logics_batch(blobs: list[str]) -> list[dict | None]:
    """
    複数の D 列文言を 1 回の Gemini 呼び出しで JSON 化。失敗・覝素数厳密一致時は 1 件うつにフォールバック。
    json/ai_remarks_cache.json にヒットした文言は API を呼みない。
    """
    n = len(blobs)
    if n == 0:
        return []
    ai_cache = load_ai_cache()
    out: list[dict | None] = [None] * n
    pend_i: list[int] = []
    pend_b: list[str] = []
    for i, b in enumerate(blobs):
        s = str(b).strip()
        hit = _cache_get_exclude_rule_de_parsed(ai_cache, s) if s else None
        if hit is not None:
            out[i] = hit
        else:
            pend_i.append(i)
            pend_b.append(s)
    if not pend_b:
        logging.info(
            "配台試行ルール: AIキャッシュのみで D→E ポッポ %s 件を完絝（API 呼び出しなし）。",
            n,
        )
        return out
    if not API_KEY:
        return out
    m = len(pend_b)
    if m == 1:
        out[pend_i[0]] = _ai_compile_exclude_rule_logic_to_json(pend_b[0])
        return out

    schema = _exclude_rule_logic_gemini_schema_instructions()
    numbered = "\n".join(f"[{i + 1}] {str(b).strip()}" for i, b in enumerate(pend_b))
    prompt = (
        "あなたは工場の配台システム用です。以下の N 個の「配台試行の説明」を」与ごた順庝でしれずれ JSON ルールに変杛してください。\n\n"
        f"」出力】JSON 配列のみ。先頭は [ で終ゝりは ] 。覝素数は必う {m}（Markdown・説明禁止）。\n"
        f"配列の先頭覝素は [1]」2 番目は [2] … に対応しした。\n\n"
        f"{schema}\n"
        f"」説明文】\n{numbered}\n"
    )
    try:
        ppath = os.path.join(log_dir, "ai_exclude_rule_logic_batch_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("配台試行ルール(ポッポ): プロンプト → %s", ppath)
    except OSError as ex:
        logging.warning("配台試行ルール(ポッポ): プロンプト保存失敗: %s", ex)
    try:
        client = _gemini_client(API_KEY)
        res, gem_model_used = _gemini_generate_content_with_retry(
            client, contents=prompt, log_label="配台不要ルールD→Eバッチ"
        )
        record_gemini_response_usage(res, gem_model_used)
        raw = (_gemini_result_text(res) or "").strip()
        rpath = os.path.join(log_dir, "ai_exclude_rule_logic_batch_last_response.txt")
        try:
            with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                rf.write(raw)
        except OSError:
            pass
        arr = _parse_exclude_rule_json_array_response(raw)
        if not isinstance(arr, list) or len(arr) != m:
            logging.warning(
                "配台試行ルール: ポッポ応答は正常（覝素数 %s」期待 %s）。1 件うつ再試行した。",
                len(arr) if isinstance(arr, list) else None,
                m,
            )
            for j, idx in enumerate(pend_i):
                out[idx] = _ai_compile_exclude_rule_logic_to_json(pend_b[j])
            return out
        cache_dirty = False
        for j, item in enumerate(arr):
            parsed = _validate_exclude_rule_parsed_dict(item)
            out[pend_i[j]] = parsed
            if parsed:
                _cache_put_exclude_rule_de_parsed(ai_cache, pend_b[j], parsed)
                cache_dirty = True
        if cache_dirty:
            save_ai_cache(ai_cache)
        return out
    except Exception as e:
        logging.warning("配台試行ルール: ポッポ Gemini 失敗」坘発にフォールバック: %s", e)
        for j, idx in enumerate(pend_i):
            out[idx] = _ai_compile_exclude_rule_logic_to_json(pend_b[j])
        return out
def _log_exclude_rules_sheet_debug(
    event: str,
    log_prefix: str,
    summary: str,
    details: str = "",
    exc: BaseException | None = None,
) -> None:
    """
    「設定_配台不要工程」の保守処理のイベントログ。

    設定シート処理の追跡を log/exclude_rules_sheet_debug.txt に追記し、execution_log にもタグ付しで出力れる。
    event 例: START, OPEN_OK, OPEN_RETRY, OPEN_FAIL, HEADER_FIX, SYNC_ROWS, OPENPYXL_SAVE_OK, OPENPYXL_SAVE_FAIL,
    OPENPYXL_SAVE_SKIPPED_EXCLUDE_RULES_POLICY, OPENPYXL_RETRY_WAIT, OPENPYXL_VBA_FALLBACK, MATRIX_TSV_WRITTEN,
    XLWINGS_UNAVAILABLE, XLWINGS_ATTACH_FAIL, XLWINGS_SYNC_SKIP, XLWINGS_SYNC_OK, XLWINGS_SYNC_FAIL,
    E_SIDECAR_WRITTEN, E_SIDECAR_APPLIED, FALLBACK_FAIL,
    SKIP_NO_PATH, SKIP_NO_FILE, SKIP_NO_SHEET, DATA_COMPACT
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        f"--- {ts} ---",
        f"event={event}",
        f"phase={log_prefix}",
        f"summary={summary}",
    ]
    if details:
        lines.append(f"details={details}")
    if exc is not None:
        lines.append(f"exception={type(exc).__name__}: {exc}")
        lines.append(traceback.format_exc().rstrip())
    block = "\n".join(lines) + "\n\n"
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(exclude_rules_sheet_debug_log_path, "a", encoding="utf-8", newline="\n") as df:
            df.write(block)
    except OSError as wex:
        logging.warning("exclude_rules_sheet_debug.txt へ書けません: %s", wex)

    tag = "[設定_配台不要工程]"
    msg = f"{tag} {event} | {log_prefix} | {summary}"
    if details:
        msg += f" | {details}"
    if event in (
        "OPEN_FAIL",
        "SAVE_FAIL",
        "COM_MERGE_FAIL",
        "FALLBACK_FAIL",
        "SKIP_NO_PATH",
        "SKIP_NO_FILE",
        "SKIP_NO_SHEET",
        "FATAL",
    ):
        logging.error(msg)
    elif event in (
        "OPEN_RETRY",
        "SAVE_FAIL_HINT",
        "SAVE_RETRY",
        "COM_SYNC_UNAVAILABLE",
        "COM_ATTACH_OPEN_FAIL",
        "XLWINGS_UNAVAILABLE",
        "XLWINGS_ATTACH_FAIL",
        "XLWINGS_SYNC_SKIP",
        "XLWINGS_SYNC_FAIL",
        "E_SIDECAR_WRITTEN",
        "OPENPYXL_SAVE_FAIL",
        "OPENPYXL_VBA_FALLBACK",
    ):
        logging.warning(msg)
    elif event in (
        "COM_MERGE_SKIP",
        "MATRIX_TSV_WRITTEN",
        "OPENPYXL_SAVE_OK",
        "OPENPYXL_RETRY_WAIT",
    ):
        logging.info(msg)
    else:
        logging.info(msg)
def __deleted_xlwings_block_placeholder_do_not_use():
    """ディスクパスと Excel ``Book.fullname``（パス文字列）が同一ファイルを指すか（表記ゆれを多少吸収）。"""
    try:
        fn = str(book_fullname).strip()
    except Exception:
        return False

    def _norm(p: str) -> str:
        p = os.path.normpath(str(p).strip().replace("/", "\\"))
        return os.path.normcase(os.path.abspath(p))

    try:
        if _norm(disk_path) == _norm(fn):
            return True
    except Exception:
        pass
    try:
        return os.path.samefile(disk_path, fn)
    except Exception:
        pass
    try:
        import win32api  # type: ignore

        a = _norm(win32api.GetLongPathName(disk_path))
        b = _norm(win32api.GetLongPathName(fn))
        if a == b:
            return True
    except Exception:
        pass
    try:
        if os.path.basename(_norm(disk_path)).lower() == os.path.basename(_norm(fn)).lower():
            if _norm(os.path.dirname(disk_path)) == _norm(os.path.dirname(fn)):
                return True
    except Exception:
        pass
    return False
def _xlwings_book_path_str(book) -> str | None:
    """Excel 連携 ``Book`` のディスクパス。公式 API は ``fullname``（``full_name`` は存在しない）。"""
    for attr in ("fullname", "full_name"):
        try:
            fn = getattr(book, attr, None)
            if fn is not None:
                s = str(fn).strip()
                if s:
                    return s
        except Exception:
            continue
    return None
def _xlwings_book_matches_path(book, disk_path: str) -> bool:
    fn = _xlwings_book_path_str(book)
    if not fn:
        return False
    return _xlwings_paths_equivalent(disk_path, fn)
def _xlwings_find_book_on_running_instances(abs_path: str):
    """起動中の Excel からパス一致する Book を返す。無ければ None。

    ``for app in xw.apps`` の列挙は環境によって COM 無応答になり得るが、
    ユーザー操作の Excel 上で既にマクロブックを開いているときは再利用した方が安全なため列挙する。
    見つからないときは ``_xlwings_attach_open_macro_workbook`` が非表示の新規 ``xw.App`` で開く。
    """
    if not (abs_path or "").strip():
        return None
    try:
        import xlwings as xw
    except ImportError:
        return None
    try:
        for app in xw.apps:
            try:
                for book in app.books:
                    if _xlwings_book_matches_path(book, abs_path):
                        return book
            except Exception:
                continue
    except Exception:
        return None
    return None
def _xlwings_find_book_in_app(app, abs_path: str):
    """同一 ``Application`` 内の ``Books`` からパス一致する ``Book`` を返す。無ければ None。"""
    if app is None or not (abs_path or "").strip():
        return None
    try:
        for book in app.books:
            if _xlwings_book_matches_path(book, abs_path):
                return book
    except Exception:
        pass
    return None
def _xlwings_same_excel_app(app_a, app_b) -> bool:
    """2 つの連携用 App が同一 Excel プロセスか。"""
    if app_a is None or app_b is None:
        return False
    if app_a is app_b:
        return True
    try:
        return int(app_a.pid) == int(app_b.pid)
    except Exception:
        return False
def _xlwings_attach_two_workbooks_same_app(abs_result: str, abs_input: str):
    """
    図形の Copy/Paste 用に、結果ブックと入力ブックを **同一 Application** 上に揃える。

    戻り値: ``(app, wb_out, wb_in, owns_app, opened_out_here, opened_in_here)`` または失敗時 ``None``。

    - ``owns_app`` … この関数が新規 ``xw.App`` を起動したとき True（終了時に ``quit`` する）。
    - ``opened_*_here`` … 当該ブックを本関数が ``books.open`` したとき True（終了時にそのブックだけ ``close``）。
    """
    try:
        import xlwings as xw
    except ImportError:
        return None
    abs_r = os.path.abspath(abs_result)
    abs_i = os.path.abspath(abs_input)

    bo = _xlwings_find_book_on_running_instances(abs_r)
    bi = _xlwings_find_book_on_running_instances(abs_i)

    if bo is not None and bi is not None and _xlwings_same_excel_app(bo.app, bi.app):
        try:
            bo.app.display_alerts = False
        except Exception:
            pass
        return (bo.app, bo, bi, False, False, False)

    if bo is not None:
        try:
            ap = bo.app
            try:
                ap.display_alerts = False
            except Exception:
                pass
            bi2 = _xlwings_find_book_in_app(ap, abs_i)
            if bi2 is None:
                bi2 = ap.books.open(abs_i, read_only=True, update_links=False)
                return (ap, bo, bi2, False, False, True)
            return (ap, bo, bi2, False, False, False)
        except Exception:
            pass

    if bi is not None:
        try:
            ap = bi.app
            try:
                ap.display_alerts = False
            except Exception:
                pass
            bo2 = _xlwings_find_book_in_app(ap, abs_r)
            if bo2 is None:
                bo2 = ap.books.open(abs_r, update_links=False)
                return (ap, bo2, bi, False, True, False)
            return (ap, bo2, bi, False, False, False)
        except Exception:
            pass

    try:
        ap = xw.App(visible=False, add_book=False)
        ap.display_alerts = False
        wb_o = ap.books.open(abs_r, update_links=False)
        wb_i = ap.books.open(abs_i, read_only=True, update_links=False)
        return (ap, wb_o, wb_i, True, True, True)
    except Exception:
        return None
def _xlwings_release_book_after_mutation(xw_book, info: dict, mutation_ok: bool) -> None:
    """専用起動した Excel は終了れる。実行中 Excel での値 Open したブックは失敗時のみ閉もる。"""
    if xw_book is None:
        return
    mode = info.get("mode", "keep")
    opened_here = bool(info.get("opened_wb_here"))
    if mode == "quit_excel":
        defer_despite = bool(info.get("defer_quit_despite_failure"))
        should_defer = _env_xlwings_defer_quit_on_success() and (
            mutation_ok or (defer_despite and _env_xlwings_defer_after_exclude_sheet_miss())
        )
        if should_defer:
            try:
                _xlwings_register_deferred_hidden_app(xw_book.app)
            except Exception:
                pass
            return
        try:
            xw_book.close()
        except Exception:
            pass
        try:
            xw_book.app.quit()
        except Exception:
            pass
        return
    if opened_here and not mutation_ok:
        try:
            xw_book.close()
        except Exception:
            pass
def _xlwings_attach_open_macro_workbook(macro_wb_path: str, log_prefix: str):
    """
    マクロブックを Excel 連携で取得する（本番・テスト共通）。
    戻り値: (Book, release_info) / 失敗時 None。
    release_info: mode は keep または quit_excel」opened_wb_here は bool。
    """
    try:
        import xlwings as xw  # noqa: F401
    except ImportError:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_UNAVAILABLE",
            log_prefix,
            "Excel 連携モジュールを import できません（requirements.txt の依存を確認）。",
        )
        return None

    abs_path = os.path.abspath(macro_wb_path)

    book_existing = _xlwings_find_book_on_running_instances(abs_path)
    if book_existing is not None:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_BOOK_REUSED",
            log_prefix,
            "起動中の Excel から同一路径のブックを再利用します（xw.apps 列挙）。",
            details=f"path={abs_path}",
        )
        return book_existing, {"mode": "keep", "opened_wb_here": False}

    _log_exclude_rules_sheet_debug(
        "XLWINGS_NEW_APP_FALLBACK",
        log_prefix,
        "起動中に該当ブックは見つかりません。非表示の新規 Excel で開きます。",
        details=f"path={abs_path}",
    )

    try:
        import xlwings as xw

        app = xw.App(visible=False, add_book=False)
        try:
            app.display_alerts = False
        except Exception:
            pass
        book = app.books.open(abs_path, update_links=False)
        return book, {"mode": "quit_excel", "opened_wb_here": True}
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_ATTACH_FAIL",
            log_prefix,
            "Excel でブックを開けませんでした。",
            details=f"path={abs_path}",
            exc=ex,
        )
        return None
def _xlwings_attach_workbook_for_tests(
    book_path: str,
    label: str,
    *,
    allow_dispatch_open: bool = False,
):
    """
    検証スクリプト用: 起動中ブックを優先し、必須なら表示付し Excel で開し。
    戻り値: (Book, info, 説明文字列) または None。
    """
    abs_path = os.path.abspath(book_path)
    if not allow_dispatch_open:
        return None
    try:
        import xlwings as xw

        app = xw.App(visible=True, add_book=False)
        try:
            app.display_alerts = False
        except Exception:
            pass
        book = app.books.open(abs_path, update_links=False)
        return book, {"mode": "keep", "opened_wb_here": True}, f"{label}:dispatch-open"
    except Exception:
        return None
def _env_xlw_suspend_auto_calculation() -> bool:
    """環境変数 XLWINGS_SUSPEND_AUTO_CALCULATION（既定 1）が off のとき、Calculation は変更しない。"""
    raw = (os.environ.get("XLWINGS_SUSPEND_AUTO_CALCULATION", "1") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off", "n", "いいえ", "無効", "×")
def _xlwings_app_save_perf_state_push(app):
    """VBA スプラッシュポーリングと競合しやすいため、同期・保存の短時間 Excel を静かにする。
    自動計算は手動（xlCalculationManual）へ切り替え、終了時に _xlwings_app_save_perf_state_pop で復帰する。
    XLWINGS_SUSPEND_AUTO_CALCULATION=0 等のときは Calculation の取得・変更を行わない。
    """
    suspend_calc = _env_xlw_suspend_auto_calculation()
    snap: dict = {"suspend_auto_calculation": suspend_calc}
    for attr in ("screen_updating", "enable_events"):
        try:
            snap[attr] = getattr(app, attr)
        except Exception:
            snap[attr] = None
    if suspend_calc:
        try:
            snap["calculation"] = getattr(app, "calculation")
        except Exception:
            snap["calculation"] = None
            try:
                snap["calculation"] = app.api.Calculation
            except Exception:
                pass
    else:
        snap["calculation"] = None
    try:
        app.screen_updating = False
    except Exception:
        pass
    if suspend_calc:
        try:
            app.calculation = "manual"
        except Exception:
            try:
                app.api.Calculation = -4135  # xlCalculationManual
            except Exception:
                pass
    try:
        app.enable_events = False
    except Exception:
        pass
    return snap
def _xlwings_app_save_perf_state_pop(app, snap):
    if not snap:
        return
    try:
        prev_ev = snap.get("enable_events")
        if prev_ev is not None:
            setattr(app, "enable_events", prev_ev)
    except Exception:
        pass
    if snap.get("suspend_auto_calculation"):
        prev_calc = snap.get("calculation")
        if prev_calc is not None:
            try:
                setattr(app, "calculation", prev_calc)
            except Exception:
                try:
                    app.api.Calculation = prev_calc
                except Exception:
                    pass
    try:
        prev_su = snap.get("screen_updating")
        if prev_su is not None:
            setattr(app, "screen_updating", prev_su)
    except Exception:
        pass
def _xlwings_sync_exclude_rules_sheet_from_openpyxl(
    wb_path: str, ws_oxl, log_prefix: str
) -> bool:
    """
    openpyxl で保存できないとき、Excel 上の「設定_配台不要工程」A:E をメモリ上の値で上書きし Save。

    表示中シートに対れる一括 .value の値てと」スプラッシュ＋ポーリング（D3=true）下で
    Range 代入は数分かかる計測はあり得る。同期中のみシートを一時非表示にし api.Value2 で書き。
    """
    global _exclude_rules_effective_read_path

    attached = _xlwings_attach_open_macro_workbook(wb_path, log_prefix)
    if attached is None:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_SYNC_SKIP",
            log_prefix,
            "Excel にブックを接続できず A:E 同期をスキップ。",
            details=f"path={wb_path}",
        )
        return False

    xw_book, info = attached
    ok = False
    try:
        try:
            xw_book.app.display_alerts = False
        except Exception:
            pass
        # 全シート名を列挙れるとシート数分の COM 往復になり」D3=true 時は VBA ポーリングと競坈して
        # 1 シート数秒〜坝数秒かかることはある（計測で 40 シート≈213s）。坝剝で直接解決れる。
        try:
            sht = xw_book.sheets[EXCLUDE_RULES_SHEET_NAME]
        except Exception:
            try:
                info["defer_quit_despite_failure"] = True
            except Exception:
                pass
            _log_exclude_rules_sheet_debug(
                "XLWINGS_SYNC_SKIP",
                log_prefix,
                f"Excel 側にシート「{EXCLUDE_RULES_SHEET_NAME}」はありません。",
                details=f"path={wb_path}",
            )
            return False
        max_r = max(1, int(ws_oxl.max_row or 1))
        ncols = EXCLUDE_RULES_SHEET_COM_SYNC_MAX_COL
        data = [
            [ws_oxl.cell(row=r, column=c).value for c in range(1, ncols + 1)]
            for r in range(1, max_r + 1)
        ]
        _perf_snap = _xlwings_app_save_perf_state_push(xw_book.app)
        rng = sht.range((1, 1)).resize(len(data), ncols)
        hid_sheet_for_write = False
        try:
            try:
                if int(sht.api.Visible) == -1:  # xlSheetVisible
                    sht.api.Visible = 0  # xlSheetHidden（同期中の値。再杝画・ウィンドウ更新負蝷を抑ごる）
                    hid_sheet_for_write = True
            except Exception:
                pass
            try:
                rng.api.Value2 = data
            except Exception:
                rng.value = data
            xw_book.save()
        finally:
            if hid_sheet_for_write:
                try:
                    sht.api.Visible = -1
                except Exception:
                    pass
            _xlwings_app_save_perf_state_pop(xw_book.app, _perf_snap)
        ok = True
        _exclude_rules_effective_read_path = wb_path
        _clear_exclude_rules_e_apply_files()
        _log_exclude_rules_sheet_debug(
            "XLWINGS_SYNC_OK",
            log_prefix,
            "Excel 経由で設定シート A〜E を同期しブックを保存しました。",
            details=f"path={wb_path} rows={max_r}",
        )
        logging.info(
            "%s: 設定シートを Excel でマクロブックに保存しました（A〜E）。",
            log_prefix,
        )
        return True
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_SYNC_FAIL",
            log_prefix,
            "Excel での A:E 同期または Save に失敗しました。",
            details=f"path={wb_path}",
            exc=ex,
        )
        return False
    finally:
        _xlwings_release_book_after_mutation(xw_book, info, ok)
def _persist_exclude_rules_workbook(_wb, wb_path: str, ws, log_prefix: str) -> bool:
    """
    設定シートのディスク反映。openpyxl でマクロ付きブックを保存する（keep_vba）。

    失敗時は log に行列 TSV を出し、VBA「設定_配台不要工程_AからE_TSVから反映」で反映できる。
    """
    global _exclude_rules_effective_read_path

    abs_p = os.path.abspath(str(wb_path or "").strip())
    if not abs_p:
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_PATH", log_prefix, "設定シート保存: パスが空です。", details=""
        )
        return False
    keep_vba = abs_p.lower().endswith(".xlsm")
    try:
        _wb.save(abs_p)
        _exclude_rules_effective_read_path = abs_p
        _clear_exclude_rules_e_apply_files()
        _log_exclude_rules_sheet_debug(
            "OPENPYXL_SAVE_OK",
            log_prefix,
            "openpyxl で設定シートを含むブックを保存しました。",
            details=f"path={abs_p}",
        )
        logging.info(
            "%s: 「%s」を openpyxl で保存しました（設定_配台不要工程）。",
            log_prefix,
            os.path.basename(abs_p),
        )
        return True
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "OPENPYXL_SAVE_FAIL",
            log_prefix,
            "openpyxl でのブック保存に失敗しました。",
            details=f"path={abs_p}",
            exc=ex,
        )
        logging.warning(
            "%s: 設定シートの openpyxl 保存に失敗: %s", log_prefix, ex
        )
    if _write_exclude_rules_matrix_vba_tsv(abs_p, ws, log_prefix):
        logging.warning(
            "%s: 設定シートを log\\%s に出力しました。"
            " Excel でマクロ「設定_配台不要工程_AからE_TSVから反映」を実行してください。",
            log_prefix,
            EXCLUDE_RULES_MATRIX_VBA_FILENAME,
        )

    _log_exclude_rules_sheet_debug(
        "OPENPYXL_VBA_FALLBACK",
        log_prefix,
        "保存に失敗したため、VBA 用行列 TSV を出力しました（必要なら手動反映）。",
        details=f"path={abs_p}",
    )
    return False
def _exclude_rules_e_sidecar_path() -> str:
    path = os.path.join(json_data_dir, EXCLUDE_RULES_E_SIDECAR_FILENAME)
    legacy = os.path.join(log_dir, EXCLUDE_RULES_E_SIDECAR_FILENAME)
    if os.path.isfile(legacy) and not os.path.isfile(path):
        try:
            shutil.move(legacy, path)
        except OSError:
            pass
    return path
def _exclude_rules_e_vba_tsv_path() -> str:
    return os.path.join(log_dir, EXCLUDE_RULES_E_VBA_TSV_FILENAME)
def _exclude_rules_matrix_vba_path() -> str:
    return os.path.join(log_dir, EXCLUDE_RULES_MATRIX_VBA_FILENAME)
def _serialize_cell_for_matrix_tsv(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    return str(val)
def _write_exclude_rules_matrix_vba_tsv(
    wb_path: str, ws, log_prefix: str
) -> bool:
    """VBA 用: 設定シート 1 行目〜 max_row の A〜E を Base64(UTF-8) 付し TSV で出力れる。"""
    max_r = max(1, int(ws.max_row or 1))
    lines = [
        "v1",
        "workbook\t" + os.path.abspath(wb_path),
        "sheet\t" + EXCLUDE_RULES_SHEET_NAME,
        "ncols\t5",
        "---",
    ]
    for r in range(1, max_r + 1):
        parts: list[str] = [str(r)]
        for c in range(1, 6):
            s = _serialize_cell_for_matrix_tsv(ws.cell(row=r, column=c).value)
            parts.append(base64.b64encode(s.encode("utf-8")).decode("ascii"))
        lines.append("\t".join(parts))
    path = _exclude_rules_matrix_vba_path()
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines) + "\n")
        _log_exclude_rules_sheet_debug(
            "MATRIX_TSV_WRITTEN",
            log_prefix,
            "設定シート A〜E を VBA 反映用 TSV に書き出しました（Excel 保存失敗時のフォールバック）。",
            details=f"path={path} rows={max_r}",
        )
        return True
    except OSError as ex:
        logging.warning("%s: 行列 VBA 用 TSV を書けません: %s", log_prefix, ex)
        return False
def _build_exclude_rules_list_from_openpyxl_ws(
    ws, c_proc: int, c_mach: int, c_flag: int, c_e: int
) -> list[dict]:
    """openpyxl 上の設定シートから _load_exclude_rules_from_workbook と同形のリストを構築。"""
    rules: list[dict] = []
    max_r = int(ws.max_row or 1)
    for r in range(2, max_r + 1):
        pv = ws.cell(row=r, column=c_proc).value
        proc = (
            ""
            if pv is None or (isinstance(pv, float) and pd.isna(pv))
            else str(pv).strip()
        )
        if not proc:
            continue
        mv = ws.cell(row=r, column=c_mach).value
        mach = (
            ""
            if mv is None or (isinstance(mv, float) and pd.isna(mv))
            else str(mv).strip()
        )
        cv = ws.cell(row=r, column=c_flag).value
        ev = ws.cell(row=r, column=c_e).value
        parsed = _parse_exclude_rule_json_cell(ev)
        rules.append(
            {"proc": proc, "mach": mach, "c_val": cv, "parsed": parsed}
        )
    return rules
def _set_exclude_rules_snapshot_from_ws(
    wb_path: str, ws, c_proc: int, c_mach: int, c_flag: int, c_e: int
) -> None:
    global _exclude_rules_rules_snapshot, _exclude_rules_snapshot_wb
    _exclude_rules_rules_snapshot = _build_exclude_rules_list_from_openpyxl_ws(
        ws, c_proc, c_mach, c_flag, c_e
    )
    _exclude_rules_snapshot_wb = os.path.normcase(os.path.abspath(wb_path))
def _clear_exclude_rules_e_apply_files() -> None:
    for p in (
        os.path.join(json_data_dir, EXCLUDE_RULES_E_SIDECAR_FILENAME),
        os.path.join(log_dir, EXCLUDE_RULES_E_SIDECAR_FILENAME),
    ):
        if os.path.isfile(p):
            try:
                os.remove(p)
            except OSError:
                pass
    for rel in (EXCLUDE_RULES_E_VBA_TSV_FILENAME, EXCLUDE_RULES_MATRIX_VBA_FILENAME):
        p = os.path.join(log_dir, rel)
        if os.path.isfile(p):
            try:
                os.remove(p)
            except OSError:
                pass
def _write_exclude_rules_e_vba_tsv_from_cells(
    wb_path: str, c_e: int, cells: dict[str, str], log_prefix: str
) -> None:
    """VBA 用: 行番坷と Base64(UTF-8) セル文字列の TSV。"""
    lines = [
        "v1",
        "workbook\t" + os.path.abspath(wb_path),
        "sheet\t" + EXCLUDE_RULES_SHEET_NAME,
        "column_e\t" + str(int(c_e)),
        "---",
    ]
    for rk in sorted(cells.keys(), key=lambda x: int(x)):
        s = cells[rk]
        b64 = base64.b64encode(s.encode("utf-8")).decode("ascii")
        lines.append(rk + "\t" + b64)
    path_tsv = _exclude_rules_e_vba_tsv_path()
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path_tsv, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines) + "\n")
        _log_exclude_rules_sheet_debug(
            "E_VBA_TSV_WRITTEN",
            log_prefix,
            "E 列を VBA 反映用 TSV に書き出しました（保存失敗時のフォールバック用）。",
            details=f"path={path_tsv} cells={len(cells)}",
        )
    except OSError as ex:
        logging.warning("%s: E 列 VBA 用 TSV を書けません: %s", log_prefix, ex)
def _write_exclude_rules_e_apply_artifacts(
    wb_path: str, ws, c_e: int, log_prefix: str
) -> None:
    """
    E 列（非空）を JSON サイドカードと VBA 用 TSV に書き。空なら両ファイルを削除。
    Python 次回起動時の E 復元用 JSON と」マクロからの E 書込み用 TSV。
    """
    cells: dict[str, str] = {}
    max_r = int(ws.max_row or 1)
    for r in range(2, max_r + 1):
        ev = ws.cell(row=r, column=c_e).value
        if _cell_is_blank_for_rule(ev):
            continue
        s = str(ev).strip() if ev is not None else ""
        if not s:
            continue
        cells[str(r)] = s
    if not cells:
        _clear_exclude_rules_e_apply_files()
        return
    payload = {
        "version": 1,
        "workbook": os.path.abspath(wb_path),
        "sheet": EXCLUDE_RULES_SHEET_NAME,
        "column_e": c_e,
        "cells": cells,
    }
    path_sc = _exclude_rules_e_sidecar_path()
    try:
        os.makedirs(json_data_dir, exist_ok=True)
        with open(path_sc, "w", encoding="utf-8", newline="\n") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except OSError as ex:
        logging.warning("%s: E 列 JSON を書けません: %s", log_prefix, ex)
    _write_exclude_rules_e_vba_tsv_from_cells(wb_path, c_e, cells, log_prefix)
    _log_exclude_rules_sheet_debug(
        "E_APPLY_FILES_WRITTEN",
        log_prefix,
        "E 列を JSON と VBA 用 TSV に書き出しました（マクロで E 列を反映後」ファイル削除）。",
        details=f"cells={len(cells)}",
    )
def _try_apply_pending_exclude_rules_e_column(
    wb_path: str, ws, c_e: int, log_prefix: str
) -> int:
    """
    剝回保存に失敗したとし書き出した JSON から E 列を復元れる。
    ブックパスは一致しなけれみ何もしない。適用後はサイドカードを削除する。
    """
    path_sc = _exclude_rules_e_sidecar_path()
    if not os.path.isfile(path_sc):
        return 0
    try:
        with open(path_sc, "r", encoding="utf-8") as f:
            payload = json.load(f)
    except Exception:
        return 0
    if int(payload.get("version") or 0) != 1:
        return 0
    target = os.path.normcase(os.path.abspath(wb_path))
    if os.path.normcase(str(payload.get("workbook") or "")) != target:
        return 0
    if str(payload.get("sheet") or "") != EXCLUDE_RULES_SHEET_NAME:
        return 0
    cells = payload.get("cells")
    if not isinstance(cells, dict):
        return 0
    n = 0
    for rk, val in cells.items():
        try:
            ri = int(rk)
        except (TypeError, ValueError):
            continue
        if ri < 2:
            continue
        if isinstance(val, dict):
            sval = json.dumps(val, ensure_ascii=False)
        else:
            sval = "" if val is None else str(val).strip()
        if not sval:
            continue
        ws.cell(row=ri, column=c_e, value=sval)
        n += 1
    try:
        os.remove(path_sc)
    except OSError:
        pass
    if n:
        _log_exclude_rules_sheet_debug(
            "E_SIDECAR_APPLIED",
            log_prefix,
            f"未保存だった E 列をサイドカードから {n} セル復元しました。",
            details=path_sc,
        )
        logging.info(
            "%s: %s の内容をシートのロジック式列へ適用しました（続けて保存を試みした）。",
            log_prefix,
            path_sc,
        )
    return n
def _read_exclude_rules_d_cells_data_only_for_rows(
    wb_path: str, rows: list[int], c_d: int
) -> dict[int, object]:
    """
    D 列は数弝のとき」openpyxl の通常読込では '=...' しか取れない。
    data_only=True でキャッシュ値を読む（Excel は一度でも保存・計算済みのブックで有効）。
    """
    out: dict[int, object] = {}
    if not rows or not os.path.isfile(wb_path):
        return out
    if _workbook_should_skip_openpyxl_io(wb_path):
        return out
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wbro = None
    try:
        wbro = load_workbook(
            wb_path,
            read_only=True,
            data_only=True,
            keep_vba=keep_vba,
        )
    except Exception:
        return out
    try:
        if EXCLUDE_RULES_SHEET_NAME not in wbro.sheetnames:
            return out
        wsro = wbro[EXCLUDE_RULES_SHEET_NAME]
        for r in rows:
            if r < 2:
                continue
            try:
                out[r] = wsro.cell(row=r, column=c_d).value
            except Exception:
                pass
    finally:
        if wbro is not None:
            try:
                wbro.close()
            except Exception:
                pass
    return out
def run_exclude_rules_sheet_maintenance(
    wb_path: str,
    pairs: list[tuple[str, str]],
    log_prefix: str,
    *,
    compile_exclude_rules_d_to_e_with_ai: bool = True,
) -> None:
    """
    「設定_配台不要工程」の行同期・（任意で）D→E の AI 補完・ディスク反映（マクロブックは **openpyxl save なし**、Excel で A〜E 同期→Save）。

    ``compile_exclude_rules_d_to_e_with_ai=False`` のときは D 列→E 列（ロジック式 JSON）の
    Gemini 補完のみスキップする（行同期・空行詰め・退避 E の復元・保存は従来どおり）。
    段階2の ``load_planning_tasks_df`` 経路では False を渡す。

    Excel 経由でも保存に失敗した場合は ``log/exclude_rules_matrix_vba.tsv`` を残し、マクロ
    ``設定_配台不要工程_AからE_TSVから反映`` で A〜E を反映れる。
    併せで従来どおり E 列のみの ``exclude_rules_e_column_vba.tsv`` も出力され得る（行列 TSV 優先で反映後は削除）。
    保存成功時は TSV/JSON は削除される。

    ``json/exclude_rules_e_column_pending.json`` は Python 次回起動時の E 列復元用。
    シートの新規作成と 1 行目見出しは VBA「設定_配台不要工程_シートを確保」。
    """
    import hashlib

    needs_disk_sync = False
    _st_added = 0
    _st_removed_empty = 0
    _st_pending_rows = 0
    _st_ai_filled = 0
    _st_header_changed = False
    _st_sidecar_applied = 0

    # 直前と同じ入力・同じブック状態なら、openpyxl でブックを開く前に保守全体をスキップできる場合がある。
    # （book open が重いため。スキップ条件は「前回は変更なし」で、今回もブックの mtime/size と pairs が同一、かつ sidecar 無し）
    try:
        _wb_abs = os.path.abspath(wb_path) if wb_path else ""
        _st = os.stat(_wb_abs) if _wb_abs and os.path.exists(_wb_abs) else None
        _sig = {
            "wb": os.path.normcase(_wb_abs),
            "mtime_ns": int(getattr(_st, "st_mtime_ns", 0) or 0) if _st else 0,
            "size": int(getattr(_st, "st_size", 0) or 0) if _st else 0,
            "ai": bool(compile_exclude_rules_d_to_e_with_ai),
            "pairs_hash": hashlib.sha256(
                ("\n".join(f"{p}\t{m}" for p, m in (pairs or []))).encode("utf-8", "ignore")
            ).hexdigest(),
        }
        _cache_path = os.path.join(json_data_dir, "exclude_rules_maintenance_cache.json")
        _have_sidecar = os.path.isfile(_exclude_rules_e_sidecar_path())
        if os.path.isfile(_cache_path) and (not _have_sidecar):
            try:
                with open(_cache_path, "r", encoding="utf-8") as f:
                    _cache = json.load(f)
            except Exception:
                _cache = None
            if isinstance(_cache, dict) and _cache.get("sig") == _sig and _cache.get("can_skip") is True:
                _log_exclude_rules_sheet_debug(
                    "SKIP_CACHE_NO_CHANGES",
                    log_prefix,
                    "前回と同一シグネチャかつ変更無しだったため、設定シート保守（openpyxl open 含む）をスキップしました。",
                    details=f"path={wb_path} pairs={len(pairs)} ai_d_to_e={compile_exclude_rules_d_to_e_with_ai}",
                )
                return
    except Exception:
        pass

    if not wb_path:
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_PATH",
            log_prefix,
            "ブックパスが空のため、設定シート処理をしません。",
        )
        return
    if not os.path.exists(wb_path):
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_FILE",
            log_prefix,
            "ブックは存在しません。",
            details=f"path={wb_path}",
        )
        return

    _log_exclude_rules_sheet_debug(
        "START",
        log_prefix,
        "設定シート保守開始",
        details=f"path={wb_path} pairs={len(pairs)} ai_d_to_e={compile_exclude_rules_d_to_e_with_ai}",
    )
    global _exclude_rules_effective_read_path
    _exclude_rules_effective_read_path = None

    if _workbook_should_skip_openpyxl_io(wb_path):
        _log_exclude_rules_sheet_debug(
            "SKIP_OPENPYXL_INCOMPATIBLE_BOOK",
            log_prefix,
            f"ブックに「{OPENPYXL_INCOMPATIBLE_SHEET_MARKER}」は含まれるため、openpyxl による設定シート保守は行いません。",
            details=f"path={wb_path}",
        )
        logging.warning(
            "%s: 「%s」併有のため、「%s」の openpyxl 保守をスキップしました（Excel で直接編集してください）。",
            log_prefix,
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
            EXCLUDE_RULES_SHEET_NAME,
        )
        return

    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = None
    try:
        wb = load_workbook(wb_path, keep_vba=keep_vba, read_only=False, data_only=False)
    except Exception as e1:
        if keep_vba:
            _log_exclude_rules_sheet_debug(
                "OPEN_RETRY",
                log_prefix,
                "keep_vba=True でブックを開けう keep_vba=False で再試行した（マクロは失ゝれる可能性）。",
                exc=e1,
            )
            try:
                wb = load_workbook(wb_path, keep_vba=False, read_only=False, data_only=False)
            except Exception as e2:
                _log_exclude_rules_sheet_debug(
                    "OPEN_FAIL",
                    log_prefix,
                    "ブックを開きません。シートは作成・保存されません。",
                    details=f"path={wb_path}",
                    exc=e2,
                )
                return
        else:
            _log_exclude_rules_sheet_debug(
                "OPEN_FAIL",
                log_prefix,
                "ブックを開きません。シートは作成・保存されません。",
                details=f"path={wb_path}",
                exc=e1,
            )
            return

    _log_exclude_rules_sheet_debug(
        "OPEN_OK",
        log_prefix,
        "ブックを開しました。",
        details=f"keep_vba={keep_vba} sheets={len(wb.sheetnames)}",
    )

    try:
        if EXCLUDE_RULES_SHEET_NAME not in wb.sheetnames:
            _log_exclude_rules_sheet_debug(
                "SKIP_NO_SHEET",
                log_prefix,
                "シートはありません。VBA の「設定_配台不要工程_シートを確保」を実行れるか」段階1/2 をマクロから起動してください。",
                details=f"path={wb_path}",
            )
            logging.error(
                "%s: 「%s」はありません。Python ではシートを作成しません。",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
            )
            return

        ws = wb[EXCLUDE_RULES_SHEET_NAME]
        hm_before = _exclude_rules_sheet_header_map(ws)
        c_proc, c_mach, c_flag, c_d, c_e = _ensure_exclude_rules_sheet_headers_and_columns(
            ws, log_prefix
        )
        hm_after = _exclude_rules_sheet_header_map(ws)
        header_changed = tuple(hm_before.get(x) for x in (
            EXCLUDE_RULE_COL_PROCESS,
            EXCLUDE_RULE_COL_MACHINE,
            EXCLUDE_RULE_COL_FLAG,
            EXCLUDE_RULE_COL_LOGIC_JA,
            EXCLUDE_RULE_COL_LOGIC_JSON,
        )) != tuple(hm_after.get(x) for x in (
            EXCLUDE_RULE_COL_PROCESS,
            EXCLUDE_RULE_COL_MACHINE,
            EXCLUDE_RULE_COL_FLAG,
            EXCLUDE_RULE_COL_LOGIC_JA,
            EXCLUDE_RULE_COL_LOGIC_JSON,
        ))
        _st_header_changed = bool(header_changed)
        if header_changed:
            needs_disk_sync = True
            _log_exclude_rules_sheet_debug(
                "HEADER_FIX",
                log_prefix,
                "1行目に標準見出しを書き込みました（空シート・列名厳密一致の補正）。",
                details=f"cols=({c_proc},{c_mach},{c_flag},{c_d},{c_e})",
            )

        # 剝回ブック保存に失敗したとし退避した E 列を」先にワークシートへ戻れ（続し保存でディスクへ載る）
        n_e_sidecar = int(_try_apply_pending_exclude_rules_e_column(wb_path, ws, c_e, log_prefix) or 0)
        _st_sidecar_applied = n_e_sidecar
        if n_e_sidecar > 0:
            needs_disk_sync = True

        existing_keys: set[tuple[str, str]] = set()
        max_r = max(2, int(ws.max_row or 2))
        for r in range(2, max_r + 1):
            pv = ws.cell(row=r, column=c_proc).value
            mv = ws.cell(row=r, column=c_mach).value
            p = str(pv).strip() if pv is not None and not (isinstance(pv, float) and pd.isna(pv)) else ""
            m = str(mv).strip() if mv is not None and not (isinstance(mv, float) and pd.isna(mv)) else ""
            if not p:
                continue
            existing_keys.add(
                (_normalize_process_name_for_rule_match(p), _normalize_equipment_match_key(m))
            )

        added = 0
        for p, m in pairs:
            key = (_normalize_process_name_for_rule_match(p), _normalize_equipment_match_key(m))
            if key in existing_keys:
                continue
            ws.append([p, m, None, None, None])
            existing_keys.add(key)
            added += 1
        _st_added = int(added)
        if added:
            needs_disk_sync = True
            _log_exclude_rules_sheet_debug(
                "SYNC_ROWS",
                log_prefix,
                f"工程+機械の行を {added} 件追加しました。",
            )
            logging.info(
                "%s: 「%s」に工程+機械の組み合わせを %s 行追加しました。",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
                added,
            )

        # 加工計画からペアは1件も坖れう」シートにもデータ行は無いとしは例行のみ（従来の新規シート相当）
        if added == 0 and not existing_keys:
            needs_disk_sync = True
            ws.append(["梱包", "", "yes", "", ""])
            existing_keys.add(
                (_normalize_process_name_for_rule_match("梱包"), _normalize_equipment_match_key(""))
            )
            _log_exclude_rules_sheet_debug(
                "EXAMPLE_ROW",
                log_prefix,
                "データ行はなかったため、例（梱包=yes）を1行追加。",
            )
            logging.info(
                "%s: 「%s」にデータ行はなかったため、例（梱包=yes）を1行追加しました。",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
            )

        # 空行詰ゝは AI より先に行ご（後から詰ゝると」書き込んて行番坷と画面上の行はうれる）
        n_kept, n_removed_empty = _compact_exclude_rules_data_rows(
            ws, c_proc, c_mach, c_flag, c_d, c_e, log_prefix
        )
        _st_removed_empty = int(n_removed_empty or 0)
        if n_removed_empty:
            needs_disk_sync = True
            _log_exclude_rules_sheet_debug(
                "DATA_COMPACT",
                log_prefix,
                "空行を削除してデータ行を詰ゝました（並よ順は維挝）。AI 補完より剝。",
                details=f"rows={n_kept} removed_empty={n_removed_empty}",
            )

        if compile_exclude_rules_d_to_e_with_ai:
            max_r = int(ws.max_row or 1)
            pending_rows: list[int] = []
            for r in range(2, max_r + 1):
                dv = ws.cell(row=r, column=c_d).value
                ev = ws.cell(row=r, column=c_e).value
                # C 列の有無に関係なく、D に説明があり E が空なら D→E を試す
                if _cell_is_blank_for_rule(dv):
                    continue
                if not _cell_is_blank_for_rule(ev):
                    continue
                pending_rows.append(r)
            _st_pending_rows = int(len(pending_rows))

            # D が数式のときは通常読込では '=...' だけ取れる。data_only でキャッシュ表示値を補う。
            formula_rows = [
                r
                for r in pending_rows
                if isinstance(ws.cell(row=r, column=c_d).value, str)
                and str(ws.cell(row=r, column=c_d).value).strip().startswith("=")
            ]
            d_cached = (
                _read_exclude_rules_d_cells_data_only_for_rows(wb_path, formula_rows, c_d)
                if formula_rows
                else {}
            )
            pending_texts: list[str] = []
            filtered_rows: list[int] = []
            for r in pending_rows:
                dv = ws.cell(row=r, column=c_d).value
                blob = (
                    ""
                    if dv is None or (isinstance(dv, float) and pd.isna(dv))
                    else str(dv).strip()
                )
                if blob.startswith("="):
                    alt = d_cached.get(r)
                    if alt is not None and not (isinstance(alt, float) and pd.isna(alt)):
                        blob = str(alt).strip()
                    else:
                        logging.warning(
                            "%s: 「%s」%s 行目の D 列が数式で、キャッシュ値を読めませんでした（Excel で一度保存するか D を値にしてください）。",
                            log_prefix,
                            EXCLUDE_RULES_SHEET_NAME,
                            r,
                        )
                        continue
                if _cell_is_blank_for_rule(blob):
                    continue
                filtered_rows.append(r)
                pending_texts.append(blob)
            pending_rows = filtered_rows

            ai_filled = 0
            ai_e_cell_addrs: list[str] = []
            if pending_texts:
                parsed_list = _ai_compile_exclude_rule_logics_batch(pending_texts)
                for r, parsed in zip(pending_rows, parsed_list):
                    if not parsed:
                        logging.warning(
                            "%s: 「%s」%s 行目の D 列を JSON にできませんでした（APIキー・応答を確認）。",
                            log_prefix,
                            EXCLUDE_RULES_SHEET_NAME,
                            r,
                        )
                        continue
                    jstr = json.dumps(parsed, ensure_ascii=False)
                    ws.cell(row=r, column=c_e, value=jstr)
                    cell_addr = f"{get_column_letter(c_e)}{r}"
                    ai_e_cell_addrs.append(cell_addr)
                    preview = jstr if len(jstr) <= 160 else (jstr[:160] + "…")
                    logging.info(
                        "%s: 「%s」ロジック式列「%s」セル %s に JSON を書き込み: %s",
                        log_prefix,
                        EXCLUDE_RULES_SHEET_NAME,
                        EXCLUDE_RULE_COL_LOGIC_JSON,
                        cell_addr,
                        preview,
                    )
                    ai_filled += 1
            _st_ai_filled = int(ai_filled or 0)
            if ai_filled:
                needs_disk_sync = True
                _log_exclude_rules_sheet_debug(
                    "AI_E_FILLED",
                    log_prefix,
                    f"D→E の AI 補完を {ai_filled} 行実施。",
                    details="cells=" + ",".join(ai_e_cell_addrs),
                )
                logging.info(
                    "%s: 「%s」で D→E の AI 補完を %s 行（セル: %s）。",
                    log_prefix,
                    EXCLUDE_RULES_SHEET_NAME,
                    ai_filled,
                    ",".join(ai_e_cell_addrs),
                )
        else:
            _log_exclude_rules_sheet_debug(
                "SKIP_AI_D_TO_E",
                log_prefix,
                "D→E の AI 補完をスキップ（呼び出し側指定）。",
            )

        _er_test = os.environ.get("EXCLUDE_RULES_TEST_E1234", "").strip().lower()
        if _er_test in ("1", "yes", "true"):
            try:
                _er_row = int(os.environ.get("EXCLUDE_RULES_TEST_E1234_ROW", "9") or "9")
            except ValueError:
                _er_row = 9
            if _er_row < 2:
                _er_row = 9
            ws.cell(row=_er_row, column=c_e, value="1234")
            needs_disk_sync = True
            _e_addr = f"{get_column_letter(c_e)}{_er_row}"
            _log_exclude_rules_sheet_debug(
                "TEST_E1234",
                log_prefix,
                f'E列 {_e_addr} にテストで "1234" を書き込み',
                details=f"row={_er_row}",
            )
            logging.warning(
                '%s: 」テスト】%s に "1234" を書き込み（EXCLUDE_RULES_TEST_E1234）。',
                log_prefix,
                _e_addr,
            )

        _set_exclude_rules_snapshot_from_ws(
            wb_path, ws, c_proc, c_mach, c_flag, c_e
        )
        _write_exclude_rules_e_apply_artifacts(wb_path, ws, c_e, log_prefix)
        if needs_disk_sync:
            persisted = _persist_exclude_rules_workbook(wb, wb_path, ws, log_prefix)
        else:
            persisted = True
            _log_exclude_rules_sheet_debug(
                "SKIP_XLWINGS_SYNC_NO_CHANGES",
                log_prefix,
                "設定シート（openpyxl側）に変更が無いため、Excel 経由の同期→Save をスキップしました。",
                details=f"path={wb_path}",
            )
        if needs_disk_sync and (not persisted):
            logging.warning(
                "%s: 設定シートの Excel 保存に失敗しました。"
                " log の行列 TSV をマクロ「設定_配台不要工程_AからE_TSVから反映」"
                "または E 列のみ「設定_配台不要工程_E列_TSVから反映」で反映してください。",
                log_prefix,
            )

        # 次回スキップ用キャッシュ（「変更なし」で成功した場合のみ can_skip=True）
        try:
            _wb_abs = os.path.abspath(wb_path) if wb_path else ""
            _st = os.stat(_wb_abs) if _wb_abs and os.path.exists(_wb_abs) else None
            _sig2 = {
                "wb": os.path.normcase(_wb_abs),
                "mtime_ns": int(getattr(_st, "st_mtime_ns", 0) or 0) if _st else 0,
                "size": int(getattr(_st, "st_size", 0) or 0) if _st else 0,
                "ai": bool(compile_exclude_rules_d_to_e_with_ai),
                "pairs_hash": hashlib.sha256(
                    ("\n".join(f"{p}\t{m}" for p, m in (pairs or []))).encode("utf-8", "ignore")
                ).hexdigest(),
            }
            os.makedirs(json_data_dir, exist_ok=True)
            with open(
                os.path.join(json_data_dir, "exclude_rules_maintenance_cache.json"),
                "w",
                encoding="utf-8",
                newline="\n",
            ) as f:
                json.dump(
                    {
                        "sig": _sig2,
                        "can_skip": (not needs_disk_sync),
                        "stats": {
                            "added": _st_added,
                            "removed_empty": _st_removed_empty,
                            "pending_rows": _st_pending_rows,
                            "ai_filled": _st_ai_filled,
                            "header_changed": bool(_st_header_changed),
                            "sidecar_applied": _st_sidecar_applied,
                        },
                    },
                    f,
                    ensure_ascii=False,
                    indent=2,
                )
        except Exception:
            pass
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "FATAL",
            log_prefix,
            "設定シート処理中に未杕杉例外は発生しました。",
            exc=ex,
        )
        logging.exception("%s: 設定_配台不要工程の処理で例外", log_prefix)
    finally:
        if wb is not None:
            wb.close()
            _log_exclude_rules_sheet_debug("CLOSED", log_prefix, "ブックをクローズしました。")
def _resolve_exclude_rules_workbook_path_for_read(wb_path: str) -> str:
    """直後の保守で実効パスは変ゝったとし（通常は保存成功後の元ブック）にしれを使う。"""
    p = _exclude_rules_effective_read_path
    if p and os.path.exists(p):
        return p
    return wb_path
_EXCLUDE_RULES_JSON_ENV_MEMO_UNSET = object()
_exclude_rules_json_env_memo: object = _EXCLUDE_RULES_JSON_ENV_MEMO_UNSET
def _get_exclude_rules_from_json_env() -> list[dict] | None:
    """PM_AI_EXCLUDE_RULES_JSON が実在し形式が有効なら rules のリスト、無効時は None（Excel にフォールバック）。

    同一プロセス内は結果をメモ化する（段階1の保守スキップ判定と適用で二重読込しない）。
    """
    global _exclude_rules_json_env_memo
    if _exclude_rules_json_env_memo is not _EXCLUDE_RULES_JSON_ENV_MEMO_UNSET:
        return _exclude_rules_json_env_memo  # type: ignore[return-value]
    json_env = (os.environ.get(ENV_EXCLUDE_RULES_JSON) or "").strip()
    if not json_env or not os.path.isfile(json_env):
        _exclude_rules_json_env_memo = None
        return None
    try:
        with open(json_env, encoding="utf-8-sig") as f:
            data = json.load(f)
    except Exception as e:
        logging.warning(
            "配台試行ルール: %s の JSON 読込に失敗しました（%s）。ブックのシート読込にフォールバックします。",
            json_env,
            e,
        )
        _exclude_rules_json_env_memo = None
        return None
    rows = data.get("rules") if isinstance(data, dict) and "rules" in data else data
    if not isinstance(rows, list):
        logging.warning(
            "配台試行ルール: %s は list または {\"rules\": [...]} ではありません。フォールバックします。",
            json_env,
        )
        _exclude_rules_json_env_memo = None
        return None
    rules: list[dict] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        proc = str(
            row.get(EXCLUDE_RULE_COL_PROCESS)
            or row.get("process")
            or row.get("Process")
            or ""
        ).strip()
        mach = str(
            row.get(EXCLUDE_RULE_COL_MACHINE)
            or row.get("machine")
            or row.get("Machine")
            or ""
        ).strip()
        if not proc:
            continue
        c_val = row.get(EXCLUDE_RULE_COL_FLAG)
        if c_val is None:
            c_val = row.get("exclude_flag") or row.get("配台不要")
        e_raw = row.get(EXCLUDE_RULE_COL_LOGIC_JSON)
        if e_raw is None:
            e_raw = row.get("logic_json") or row.get("logic")
        parsed = _parse_exclude_rule_json_cell(e_raw)
        rules.append(
            {
                "proc": proc,
                "mach": mach,
                "c_val": c_val,
                "parsed": parsed,
            }
        )
    logging.info(
        "配台試行ルール: JSON '%s' から %s 件（Excel の「%s」read_excel はスキップ）。",
        json_env,
        len(rules),
        EXCLUDE_RULES_SHEET_NAME,
    )
    _exclude_rules_json_env_memo = rules
    return rules
def _reset_exclude_rules_json_env_memo() -> None:
    global _exclude_rules_json_env_memo
    _exclude_rules_json_env_memo = _EXCLUDE_RULES_JSON_ENV_MEMO_UNSET
def _json_safe_cell_for_exclude_rules_export(value) -> object:
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        try:
            return value.isoformat()
        except Exception:
            return str(value)
    if isinstance(value, (int, float, str, bool)):
        return value
    return str(value)
def _exclude_rule_json_row_proc_mach(row: dict) -> tuple[str, str]:
    """JSON ルール行から工程名・機械名（表示用）を取り出す。"""
    proc = str(
        row.get(EXCLUDE_RULE_COL_PROCESS)
        or row.get("process")
        or row.get("Process")
        or ""
    ).strip()
    mach = str(
        row.get(EXCLUDE_RULE_COL_MACHINE)
        or row.get("machine")
        or row.get("Machine")
        or ""
    ).strip()
    return proc, mach
def _merge_exclude_rules_json_with_plan_pairs(
    json_path: str,
    pairs: list[tuple[str, str]],
    log_prefix: str,
) -> int:
    """``PM_AI_EXCLUDE_RULES_JSON`` 正本に、計画タスク上の (工程名, 機械名) で未登録の行を追記する。

    既存行の ``配台不要`` / ロジック列は変更しない。追記のみ。
    成功時に ``_reset_exclude_rules_json_env_memo`` でメモを無効化する。
    """
    path = (json_path or "").strip()
    if not path or not os.path.isfile(path) or not pairs:
        return 0
    try:
        with open(path, encoding="utf-8-sig") as f:
            data = json.load(f)
    except Exception as ex:
        logging.warning(
            "%s: 配台不要ルール JSON の読込に失敗し行同期をスキップ: %s (%s)",
            log_prefix,
            path,
            ex,
        )
        return 0

    if isinstance(data, list):
        top: dict = {"rules": data}
        rules_list: list = data
    elif isinstance(data, dict):
        top = dict(data)
        rules_list = top.get("rules")
        if not isinstance(rules_list, list):
            rules_list = []
            top["rules"] = rules_list
    else:
        logging.warning(
            "%s: 配台不要ルール JSON が list / object ではないため行同期をスキップ: %s",
            log_prefix,
            path,
        )
        return 0

    existing_keys: set[tuple[str, str]] = set()
    for row in rules_list:
        if not isinstance(row, dict):
            continue
        proc, mach = _exclude_rule_json_row_proc_mach(row)
        if not proc:
            continue
        existing_keys.add(
            (
                _normalize_process_name_for_rule_match(proc),
                _normalize_equipment_match_key(mach),
            )
        )

    added = 0
    for p, m in pairs:
        key = (
            _normalize_process_name_for_rule_match(p),
            _normalize_equipment_match_key(m),
        )
        if key in existing_keys:
            continue
        rules_list.append(
            {
                EXCLUDE_RULE_COL_PROCESS: p,
                EXCLUDE_RULE_COL_MACHINE: m,
                EXCLUDE_RULE_COL_FLAG: None,
                EXCLUDE_RULE_COL_LOGIC_JA: None,
                EXCLUDE_RULE_COL_LOGIC_JSON: None,
            }
        )
        existing_keys.add(key)
        added += 1

    if not added:
        return 0

    abs_out = os.path.abspath(path)
    parent = os.path.dirname(abs_out)
    try:
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(abs_out, "w", encoding="utf-8", newline="\n") as f:
            json.dump(top, f, ensure_ascii=False, indent=2)
    except OSError as ex:
        logging.warning(
            "%s: 配台不要ルール JSON の書き込みに失敗（追記 %s 件は未保存）: %s",
            log_prefix,
            added,
            ex,
        )
        return 0

    _reset_exclude_rules_json_env_memo()
    logging.info(
        "%s: 配台不要ルール JSON に工程+機械の組み合わせを %s 件追加しました（%s）。",
        log_prefix,
        added,
        abs_out,
    )
    return int(added)
def _exclude_rules_json_env_supersedes_excel_sheet() -> bool:
    """True のときルール読込は JSON を正とし Excel「設定_配台不要工程」は読まない。

    行同期は ``_merge_exclude_rules_json_with_plan_pairs`` で JSON へ追記する。
    """
    return _get_exclude_rules_from_json_env() is not None
def _resolve_summary_ai_dispatch_workbook_path() -> str:
    """Java AppPaths.summaryAiDispatchXlsxPath と同じ解決（共有 DATA 内のサマリ_AI配台.xlsx）。"""
    from planning_core.core.summary_shared_data_paths import (
        resolve_summary_ai_dispatch_workbook_path,
    )

    return resolve_summary_ai_dispatch_workbook_path()
def _resolve_stage1_exclude_rules_json_work_path() -> str:
    """ローカル output 配下の stage1_exclude_rules.json 絶対パス（共有へは出さない）。"""
    override = (os.environ.get("PM_AI_OUTPUT_DIR") or "").strip()
    if override:
        abs_override = os.path.abspath(override)
        s = abs_override.replace("/", "\\")
        shared = (
            s.startswith("\\\\")
            or s.startswith("//")
            or "\\共有DATA" in s
            or "\\●配台AIシステム\\" in s
            or "\\配台AIシステム\\" in s
            or s.lower().startswith("m:\\湖南工場")
        )
        if not shared:
            return os.path.normpath(
                os.path.join(abs_override, STAGE1_EXCLUDE_RULES_JSON_FILENAME)
            )
    repo = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
    if repo:
        parent = os.path.join(os.path.abspath(repo), "output")
    else:
        parent = os.path.join(os.getcwd(), "output")
    return os.path.normpath(os.path.join(parent, STAGE1_EXCLUDE_RULES_JSON_FILENAME))


def _resolve_stage1_exclude_rules_json_legacy_beside_summary() -> str:
    """旧作業先: サマリ Excel と同フォルダ（共有に残っている場合の移行元）。"""
    summary = _resolve_summary_ai_dispatch_workbook_path()
    parent = os.path.dirname(summary)
    if not parent:
        repo = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
        parent = os.path.join(repo or os.getcwd(), "code")
    return os.path.normpath(os.path.join(parent, STAGE1_EXCLUDE_RULES_JSON_FILENAME))


def _copy_exclude_rules_json_if_missing(target: str, source: str) -> bool:
    if os.path.isfile(target):
        return True
    if not source or not os.path.isfile(source):
        return False
    parent = os.path.dirname(target)
    if parent:
        os.makedirs(parent, exist_ok=True)
    try:
        shutil.copy2(source, target)
    except OSError:
        return False
    return os.path.isfile(target)


def _ensure_stage1_exclude_rules_json_at_work_path() -> str | None:
    """ローカル作業先に JSON が無ければリポジトリ同梱・旧サマリ同階層・旧 cwd/json からコピー。"""
    target = _resolve_stage1_exclude_rules_json_work_path()
    if os.path.isfile(target):
        return target
    bundled = _resolve_default_exclude_rules_json_path_for_env()
    if bundled and _copy_exclude_rules_json_if_missing(target, bundled):
        logging.info(
            "配台不要ルール JSON をリポジトリ同梱から作業先へコピーしました（%s → %s）。",
            bundled,
            target,
        )
        return target
    legacy_summary = _resolve_stage1_exclude_rules_json_legacy_beside_summary()
    if legacy_summary != target and _copy_exclude_rules_json_if_missing(
        target, legacy_summary
    ):
        logging.info(
            "配台不要ルール JSON を旧配置（サマリ同階層）から作業先へコピーしました（%s → %s）。",
            legacy_summary,
            target,
        )
        return target
    legacy = os.path.join(json_data_dir, STAGE1_EXCLUDE_RULES_JSON_FILENAME)
    if legacy != target and _copy_exclude_rules_json_if_missing(target, legacy):
        logging.info(
            "配台不要ルール JSON を旧配置から作業先へコピーしました（%s → %s）。",
            legacy,
            target,
        )
        return target
    return None
def _resolve_repo_root_exclude_rules_json_path() -> str | None:
    """リポジトリの ``code/exclude_rules.json``（配台不要ルールの既定・JavaFX と同じ）があれば絶対パス。"""
    repo = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
    candidates: list[str] = []
    if repo:
        candidates.append(
            os.path.normpath(os.path.join(repo, "code", "exclude_rules.json"))
        )
    cwd = os.path.abspath(os.getcwd())
    candidates.append(os.path.join(cwd, "code", "exclude_rules.json"))
    try:
        _pkg_dir = os.path.dirname(os.path.abspath(__file__))
        _code_dir = os.path.normpath(os.path.join(_pkg_dir, "..", ".."))
        candidates.append(os.path.join(_code_dir, "exclude_rules.json"))
    except Exception:
        pass
    seen: set[str] = set()
    for p in candidates:
        if not p or p in seen:
            continue
        seen.add(p)
        if os.path.isfile(p):
            return p
    return None
def _resolve_default_exclude_rules_json_path_for_env() -> str | None:
    """リポジトリ同梱テンプレート: ``code/exclude_rules.json`` を優先し、無ければ ``code/json/stage1_exclude_rules.json``。"""
    primary = _resolve_repo_root_exclude_rules_json_path()
    if primary:
        return primary
    repo = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
    name = STAGE1_EXCLUDE_RULES_JSON_FILENAME
    candidates: list[str] = []
    if repo:
        candidates.append(os.path.normpath(os.path.join(repo, "code", "json", name)))
    cwd = os.path.abspath(os.getcwd())
    candidates.append(os.path.join(cwd, "code", "json", name))
    try:
        _pkg_dir = os.path.dirname(os.path.abspath(__file__))
        _code_dir = os.path.normpath(os.path.join(_pkg_dir, "..", ".."))
        candidates.append(os.path.join(_code_dir, "json", name))
    except Exception:
        pass
    seen: set[str] = set()
    for cand in candidates:
        if not cand or cand in seen:
            continue
        seen.add(cand)
        if os.path.isfile(cand):
            return cand
    return None
def _ensure_stage1_exclude_rules_json_env_from_repo_default() -> None:
    """``PM_AI_EXCLUDE_RULES_JSON`` が未設定・共有上・実在しないとき、ローカル作業 JSON を正本として載せる。"""
    cur = (os.environ.get(ENV_EXCLUDE_RULES_JSON) or "").strip()
    if cur and os.path.isfile(cur):
        s = cur.replace("/", "\\")
        shared = (
            s.startswith("\\\\")
            or s.startswith("//")
            or "\\共有DATA" in s
            or "\\●配台AIシステム\\" in s
            or "\\配台AIシステム\\" in s
            or s.lower().startswith("m:\\湖南工場")
        )
        if not shared:
            _reset_exclude_rules_json_env_memo()
            return
    work = _ensure_stage1_exclude_rules_json_at_work_path()
    if not work:
        return
    os.environ[ENV_EXCLUDE_RULES_JSON] = work
    _reset_exclude_rules_json_env_memo()
    logging.info(
        "段階1: 配台不要ルールの正本として JSON を使用します（%s）。Excel「%s」は参照しません。",
        work,
        EXCLUDE_RULES_SHEET_NAME,
    )
def _load_exclude_rules_from_workbook(wb_path: str) -> list[dict]:
    """シートからルール行を読み」評価用リストを返す。"""
    if not wb_path:
        return []
    global _exclude_rules_rules_snapshot, _exclude_rules_snapshot_wb
    ap_arg = os.path.normcase(os.path.abspath(wb_path))
    if (
        _exclude_rules_rules_snapshot is not None
        and _exclude_rules_snapshot_wb == ap_arg
    ):
        snap = list(_exclude_rules_rules_snapshot)
        _exclude_rules_rules_snapshot = None
        _exclude_rules_snapshot_wb = None
        return snap
    json_rules = _get_exclude_rules_from_json_env()
    if json_rules is not None:
        return json_rules
    path = _resolve_exclude_rules_workbook_path_for_read(wb_path)
    if not os.path.exists(path):
        return []
    if _workbook_should_skip_openpyxl_io(path):
        logging.warning(
            "配台試行ルール: ブックに「%s」があるため、pandas(openpyxl) での「%s」読込をスキップしました（ルールは未適用）。",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
            EXCLUDE_RULES_SHEET_NAME,
        )
        return []
    try:
        df = pd.read_excel(path, sheet_name=EXCLUDE_RULES_SHEET_NAME)
    except Exception:
        return []
    df.columns = df.columns.str.strip()
    need = [EXCLUDE_RULE_COL_PROCESS, EXCLUDE_RULE_COL_MACHINE]
    for c in need:
        if c not in df.columns:
            return []
    rules = []
    for _, row in df.iterrows():
        proc = str(row.get(EXCLUDE_RULE_COL_PROCESS, "") or "").strip()
        if not proc:
            continue
        mach = str(row.get(EXCLUDE_RULE_COL_MACHINE, "") or "").strip()
        c_val = row.get(EXCLUDE_RULE_COL_FLAG)
        e_raw = row.get(EXCLUDE_RULE_COL_LOGIC_JSON)
        parsed = _parse_exclude_rule_json_cell(e_raw)
        rules.append(
            {
                "proc": proc,
                "mach": mach,
                "c_val": c_val,
                "parsed": parsed,
            }
        )
    return rules
def apply_exclude_rules_config_to_plan_df(
    df: pd.DataFrame, wb_path: str, log_prefix: str
) -> pd.DataFrame:
    """設定シートに基づき「配台不要」を設定（C=yes または E の JSON が真）。

    工程名が「分割」の行は、同一依頼NO内に非空の同一「機械名」が複数行ある場合に限り
    （``_apply_auto_exclude_bunkatsu_duplicate_machine`` と同じ重複条件）、C/E を適用する。
    EC と分割で機械が異なる依頼では、設定行が残っていても当該分割行は配台対象のままとする。

    運用上は **段階1**（``run_stage1_extract``）から呼ぶ。``PM_AI_EXCLUDE_RULES_JSON`` が有効なときは
    ルール元は JSON のみ（Excel シートは読まない）。段階2の ``load_planning_tasks_df`` では
    ``_apply_planning_sheet_post_load_mutations(..., apply_exclude_rules_from_config=False,
    compile_exclude_rules_d_to_e_with_ai=False)`` とし、本関数でシートの C/E を計画 DataFrame に
    再適用しない（設定シートの D→E AI も段階2では行わない）。
    """
    if df is None or df.empty:
        return df
    if TASK_COL_MACHINE not in df.columns or PLAN_COL_EXCLUDE_FROM_ASSIGNMENT not in df.columns:
        return df
    rules = _load_exclude_rules_from_workbook(wb_path)
    if not rules:
        return df
    df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT].astype(object)
    by_tid_idx: dict[str, list] = defaultdict(list)
    if TASK_COL_TASK_ID in df.columns:
        for j in df.index:
            tid_j = _normalize_task_id_for_dup_grouping(df.at[j, TASK_COL_TASK_ID])
            if tid_j:
                by_tid_idx[tid_j].append(j)
    n = 0
    for i in df.index:
        try:
            row = df.loc[i]
        except Exception:
            continue
        tp = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        tm = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
        if not tp:
            continue
        tid_norm = _normalize_task_id_for_dup_grouping(row.get(TASK_COL_TASK_ID))
        is_bunkatsu = _process_name_is_bunkatsu_for_auto_exclude(tp)
        dup_ge2_for_tid = False
        if tid_norm:
            dup_ge2_for_tid, _mc = _same_tid_nonempty_machine_dup_ge2(
                df, by_tid_idx.get(tid_norm, [])
            )
        bunkatsu_block_cfg = is_bunkatsu and bool(tid_norm) and not dup_ge2_for_tid
        for ru in rules:
            if not _task_row_matches_exclude_rule_target(tp, tm, ru["proc"], ru["mach"]):
                continue
            if bunkatsu_block_cfg:
                continue
            if _exclude_rule_c_column_is_yes(ru["c_val"]):
                df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
                n += 1
                break
            if ru.get("parsed") and evaluate_exclude_rule_json_for_row(ru["parsed"], row):
                df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
                n += 1
                break
    if n:
        logging.info("%s: 設定「%s」により配台不要=yes を %s 行に設定しました。", log_prefix, EXCLUDE_RULES_SHEET_NAME, n)
    return df
def _sort_stage1_plan_df_by_dispatch_trial_order_asc(plan_df: "pd.DataFrame") -> "pd.DataFrame":
    """
    段階1出力直後: 配台試行順番の昇順に行を並き替ごた DataFrame を返す。
    正の整数でないセルは最後（同一帯内は元の行順）。
    """
    col = RESULT_TASK_COL_DISPATCH_TRIAL_ORDER
    if plan_df is None or getattr(plan_df, "empty", True) or col not in plan_df.columns:
        return plan_df
    dto_positions = [i for i, c in enumerate(plan_df.columns) if c == col]
    if not dto_positions:
        return plan_df
    loc = dto_positions[0]
    n = len(plan_df)
    if n <= 1:
        return plan_df
    keys = []
    from planning_core.core.plan_input import dispatch_trial_order_positive_finite

    for i in range(n):
        dto = dispatch_trial_order_positive_finite(plan_df.iat[i, loc])
        if dto is not None:
            keys.append((0, float(dto), i))
        else:
            keys.append((1, 10**9, i))
    order = sorted(range(n), key=lambda j: keys[j])
    if order == list(range(n)):
        return plan_df
    return plan_df.iloc[order].reset_index(drop=True)
def _raw_fabric_width_table_search_paths() -> list[str]:
    """原反幅テーブル CSV の探索順（先に見つかったパスを採用）。"""
    return _material_mm_table_search_paths(
        RAW_FABRIC_WIDTH_TABLE_DEFAULT_FILENAME, RAW_FABRIC_WIDTH_TABLE_PATH_ENV
    )
def _summary_ai_dispatch_workbook_sibling_path(filename: str) -> str:
    from planning_core.core.summary_shared_data_paths import summary_shared_data_sibling_path

    return summary_shared_data_sibling_path(filename)
def _planning_code_dir_candidates() -> list[str]:
    """Java {@code PM_AI_CODE_DIR} / python 隣接 code / repo/code の候補（順序付き・重複除去前）。"""
    out: list[str] = []
    code_dir = (os.environ.get("PM_AI_CODE_DIR") or "").strip()
    if code_dir:
        out.append(code_dir)
    py_dir = (os.environ.get("PM_AI_CODE_PYTHON_DIR") or "").strip()
    if py_dir:
        parent = os.path.abspath(os.path.join(py_dir, os.pardir))
        if parent:
            out.append(parent)
    repo = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
    if repo:
        out.append(os.path.join(repo, "code"))
    try:
        pkg_code = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if os.path.basename(pkg_code).lower() == "code":
            out.append(pkg_code)
    except OSError:
        pass
    return out
def _resolve_planning_code_dir() -> str:
    for candidate in _planning_code_dir_candidates():
        if candidate and os.path.isdir(candidate):
            return os.path.abspath(candidate)
    return ""
def _is_under_system_temp(path: str) -> bool:
    """OS 一時ディレクトリ配下（Excel 展開の /tmp 等）か。材料テーブル正本の取りこぼし防止。"""
    if not path:
        return False
    try:
        abs_p = os.path.abspath(path)
        parts = abs_p.replace("\\", "/").split("/")
        if "tmp" in parts or "temp" in parts:
            return True
        import tempfile

        temp_root = os.path.abspath(tempfile.gettempdir())
        common = os.path.commonpath([abs_p, temp_root])
        return common == temp_root
    except (OSError, ValueError):
        return False
def _canonical_material_table_path(default_filename: str) -> str:
    sibling = _summary_ai_dispatch_workbook_sibling_path(default_filename)
    if sibling:
        return sibling
    code_dir = _resolve_planning_code_dir()
    if code_dir:
        return os.path.join(code_dir, default_filename)
    return ""
def _pick_material_table_path_for_read(
    search_paths: list[str], default_filename: str
) -> str:
    """読込: 正本 code/ を OS 一時フォルダ内の同名 CSV より優先する。"""
    canonical = _canonical_material_table_path(default_filename)
    if canonical and os.path.isfile(canonical):
        return canonical
    for p in search_paths:
        if not os.path.isfile(p):
            continue
        if canonical and _is_under_system_temp(p):
            continue
        return p
    return ""
def _material_mm_table_search_paths(default_filename: str, path_env_var: str) -> list[str]:
    """mm 系材料テーブル CSV の探索順。作業先（サマリ Excel 同フォルダ）を code/ より先にする。"""
    paths: list[str] = []
    env = (os.environ.get(path_env_var) or "").strip()
    if env:
        paths.append(env)
    sibling = _summary_ai_dispatch_workbook_sibling_path(default_filename)
    if sibling:
        paths.append(sibling)
    for code_dir in _planning_code_dir_candidates():
        paths.append(os.path.join(code_dir, default_filename))
    wb = (_excel_plan_input_wb() or "").strip()
    if wb:
        paths.append(
            os.path.join(os.path.dirname(os.path.abspath(wb)), default_filename)
        )
    paths.append(os.path.join(os.getcwd(), default_filename))
    paths.append(os.path.join(os.getcwd(), "code", default_filename))
    out: list[str] = []
    seen: set[str] = set()
    for p in paths:
        key = os.path.normcase(os.path.abspath(p))
        if key not in seen:
            seen.add(key)
            out.append(p)
    return out
def _normalize_mm_table_lookup_key(val) -> str:
    """
    製品長・製品幅（製品名, 製品幅.txt）・製品厚み（製品名,製品厚み.txt）・原反幅（使用原反, 加工幅.txt）など mm 系 CSV の照会キーを正規化する。
    NFKC で全角英数字等を寄せたうえで、半角・全角などあらゆる空白類（isspace）を除去する
    （ロール単位長さテーブルの `_normalize_roll_unit_length_table_key` と同趣旨）。
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = unicodedata.normalize("NFKC", str(val).strip())
    return "".join(ch for ch in s if not ch.isspace())
def _resolve_code_lookup_table_path_for_write(
    search_paths: list[str], default_filename: str
) -> str:
    """材料テーブル CSV の追記先。正本 code/ が解決できれば常にそこ（/tmp 側の同名ファイルは無視）。"""
    canonical = _canonical_material_table_path(default_filename)
    if canonical:
        return canonical
    for p in search_paths:
        if os.path.isfile(p) and not _is_under_system_temp(p):
            return p
    repo = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
    if repo:
        return os.path.join(repo, "code", default_filename)
    return os.path.join(os.getcwd(), "code", default_filename)
def _lookup_display_for_normalized_key(
    row: "pd.Series", cols: tuple[str, ...], normalized_key: str
) -> str:
    for col in cols:
        if col not in row.index:
            continue
        raw = row.get(col)
        if _normalize_mm_table_lookup_key(raw) == normalized_key:
            display = str(raw or "").strip()
            if display:
                return display
    return normalized_key
def _append_code_dispatch_lookup_table_row_if_missing(
    raw_key_display,
    normalized_key: str,
    *,
    table_path: str,
    header_line: str,
    known_keys: set[str],
    appended: set[str],
    log_table_label: str,
) -> bool:
    """段階1: 材料テーブルに未登録キーを空欄値で追記する。"""
    skip_reason = ""
    if not normalized_key:
        skip_reason = "empty_key"
    elif normalized_key in known_keys:
        skip_reason = "in_known_keys"
    elif normalized_key in appended:
        skip_reason = "in_appended"
    elif not (table_path or "").strip():
        skip_reason = "empty_table_path"
    if not normalized_key or normalized_key in known_keys or normalized_key in appended:
        if skip_reason:
            logging.info(
                "%sテーブル追記スキップ（%s）: %r path=%s",
                log_table_label,
                skip_reason,
                normalized_key,
                table_path,
            )
        return False
    display = str(raw_key_display or "").strip()
    if not display:
        display = normalized_key
    path = (table_path or "").strip()
    if not path:
        return False
    line = f"{display},\n"
    try:
        parent = os.path.dirname(os.path.abspath(path))
        if parent:
            os.makedirs(parent, exist_ok=True)
        if not os.path.isfile(path):
            with open(path, "w", encoding="utf-8", newline="") as f:
                f.write(f"{header_line}\n{line}")
        else:
            with open(path, "a", encoding="utf-8", newline="") as f:
                f.write(line)
        known_keys.add(normalized_key)
        appended.add(normalized_key)
        logging.info(
            "%sテーブルに未登録キーを追記しました（値は空欄）: %r → %s",
            log_table_label,
            display,
            path,
        )
        return True
    except OSError as ex:
        logging.warning(
            "%sテーブルへの追記に失敗: キー=%r パス=%s (%s)",
            log_table_label,
            display,
            path,
            ex,
        )
        return False
def _load_int_mm_lookup_table(
    search_paths: list[str],
    *,
    default_filename: str,
    path_env: str,
    table_label_ja: str,
    key_header: str,
    value_header: str,
) -> tuple[dict[str, int], set[str], str]:
    """mm 整数の材料テーブルを読み込む。値が空欄の行は known_keys のみ（dict には載せない）。"""
    path_found = _pick_material_table_path_for_read(search_paths, default_filename)
    if not path_found:
        hint = " / ".join(search_paths[:4])
        raise PlanningValidationError(
            f"{table_label_ja}テーブルが見つかりません。{default_filename} を配置するか、"
            f"環境変数 {path_env} で CSV のフルパスを指定してください。探索: {hint}"
        )
    out: dict[str, int] = {}
    known_keys: set[str] = set()
    with open(path_found, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        raise PlanningValidationError(f"{table_label_ja}テーブルが空です: {path_found}")
    hdr = [_normalize_mm_table_lookup_key(x) for x in rows[0]]
    try:
        i_key = hdr.index(_normalize_mm_table_lookup_key(key_header))
    except ValueError:
        i_key = 0
    try:
        i_w = hdr.index(_normalize_mm_table_lookup_key(value_header))
    except ValueError:
        i_w = 1 if len(hdr) > 1 else 0
    for parts in rows[1:]:
        if not parts or all(not str(x).strip() for x in parts):
            continue
        while len(parts) <= max(i_key, i_w):
            parts.append("")
        raw_k = parts[i_key]
        raw_w = parts[i_w]
        key = _normalize_mm_table_lookup_key(raw_k)
        if not key:
            continue
        known_keys.add(key)
        try:
            w = _parse_int_mm_width_table_cell(raw_w)
        except ValueError:
            continue
        if key in out and out[key] != w:
            raise PlanningValidationError(
                f"{table_label_ja}テーブルで同一キーに矛盾する値があります: {key!r} → {out[key]} と {w} ({path_found})"
            )
        out[key] = w
    logging.info("%sテーブルを読み込みました: %s (%s 件)", table_label_ja, path_found, len(out))
    return out, known_keys, path_found
def _parse_int_mm_width_table_cell(val) -> int:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        raise ValueError("empty")
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if isinstance(val, float) and (not math.isfinite(val)):
            raise ValueError("non-finite")
        n = int(round(float(val)))
        if n <= 0:
            raise ValueError("non-positive")
        return n
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none"):
        raise ValueError("empty")
    s = unicodedata.normalize("NFKC", s)
    s = s.replace(",", "")
    n = int(round(float(s)))
    if n <= 0:
        raise ValueError("non-positive")
    return n
def _infer_width_mm_from_last_dim_pair_left(text: str) -> int | None:
    """
    最後の「左数 x 右数」（2〜6 桁）ペアの左側を幅(mm)候補とする。
    原反幅・製品幅テーブル未登録時のフォールバック。寸法区切りは infer_unit_m と同様に正規化する。
    """
    s = _normalize_product_dim_separators_for_roll_inference(str(text or ""))
    dim_pairs = re.findall(r"(\d{2,6})\s*[xX]\s*(\d{2,6})", s)
    if not dim_pairs:
        return None
    try:
        left = int(dim_pairs[-1][0])
        return left if left > 0 else None
    except ValueError:
        return None
def _infer_length_mm_from_last_dim_pair_right(text: str) -> int | None:
    """
    最後の「左数 x 右数」（2〜6 桁）ペアの右側を長さ(mm)候補とする。
    製品長テーブル未登録時のフォールバック。寸法区切りは infer_unit_m と同様に正規化する。
    """
    s = _normalize_product_dim_separators_for_roll_inference(str(text or ""))
    dim_pairs = re.findall(r"(\d{2,6})\s*[xX]\s*(\d{2,6})", s)
    if not dim_pairs:
        return None
    try:
        right = int(dim_pairs[-1][1])
        return right if right > 0 else None
    except ValueError:
        return None
def _load_raw_fabric_width_mm_table() -> tuple[dict[str, int], set[str], str]:
    """原反幅テーブル（使用原反→原反幅）を読み込む。"""
    return _load_int_mm_lookup_table(
        _raw_fabric_width_table_search_paths(),
        default_filename=RAW_FABRIC_WIDTH_TABLE_DEFAULT_FILENAME,
        path_env=RAW_FABRIC_WIDTH_TABLE_PATH_ENV,
        table_label_ja="原反幅",
        key_header="使用原反",
        value_header="原反幅",
    )
def _raw_fabric_width_lookup_source_strings(row: "pd.Series") -> list[str]:
    """照会用文字列（正規化済み・重複除き）。使用原反 → 製品名。"""
    ordered: list[str] = []
    seen: set[str] = set()
    for col in (TASK_COL_USED_RAW, TASK_COL_PRODUCT):
        if col not in row.index:
            continue
        v = row.get(col)
        k = _normalize_mm_table_lookup_key(v)
        if not k or k in seen:
            continue
        seen.add(k)
        ordered.append(k)
    return ordered
def _resolve_raw_fabric_width_mm_for_stage1_row(
    row: "pd.Series",
    table: dict[str, int],
    *,
    table_path: str = "",
    known_keys: set[str] | None = None,
    appended: set[str] | None = None,
) -> int | None:
    """
    テーブル照会を優先し、未登録なら寸法パターンで原反幅(mm)を決定。
    決められなければ None（段階1では空欄出力し材料テーブルへ追記）。
    """
    keys = _raw_fabric_width_lookup_source_strings(row)
    for k in keys:
        w = table.get(k)
        if w is not None and w > 0:
            return int(w)
    for k in keys:
        inferred = _infer_width_mm_from_last_dim_pair_left(k)
        if inferred is not None and inferred > 0:
            return int(inferred)
    tid = planning_task_id_str_from_scalar(row.get(TASK_COL_TASK_ID))
    keys_set = known_keys if known_keys is not None else set(table.keys())
    pending = appended if appended is not None else set()
    if keys:
        nk = keys[0]
        display = _lookup_display_for_normalized_key(
            row, (TASK_COL_USED_RAW, TASK_COL_PRODUCT), nk
        )
        _append_code_dispatch_lookup_table_row_if_missing(
            display,
            nk,
            table_path=_resolve_code_lookup_table_path_for_write(
                _raw_fabric_width_table_search_paths(),
                RAW_FABRIC_WIDTH_TABLE_DEFAULT_FILENAME,
            ),
            header_line="使用原反,原反幅",
            known_keys=keys_set,
            appended=pending,
            log_table_label="原反幅",
        )
    logging.warning(
        "原反幅未登録（テーブル未登録かつ寸法パターンからも解釈不可）。材料テーブルへ追記し原反幅は空欄で出力。依頼NO=%s 照会キー=%r",
        tid,
        keys,
    )
    return None
def _product_width_table_search_paths() -> list[str]:
    """製品幅テーブル CSV の探索順（先に見つかったパスを採用）。"""
    return _material_mm_table_search_paths(
        PRODUCT_WIDTH_TABLE_DEFAULT_FILENAME, PRODUCT_WIDTH_TABLE_PATH_ENV
    )
def _load_product_width_mm_table() -> tuple[dict[str, int], set[str], str]:
    """製品幅テーブル（製品名→製品幅）を読み込む。"""
    return _load_int_mm_lookup_table(
        _product_width_table_search_paths(),
        default_filename=PRODUCT_WIDTH_TABLE_DEFAULT_FILENAME,
        path_env=PRODUCT_WIDTH_TABLE_PATH_ENV,
        table_label_ja="製品幅",
        key_header="製品名",
        value_header="製品幅",
    )
def _product_length_table_search_paths() -> list[str]:
    """製品長テーブル CSV の探索順（先に見つかったパスを採用）。"""
    return _material_mm_table_search_paths(
        PRODUCT_LENGTH_TABLE_DEFAULT_FILENAME, PRODUCT_LENGTH_TABLE_PATH_ENV
    )
def _load_product_length_mm_table() -> tuple[dict[str, int], set[str], str]:
    """製品長テーブル（製品名→製品長）を読み込む。"""
    return _load_int_mm_lookup_table(
        _product_length_table_search_paths(),
        default_filename=PRODUCT_LENGTH_TABLE_DEFAULT_FILENAME,
        path_env=PRODUCT_LENGTH_TABLE_PATH_ENV,
        table_label_ja="製品長",
        key_header="製品名",
        value_header="製品長",
    )
def _product_length_lookup_source_strings(row: "pd.Series") -> list[str]:
    """照会用文字列（正規化済み・重複除き）。製品名 → 使用原反。"""
    ordered: list[str] = []
    seen: set[str] = set()
    for col in (TASK_COL_PRODUCT, TASK_COL_USED_RAW):
        if col not in row.index:
            continue
        v = row.get(col)
        k = _normalize_mm_table_lookup_key(v)
        if not k or k in seen:
            continue
        seen.add(k)
        ordered.append(k)
    return ordered
def _resolve_product_length_mm_for_stage1_row(
    row: "pd.Series",
    table: dict[str, int],
    *,
    table_path: str = "",
    known_keys: set[str] | None = None,
    appended: set[str] | None = None,
) -> int | None:
    """
    テーブル照会を優先し、未登録なら寸法パターンで製品長(mm)を決定。
    決められなければ None（段階1では空欄出力し材料テーブルへ追記）。
    """
    keys = _product_length_lookup_source_strings(row)
    for k in keys:
        w = table.get(k)
        if w is not None and w > 0:
            return int(w)
    for k in keys:
        inferred = _infer_length_mm_from_last_dim_pair_right(k)
        if inferred is not None and inferred > 0:
            return int(inferred)
    tid = planning_task_id_str_from_scalar(row.get(TASK_COL_TASK_ID))
    keys_set = known_keys if known_keys is not None else set(table.keys())
    pending = appended if appended is not None else set()
    if keys:
        nk = keys[0]
        display = _lookup_display_for_normalized_key(
            row, (TASK_COL_PRODUCT, TASK_COL_USED_RAW), nk
        )
        _append_code_dispatch_lookup_table_row_if_missing(
            display,
            nk,
            table_path=_resolve_code_lookup_table_path_for_write(
                _product_length_table_search_paths(),
                PRODUCT_LENGTH_TABLE_DEFAULT_FILENAME,
            ),
            header_line="製品名,製品長",
            known_keys=keys_set,
            appended=pending,
            log_table_label="製品長",
        )
    logging.warning(
        "製品長未登録（テーブル未登録かつ寸法パターンからも解釈不可）。材料テーブルへ追記し製品長は空欄で出力。依頼NO=%s 照会キー=%r",
        tid,
        keys,
    )
    return None
def _product_width_lookup_source_strings(row: "pd.Series") -> list[str]:
    """照会用文字列（正規化済み・重複除き）。製品名 → 使用原反。"""
    ordered: list[str] = []
    seen: set[str] = set()
    for col in (TASK_COL_PRODUCT, TASK_COL_USED_RAW):
        if col not in row.index:
            continue
        v = row.get(col)
        k = _normalize_mm_table_lookup_key(v)
        if not k or k in seen:
            continue
        seen.add(k)
        ordered.append(k)
    return ordered
def _resolve_product_width_mm_for_stage1_row(
    row: "pd.Series",
    table: dict[str, int],
    *,
    table_path: str = "",
    known_keys: set[str] | None = None,
    appended: set[str] | None = None,
) -> int | None:
    """
    テーブル照会を優先し、未登録なら寸法パターンで製品幅(mm)を決定。
    決められなければ None（段階1では空欄出力し材料テーブルへ追記）。
    """
    keys = _product_width_lookup_source_strings(row)
    for k in keys:
        w = table.get(k)
        if w is not None and w > 0:
            return int(w)
    for k in keys:
        inferred = _infer_width_mm_from_last_dim_pair_left(k)
        if inferred is not None and inferred > 0:
            return int(inferred)
    tid = planning_task_id_str_from_scalar(row.get(TASK_COL_TASK_ID))
    keys_set = known_keys if known_keys is not None else set(table.keys())
    pending = appended if appended is not None else set()
    if keys:
        nk = keys[0]
        display = _lookup_display_for_normalized_key(
            row, (TASK_COL_PRODUCT, TASK_COL_USED_RAW), nk
        )
        _append_code_dispatch_lookup_table_row_if_missing(
            display,
            nk,
            table_path=_resolve_code_lookup_table_path_for_write(
                _product_width_table_search_paths(),
                PRODUCT_WIDTH_TABLE_DEFAULT_FILENAME,
            ),
            header_line="製品名,製品幅",
            known_keys=keys_set,
            appended=pending,
            log_table_label="製品幅",
        )
    logging.warning(
        "製品幅未登録（テーブル未登録かつ寸法パターンからも解釈不可）。材料テーブルへ追記し製品幅は空欄で出力。依頼NO=%s 照会キー=%r",
        tid,
        keys,
    )
    return None
def _product_thickness_table_search_paths() -> list[str]:
    """製品厚みテーブル CSV の探索順（先に見つかったパスを採用）。"""
    paths = _material_mm_table_search_paths(
        PRODUCT_THICKNESS_TABLE_DEFAULT_FILENAME, PRODUCT_THICKNESS_TABLE_PATH_ENV
    )
    # リポジトリ同梱（細川/GoogleAIStudio/配下）を直接参照したいケース向け
    try:
        _ga_dir = os.path.dirname(
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        )
        extra = os.path.join(_ga_dir, PRODUCT_THICKNESS_TABLE_DEFAULT_FILENAME)
        key = os.path.normcase(os.path.abspath(extra))
        if key not in {os.path.normcase(os.path.abspath(p)) for p in paths}:
            paths.append(extra)
    except Exception:
        pass
    return paths
def _parse_float_mm_thickness_cell(val) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        raise ValueError("empty")
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if isinstance(val, float) and (not math.isfinite(val)):
            raise ValueError("non-finite")
        x = float(val)
        if x <= 0:
            raise ValueError("non-positive")
        return x
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none"):
        raise ValueError("empty")
    s = unicodedata.normalize("NFKC", s).replace(",", "")
    x = float(s)
    if not math.isfinite(x) or x <= 0:
        raise ValueError("non-positive")
    return float(x)
def _resolve_product_thickness_table_path_for_write() -> str:
    """製品厚みテーブルの追記先パス（既存ファイルがあればそれ、無ければ code/ 配下を新規作成）。"""
    return _resolve_code_lookup_table_path_for_write(
        _product_thickness_table_search_paths(),
        PRODUCT_THICKNESS_TABLE_DEFAULT_FILENAME,
    )
def _load_product_thickness_mm_table() -> tuple[dict[str, float], set[str], str]:
    """
    製品厚みテーブル（製品名→製品厚み）を読み込む。ファイル必須。同一キーで数値が食い違うときは例外。
    製品名キーは _normalize_mm_table_lookup_key で正規化（NFKC・空白除去）して dict に格納する。
    値が空欄の行は dict には載せず known_keys のみに含める（段階1追記分など）。
    """
    path_found = _pick_material_table_path_for_read(
        _product_thickness_table_search_paths(),
        PRODUCT_THICKNESS_TABLE_DEFAULT_FILENAME,
    )
    if not path_found:
        hint = " / ".join(_product_thickness_table_search_paths()[:4])
        raise PlanningValidationError(
            f"製品厚みテーブルが見つかりません。{PRODUCT_THICKNESS_TABLE_DEFAULT_FILENAME} を配置するか、"
            f"環境変数 {PRODUCT_THICKNESS_TABLE_PATH_ENV} で CSV のフルパスを指定してください。探索: {hint}"
        )
    out: dict[str, float] = {}
    known_keys: set[str] = set()
    with open(path_found, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        raise PlanningValidationError(f"製品厚みテーブルが空です: {path_found}")
    hdr = [_normalize_mm_table_lookup_key(x) for x in rows[0]]
    try:
        i_key = hdr.index(_normalize_mm_table_lookup_key("製品名"))
    except ValueError:
        i_key = 0
    try:
        i_w = hdr.index(_normalize_mm_table_lookup_key("製品厚み"))
    except ValueError:
        i_w = 1 if len(hdr) > 1 else 0
    for parts in rows[1:]:
        if not parts or all(not str(x).strip() for x in parts):
            continue
        while len(parts) <= max(i_key, i_w):
            parts.append("")
        raw_k = parts[i_key]
        raw_w = parts[i_w]
        key = _normalize_mm_table_lookup_key(raw_k)
        if not key:
            continue
        known_keys.add(key)
        try:
            w = _parse_float_mm_thickness_cell(raw_w)
        except ValueError:
            continue
        prev = out.get(key)
        if prev is not None and abs(float(prev) - float(w)) > 1e-9:
            raise PlanningValidationError(
                f"製品厚みテーブルで同一キーに矛盾する値があります: {key!r} → {prev} と {w} ({path_found})"
            )
        out[key] = float(w)
    logging.info("製品厚みテーブルを読み込みました: %s (%s 件)", path_found, len(out))
    return out, known_keys, path_found
def _append_product_thickness_table_row_if_missing(
    raw_product_name,
    normalized_key: str,
    *,
    table_path: str,
    known_keys: set[str],
    appended: set[str],
) -> bool:
    """段階1: 製品厚みテーブルに未登録キーを空欄値で追記する。"""
    write_path = _resolve_product_thickness_table_path_for_write()
    return _append_code_dispatch_lookup_table_row_if_missing(
        raw_product_name,
        normalized_key,
        table_path=write_path,
        header_line="製品名,製品厚み",
        known_keys=known_keys,
        appended=appended,
        log_table_label="製品厚み",
    )
def _infer_product_thickness_mm_from_product_name_prefix(product_name) -> float | None:
    """
    製品名の先頭5文字の末尾3桁を厚みコードとして code/10 を返す。
    例: 0R040 → 040 → 4.0, 30100 → 100 → 10.0
    英字開始（例: FEL...）はパターン化不可とし None。
    """
    if product_name is None or (isinstance(product_name, float) and pd.isna(product_name)):
        return None
    s = unicodedata.normalize("NFKC", str(product_name).strip())
    if not s:
        return None
    if s[0].isalpha():
        return None
    if len(s) < 5:
        return None
    head5 = s[:5]
    code3 = head5[-3:]
    if not re.match(r"^\d{3}$", code3):
        return None
    v = int(code3)
    if v <= 0:
        return None
    return float(v) / 10.0
def _resolve_product_thickness_mm_for_stage1_row(
    row: "pd.Series",
    table: dict[str, float],
    *,
    table_path: str = "",
    known_keys: set[str] | None = None,
    appended: set[str] | None = None,
) -> float | None:
    """
    英字開始の製品名はテーブル必須。それ以外は先頭5文字パターンを優先し、失敗時はテーブル。
    テーブルにも無く先頭5文字からも解析できない場合は None（段階1では製品厚みを空欄で出力し、材料テーブルへ追記）。
    """
    tid = planning_task_id_str_from_scalar(row.get(TASK_COL_TASK_ID))
    pn_raw = row.get(TASK_COL_PRODUCT)
    pn = _normalize_mm_table_lookup_key(pn_raw)
    keys = known_keys if known_keys is not None else set(table.keys())
    pending = appended if appended is not None else set()

    def _ensure_table_row() -> None:
        if pn:
            _append_product_thickness_table_row_if_missing(
                pn_raw,
                pn,
                table_path=table_path,
                known_keys=keys,
                appended=pending,
            )

    if pn and pn[0].isalpha():
        if pn in table:
            return float(table[pn])
        _ensure_table_row()
        logging.warning(
            "製品厚み未登録（英字開始・テーブル未登録）。材料テーブルへ追記し製品厚みは空欄で出力。依頼NO=%s 製品名=%r",
            tid,
            pn,
        )
        return None
    inferred = _infer_product_thickness_mm_from_product_name_prefix(pn)
    if inferred is not None and inferred > 0:
        return float(inferred)
    if pn and pn in table:
        return float(table[pn])
    _ensure_table_row()
    logging.warning(
        "製品厚み未登録（先頭5文字から解析不可・テーブル未登録）。材料テーブルへ追記し製品厚みは空欄で出力。依頼NO=%s 製品名=%r",
        tid,
        pn,
    )
    return None
def _apply_stage1_in_progress_dispatch_plan_exclude_marker(df, log_prefix: str = "段階1") -> int:
    """
    ``PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH`` が有効なとき、実加工数が正の行の「配台不要」に
    ``PLAN_COL_STAGE2_DISPATCH_PLAN_EXCLUDE_MARKER``（配台計画除外）を付与する（既に含む場合は据え置き）。
    """
    if not _stage2_truthy_env("PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH"):
        return 0
    if df is None or df.empty:
        return 0
    if (
        TASK_COL_ACTUAL_DONE not in df.columns
        or PLAN_COL_EXCLUDE_FROM_ASSIGNMENT not in df.columns
    ):
        return 0
    marker = PLAN_COL_STAGE2_DISPATCH_PLAN_EXCLUDE_MARKER
    act = pd.to_numeric(df[TASK_COL_ACTUAL_DONE], errors="coerce").fillna(0.0)
    n = 0
    for idx in df.index:
        try:
            av = float(act.loc[idx])
        except (TypeError, ValueError):
            av = 0.0
        if av <= 1e-12:
            continue
        cell = df.at[idx, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT]
        if cell is None or pd.isna(cell):
            cur = ""
        else:
            cur = str(cell).strip()
        if marker in cur:
            continue
        df.at[idx, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = (
            f"{cur} {marker}" if cur else marker
        ).strip()
        n += 1
    if n:
        logging.info(
            "%s: 加工途中を配台しない — 「%s」を実加工数>0 の %s 行へ付与しました。",
            log_prefix,
            marker,
            n,
        )
    return n

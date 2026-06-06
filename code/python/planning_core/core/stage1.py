# -*- coding: utf-8 -*-
# planning_core.core.stage1 — body only (loaded via _core exec chain)
def _stage1_roll_length_for_planning_row(row) -> float:
    """
    段階1: 加工計画由来の1行から ロール単位長さ(m)を計算（``run_stage1_extract`` の merge 前と同一式）。

    製品名テーブル・製品名寸法から得た値を **100m 切上げせず**そのまま返す（使用原反列・原反ロール長テーブルは参照しない）。
    """
    _pn_stage1 = row.get(TASK_COL_PRODUCT, None)
    qty, _done_m, _qtceiled, _from_unp = _plan_row_dispatch_qty_metrics(row)
    _qty_total_s1 = max(0.0, parse_float_safe(row.get(TASK_COL_QTY), 0.0))
    _roll_len = infer_unit_m_from_product_name(
        _pn_stage1,
        _qty_total_s1 if _qty_total_s1 > 0 else qty,
    )
    try:
        _roll_len = float(_roll_len)
    except (TypeError, ValueError):
        _roll_len = _qty_total_s1 if _qty_total_s1 > 0 else qty
    if _roll_len <= 0:
        _roll_len = _qty_total_s1 if _qty_total_s1 > 0 else max(qty, 1e-9)
    return float(_roll_len)
def _write_stage1_exclude_rules_json_sidecar(
    wb_path_arg: str, out_path: str, *, use_effective_read_path: bool = True
) -> str | None:
    """設定_配台不要工程を UTF-8 JSON（{\"rules\": [...]}）に書き、絶対パスを返す。失敗時 None。

    ``use_effective_read_path=True``（既定）のときは保守直後の実効パス
    （``_exclude_rules_effective_read_path``）を優先する。

    段階1の ``stage1_exclude_rules.json`` は **master.xlsm** を読み、``False`` で呼ぶ。
    """
    try:
        if use_effective_read_path:
            path = _resolve_exclude_rules_workbook_path_for_read(wb_path_arg)
        else:
            path = (
                os.path.normpath(os.path.abspath(wb_path_arg))
                if wb_path_arg and os.path.isfile(wb_path_arg)
                else ""
            )
        if not path or not os.path.isfile(path):
            return None
        if _workbook_should_skip_openpyxl_io(path):
            logging.warning(
                "段階1: 配台不要ルールの JSON 化をスキップ（ブックに openpyxl 非互換シート）。",
            )
            return None
        df = pd.read_excel(path, sheet_name=EXCLUDE_RULES_SHEET_NAME)
    except Exception as ex:
        logging.warning(
            "段階1: 「%s」シートの読込に失敗し JSON を書けません: %s",
            EXCLUDE_RULES_SHEET_NAME,
            ex,
        )
        return None
    df.columns = df.columns.str.strip()
    need = [EXCLUDE_RULE_COL_PROCESS, EXCLUDE_RULE_COL_MACHINE]
    for c in need:
        if c not in df.columns:
            logging.warning(
                "段階1: 配台不要 JSON に必要な列「%s」が無いためスキップ。",
                c,
            )
            return None
    rules_out: list[dict] = []
    for _, row in df.iterrows():
        proc = str(row.get(EXCLUDE_RULE_COL_PROCESS, "") or "").strip()
        if not proc:
            continue
        rec: dict = {}
        for col in df.columns:
            k = str(col).strip()
            rec[k] = _json_safe_cell_for_exclude_rules_export(row.get(col))
        rules_out.append(rec)
    abs_out = os.path.abspath(out_path)
    parent = os.path.dirname(abs_out)
    try:
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(abs_out, "w", encoding="utf-8", newline="\n") as f:
            json.dump({"rules": rules_out}, f, ensure_ascii=False, indent=2)
    except OSError as ex:
        logging.warning("段階1: 配台不要 JSON の書き込みに失敗: %s", ex)
        return None
    return abs_out
def _write_stage1_task_input_preview_xlsx(df: pd.DataFrame, out_dir: str) -> str | None:
    """
    タスク一覧（plan_input_tasks）化の手前: ヘッダー行・列補完まで済んだ `load_tasks_df` の内容を保存する。
    Power Query の「依頼NO が空でない行」のフィルタに揃え、メタ行除去後の表を確認する用途。
    """
    if df is None or getattr(df, "empty", True):
        return None
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError:
        pass
    out_path = os.path.join(out_dir, STAGE1_TASK_INPUT_PREVIEW_FILENAME)
    if TASK_COL_TASK_ID in df.columns:
        _mask = df[TASK_COL_TASK_ID].map(
            lambda v: bool(planning_task_id_str_from_scalar(v))
        )
        df_out = df.loc[_mask].copy()
    else:
        df_out = df.copy()
    df_out.to_excel(out_path, sheet_name=STAGE1_TASK_INPUT_PREVIEW_SHEET, index=False)
    logging.info(
        "段階1: タスク入力整形プレビュー '%s'（シート「%s」%s 行）を出力しました。",
        out_path,
        STAGE1_TASK_INPUT_PREVIEW_SHEET,
        len(df_out),
    )
    return out_path
def run_stage1_extract():
    """
    段階1: 加工計画DATA から配台用タスク一覧を抽出し output/plan_input_tasks.xlsx へ出力。
    同一依頼NOで同一機械名が複数行あるとき、工程名「分割」行の空の「配台不要」に yes を自動設定する。
    マクロブックの「設定_配台不要工程」で工程+機械ごとの配台不要・条件式（AI）を管理する（シート作成は VBA）。
    設定シートの行同期および D 列→E 列（ロジック式）の AI 補完は、計画 DataFrame 確定後かつ
    「配台試行順番」の付与より前に行う。

    配台不要ルールの**正本**は UTF-8 JSON（list または ``{"rules":[...]}``）。作業先は
    ``PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK`` と同一フォルダの
    ``stage1_exclude_rules.json``（無ければリポジトリ同梱
    ``code/exclude_rules.json`` / ``code/json/stage1_exclude_rules.json`` から初回コピー）。
    有効な ``PM_AI_EXCLUDE_RULES_JSON`` があれば Excel の
    「設定_配台不要工程」は読まない。どちらも無い場合のみ master.xlsm の当該シートから
    作業先へ書き出して同変数を設定する。

    ``PM_AI_EXCLUDE_RULES_JSON`` が有効なときは、上記 JSON に
    ``_merge_exclude_rules_json_with_plan_pairs`` で工程+機械の行同期（追記）を行う。
    JSON を使わない場合に限り ``PM_AI_PLAN_INPUT_PATH`` が Excel ブック実ファイルのとき、そのブックへ
    ``run_exclude_rules_sheet_maintenance``（行同期・D→E）を行う。
    """
    resolve_processing_plan_path_from_env()
    _proc_plan = (os.environ.get(ENV_PROCESSING_PLAN_PATH) or "").strip()
    _has_processing_plan = bool(_proc_plan and os.path.isfile(_proc_plan))
    if not _has_processing_plan:
        logging.error(
            "段階1: タスク入力が解決できません。%s を実在ファイルにするか、"
            "%s から表ファイルを解決できるようにしてください。"
            % (ENV_PROCESSING_PLAN_PATH, "PM_AI_TASK_INPUT_SOURCE_DIR")
        )
        _tid0 = (os.environ.get("PM_AI_TASK_INPUT_SOURCE_DIR") or "").strip()
        print(
            "[stage1-error] "
            + (
                "段階1: タスク入力が解決できません。%s を実在ファイルにするか、"
                "%s から表ファイルを解決できるようにしてください。"
                % (ENV_PROCESSING_PLAN_PATH, "PM_AI_TASK_INPUT_SOURCE_DIR")
            ),
            file=sys.stderr,
            flush=True,
        )
        print(
            "[stage1-input] 現在の環境: %s=%r PM_AI_TASK_INPUT_SOURCE_DIR=%r"
            % (ENV_PROCESSING_PLAN_PATH, _proc_plan, _tid0),
            file=sys.stderr,
            flush=True,
        )
        return False
    reset_gemini_usage_tracker()
    logging.info(
        "段階1: material_table_append_build=%s _core=%s cwd=%s",
        _STAGE1_MATERIAL_TABLE_APPEND_BUILD,
        __file__,
        os.getcwd(),
    )
    df_src = load_tasks_df()
    try:
        _write_stage1_task_input_preview_xlsx(df_src, output_dir)
    except Exception as ex:
        logging.warning("段階1: タスク入力整形プレビューの出力をスキップ: %s", ex)
    rw_table, rw_known_keys, rw_table_path = _load_raw_fabric_width_mm_table()
    rw_appended_keys: set[str] = set()
    pw_table, pw_known_keys, pw_table_path = _load_product_width_mm_table()
    pw_appended_keys: set[str] = set()
    pl_table, pl_known_keys, pl_table_path = _load_product_length_mm_table()
    pl_appended_keys: set[str] = set()
    pt_table, pt_known_keys, pt_table_path = _load_product_thickness_mm_table()
    pt_write_path = _resolve_product_thickness_table_path_for_write()
    pt_appended_keys: set[str] = set()
    logging.info(
        "段階1: 製品厚みテーブル 読込=%s 追記先(正本)=%s",
        pt_table_path,
        pt_write_path,
    )
    _ur_roll_table, ur_roll_known_keys, ur_roll_table_path = (
        _load_used_raw_roll_length_table_stage1()
    )
    ur_roll_appended_keys: set[str] = set()
    records = []
    for _, row in df_src.iterrows():
        if row_has_completion_keyword(row):
            continue
        if _plan_row_exclude_as_completed_mikan_unprocessed_zero_actual_done_rule(row):
            continue
        task_id = planning_task_id_str_from_scalar(row.get(TASK_COL_TASK_ID))
        machine = str(row.get(TASK_COL_MACHINE, "")).strip()
        machine_name = str(row.get(TASK_COL_MACHINE_NAME, "")).strip()
        qty, _, _, _ = _plan_row_dispatch_qty_metrics(row)
        if qty <= 0 or not machine or not task_id:
            continue
        rec = {c: row.get(c) for c in SOURCE_BASE_COLUMNS}
        rec[TASK_COL_TASK_ID] = task_id
        _roll_len_product = _stage1_roll_length_for_planning_row(row)
        rec[PLAN_COL_ROLL_UNIT_LENGTH] = _roll_len_product
        _pw_mm = _resolve_product_width_mm_for_stage1_row(
            row,
            pw_table,
            table_path=pw_table_path,
            known_keys=pw_known_keys,
            appended=pw_appended_keys,
        )
        rec[PLAN_COL_PRODUCT_WIDTH] = _pw_mm if _pw_mm is not None else ""
        _pl_mm = _resolve_product_length_mm_for_stage1_row(
            row,
            pl_table,
            table_path=pl_table_path,
            known_keys=pl_known_keys,
            appended=pl_appended_keys,
        )
        rec[PLAN_COL_PRODUCT_LENGTH] = _pl_mm if _pl_mm is not None else ""
        _th_mm = _resolve_product_thickness_mm_for_stage1_row(
            row,
            pt_table,
            table_path=pt_write_path,
            known_keys=pt_known_keys,
            appended=pt_appended_keys,
        )
        rec[PLAN_COL_PRODUCT_THICKNESS] = _th_mm if _th_mm is not None else ""
        _rw_mm = _resolve_raw_fabric_width_mm_for_stage1_row(
            row,
            rw_table,
            table_path=rw_table_path,
            known_keys=rw_known_keys,
            appended=rw_appended_keys,
        )
        rec[PLAN_COL_RAW_FABRIC_WIDTH] = _rw_mm if _rw_mm is not None else ""
        _ur_cell = row.get(TASK_COL_USED_RAW)
        _raw_roll_tab_m = _lookup_roll_unit_length_m_from_used_raw(_ur_cell)
        if _raw_roll_tab_m is not None and float(_raw_roll_tab_m) > 0:
            rec[PLAN_COL_RAW_ROLL_UNIT_LENGTH] = float(_raw_roll_tab_m)
        else:
            _raw_dim_m = _parse_roll_unit_m_from_used_raw_dimension_only(_ur_cell)
            if _raw_dim_m is not None and _raw_dim_m > 0:
                rec[PLAN_COL_RAW_ROLL_UNIT_LENGTH] = float(_raw_dim_m)
            else:
                if _append_used_raw_roll_length_table_row_if_missing(
                    _ur_cell,
                    table_path=ur_roll_table_path,
                    known_keys=ur_roll_known_keys,
                    appended=ur_roll_appended_keys,
                ):
                    tid = planning_task_id_str_from_scalar(row.get(TASK_COL_TASK_ID))
                    logging.warning(
                        "使用原反ロール長未登録（テーブル・寸法から解釈不可）。材料テーブルへ追記。(原反)ロール単位長さは「不明」。依頼NO=%s 使用原反=%r",
                        tid,
                        _normalize_mm_table_lookup_key(_ur_cell),
                    )
                rec[PLAN_COL_RAW_ROLL_UNIT_LENGTH] = "不明"
        # 換算数量列は加工計画DATAの値のまま（§7.1）。配台使用残数量・配台ロール数は段階1列で埋める。
        if TASK_COL_QTY in rec:
            rec[TASK_COL_QTY] = max(0.0, parse_float_safe(row.get(TASK_COL_QTY), 0.0))
        # 工程名 + 機械名 を“因孝”として表示用に追加（後段は計算キーにも使用）
        if machine_name:
            rec[PLAN_COL_PROCESS_FACTOR] = f"{machine}+{machine_name}"
        else:
            rec[PLAN_COL_PROCESS_FACTOR] = f"{machine}+"
        _raw_for_dispatch = parse_optional_date(rec.get(TASK_COL_RAW_INPUT_DATE))
        rec[PLAN_COL_DISPATCHABLE_DATETIME] = format_dispatchable_datetime_cell(
            compute_dispatchable_datetime(_raw_for_dispatch)
        )
        rec[PLAN_COL_PREFERRED_OP] = ""
        rec[PLAN_COL_SPECIAL_REMARK] = ""
        rec[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = ""
        rec[PLAN_COL_AI_PARSE] = ""
        records.append(rec)
    if not records:
        logging.warning("段階1: 抽出対象タスクはありません。")
    order = plan_input_sheet_column_order()
    out_df = pd.DataFrame(records)
    if out_df.empty:
        out_df = pd.DataFrame(columns=order)
    else:
        out_df = out_df.reindex(columns=order).fillna("")
    if PLAN_COL_EXCLUDE_FROM_ASSIGNMENT in out_df.columns:
        out_df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = out_df[
            PLAN_COL_EXCLUDE_FROM_ASSIGNMENT
        ].astype(object)
    if RESULT_TASK_COL_DISPATCH_TRIAL_ORDER in out_df.columns:
        out_df[RESULT_TASK_COL_DISPATCH_TRIAL_ORDER] = ""
        out_df[RESULT_TASK_COL_DISPATCH_TRIAL_ORDER] = out_df[
            RESULT_TASK_COL_DISPATCH_TRIAL_ORDER
        ].astype(object)
    _members_stage1: list = []
    try:
        (
            _skills_d_stage1,
            _members_stage1,
            equipment_list_stage1,
            req_map,
            need_rules,
            _surplus_map_stage1,
            need_combo_col_index_stage1,
        ) = load_skills_and_needs()
    except PlanningValidationError:
        logging.error("段階1を中断: マスタ skills / 勤怠シートの検証エラー。")
        raise
    out_df = _merge_plan_sheet_user_overrides(out_df)
    _apply_master_speed_sheet_to_plan_df(out_df, log_prefix="段階1")
    try:
        from .actual_speed_apply import apply_learned_speed_to_plan_df

        apply_learned_speed_to_plan_df(out_df, log_prefix="段階1")
    except Exception as ex:
        logging.warning("段階1: 学習速度適用をスキップ（%s）", ex)
    _heal_stage1_roll_unit_no_dim_when_roll_matches_qty_mistake(out_df)
    _heal_stage1_roll_unit_if_width_ceiling_merge_spurious(out_df)
    _refresh_plan_reference_columns(out_df, req_map, need_rules)
    try:
        _apply_auto_exclude_bunkatsu_duplicate_machine(out_df, log_prefix="段階1")
    except Exception as ex:
        logging.exception("段階1: 分割行の配台不要自動設定で例外（出力は続行）: %s", ex)
    _ensure_dispatch_lookup_tables_at_work_path()
    _ensure_stage1_exclude_rules_json_env_from_repo_default()
    # 計画行確定後・試行順より前: JSON 正本なら JSON への行同期、それ以外は計画ブックの「設定_配台不要工程」（D→E 含む）。
    try:
        _pm_pairs_s1 = _collect_plan_input_process_machine_pairs_for_exclude_rules_sync(
            out_df
        )
        if _exclude_rules_json_env_supersedes_excel_sheet():
            json_env = (os.environ.get(ENV_EXCLUDE_RULES_JSON) or "").strip()
            if json_env and _pm_pairs_s1:
                _merge_exclude_rules_json_with_plan_pairs(json_env, _pm_pairs_s1, "段階1")
        else:
            _wb_maint_s1 = _excel_plan_input_wb()
            if not _wb_maint_s1 or not os.path.isfile(_wb_maint_s1):
                logging.info(
                    "段階1: 設定シート「%s」の Excel 保守をスキップ（%s がブック実ファイルでない）。",
                    EXCLUDE_RULES_SHEET_NAME,
                    ENV_PLAN_INPUT_PATH,
                )
            else:
                run_exclude_rules_sheet_maintenance(_wb_maint_s1, _pm_pairs_s1, "段階1")
    except Exception:
        logging.exception("段階1: 設定_配台不要工程の保守で例外（続行）")
    _master_er_wb = _require_master_workbook_path_exists()
    try:
        if _exclude_rules_json_env_supersedes_excel_sheet():
            logging.info(
                "段階1: 配台不要は JSON のみのため master からの %s 書き出しをスキップします。",
                STAGE1_EXCLUDE_RULES_JSON_FILENAME,
            )
        else:
            _s1_er_json = _resolve_stage1_exclude_rules_json_work_path()
            _written_er = _write_stage1_exclude_rules_json_sidecar(
                _master_er_wb, _s1_er_json, use_effective_read_path=False
            )
            if _written_er:
                os.environ[ENV_EXCLUDE_RULES_JSON] = _written_er
                _reset_exclude_rules_json_env_memo()
                logging.info(
                    "段階1: 配台不要ルール（master）を JSON に書き出し、%s=%s",
                    ENV_EXCLUDE_RULES_JSON,
                    _written_er,
                )
    except Exception:
        logging.exception("段階1: 配台不要ルールの JSON 書き出しで例外（続行）")
    try:
        out_df = apply_exclude_rules_config_to_plan_df(out_df, _master_er_wb, "段階1")
    except Exception as ex:
        logging.warning("段階1: 設定シートによる配台試行適用で例外（続行）: %s", ex)
    try:
        _ext_dt_s1, _ = _extract_data_extraction_datetime()
        _run_d_s1 = _ext_dt_s1.date() if _ext_dt_s1 is not None else datetime.now().date()
        fill_plan_dispatch_trial_order_column_stage1(
            out_df,
            _run_d_s1,
            req_map,
            need_rules,
            need_combo_col_index_stage1,
            equipment_list_stage1,
            members_for_gpo=_members_stage1,
        )
    except Exception as ex:
        logging.warning("段階1: 配台試行順番列の計算をスキップしました（続行）: %s", ex)
    out_df = _sort_stage1_plan_df_by_dispatch_trial_order_asc(out_df)
    _fill_plan_dispatch_remaining_qty_column(out_df)
    _apply_stage1_in_progress_dispatch_plan_exclude_marker(out_df, log_prefix="段階1")
    out_path = os.path.join(output_dir, STAGE1_OUTPUT_FILENAME)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name=STAGE1_PLAN_OUTPUT_SHEET, index=False)
    normalize_ooxml_shared_strings_if_missing(out_path)
    _apply_excel_date_columns_date_only_display(out_path, STAGE1_PLAN_OUTPUT_SHEET)
    _apply_plan_input_visual_format(out_path, STAGE1_PLAN_OUTPUT_SHEET)
    logging.info(f"段階1完了: '{out_path}' を出力しました。マクロで '{PLAN_INPUT_SHEET_NAME}' に坖り込んでしてさい。")
    _try_write_main_sheet_gemini_usage_summary("段階1")
    logging.info(
        "段階1: 製品厚みテーブル追記サマリ build=%s 追記先=%s 追記キー数=%s",
        _STAGE1_MATERIAL_TABLE_APPEND_BUILD,
        pt_write_path,
        len(pt_appended_keys),
    )
    if pt_appended_keys:
        logging.info(
            "段階1: 製品厚みテーブルへ追記したキー: %s",
            ", ".join(sorted(pt_appended_keys)[:50]),
        )
    return True
TARGET_YEAR = 2026
TARGET_MONTH = 3
DEFAULT_START_TIME = time(8, 45)
DEFAULT_END_TIME = time(17, 0)
DEFAULT_BREAKS = [
    (time(12, 0), time(12, 50)),
    (time(14, 45), time(15, 0))
]
ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS = max(
    0,
    int(os.environ.get("ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS", "5").strip() or 0),
)
ASSIGN_END_OF_DAY_DEFER_MINUTES = max(
    0,
    int(os.environ.get("ASSIGN_END_OF_DAY_DEFER_MINUTES", "45").strip() or 0),
)

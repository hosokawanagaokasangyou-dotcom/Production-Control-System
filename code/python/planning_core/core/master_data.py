# -*- coding: utf-8 -*-
# planning_core.core.master_data — body only (loaded via _core exec chain)
def load_skills_and_needs():
    """
    統合ファイル(MASTER_FILE)からスキルと need を動的に読み込みした。

    戻り値は7覝素。最後は need シート上の「工程名+機械名」列佝置（左ろど尝さい整数）の辞書
    ``need_combo_col_index``（配台キューソート用）。

    今回の need は（Excel上で）
      工程名行・機械名行のあと「基本必須人数」行（A列に「必須人数」を含む）
      しの直下: 配台で余剰人員はあるとしに追加で入れられる人数（工程×機械とと。未設定は 0）
      以降: 特別指定1〜99
    といご構造のため、必須OPは「工程名+機械名」で解決れる。

    skills 交差セルは OP/AS の後に優先度整数（例 OP1, AS3）。数値は尝さいろど当該工程への割当で優先。
    数字省略の OP/AS は優先度 1。
    同一列（同一工程×機械）では優先度の数値はメンバー間で重複試行（重複時は PlanningValidationError）。
    skills メンバー全員に同名の勤怠シートが無い場合も PlanningValidationError。
    """
    mp = _require_master_workbook_path_exists()
    try:
        # 同一ブックを pd.read_excel で都度開しと I/O は重いめ、ExcelFile を1回の値開いでシートを parse れる。
        _master_xls = _cached_master_pd_excel_file(mp)
        if _master_xls is None:
            raise FileNotFoundError(f"マスタブックを開けません: {mp}")
        if True:
            # skills は新仕様:
            #   1行目: 工程名
            #   2行目: 機械名
            #   A3以降: メンバー坝
            #   交差セル: OP または AS の後に割当優先度の整数（例 OP1, AS3）。数値は尝さいろど当該工程へ優先割当。
            #             数字省略の OP/AS は優先度 1（従来どおり最優先扱い）。
            # を基本としつつ」旧仕様（1行ヘッダ）にもフォールバック対応れる。
            skills_raw = pd.read_excel(
                _master_xls, sheet_name="skills", header=None
            )
            skills_dict = {}
            equipment_list = []
            members = []

            use_two_header = False
            if skills_raw.shape[0] >= 3 and skills_raw.shape[1] >= 2:
                non_empty_pm = 0
                for c in range(1, skills_raw.shape[1]):
                    p = skills_raw.iat[0, c]
                    m = skills_raw.iat[1, c]
                    if pd.isna(p) or pd.isna(m):
                        continue
                    p_s = str(p).strip()
                    m_s = str(m).strip()
                    if p_s and m_s and p_s.lower() != "nan" and m_s.lower() != "nan":
                        non_empty_pm += 1
                use_two_header = non_empty_pm > 0

            if use_two_header:
                pm_cols = []
                seen_combo = set()
                for c in range(1, skills_raw.shape[1]):
                    p = skills_raw.iat[0, c]
                    m = skills_raw.iat[1, c]
                    if pd.isna(p) or pd.isna(m):
                        continue
                    p_s = str(p).strip()
                    m_s = str(m).strip()
                    if not p_s or not m_s or p_s.lower() == "nan" or m_s.lower() == "nan":
                        continue
                    combo = f"{p_s}+{m_s}"
                    pm_cols.append((c, p_s, m_s, combo))
                    if combo not in seen_combo:
                        seen_combo.add(combo)
                        equipment_list.append(combo)

                for r in range(2, skills_raw.shape[0]):
                    m_name_raw = skills_raw.iat[r, 0]
                    if pd.isna(m_name_raw):
                        continue
                    m_name = str(m_name_raw).strip()
                    if not m_name or m_name.lower() in ("nan", "none", "null"):
                        continue
                    row_skills = {}
                    for c, p_s, m_s, combo in pm_cols:
                        v = skills_raw.iat[r, c] if c < skills_raw.shape[1] else None
                        sval = "" if pd.isna(v) else str(v).strip()
                        if not sval or sval.lower() in ("nan", "none", "null"):
                            continue
                        row_skills[combo] = sval
                        if m_s not in row_skills:
                            row_skills[m_s] = sval
                        if p_s not in row_skills:
                            row_skills[p_s] = sval
                    skills_dict[m_name] = row_skills
                members = list(skills_dict.keys())
                logging.info(
                    "skillsシート: 2段ヘッダ形式で読み込みました（工程+機械=%s列, メンバー=%s人）。",
                    len(pm_cols),
                    len(members),
                )
            else:
                skills_df = pd.read_excel(_master_xls, sheet_name="skills")
                skills_df.columns = skills_df.columns.str.strip()
                skill_cols = [
                    str(c).strip()
                    for c in skills_df.columns
                    if not str(c).startswith("Unnamed")
                ]

                member_col = None
                for c in skill_cols:
                    if c in ("メンバー", "担当者", "並び", "作業者"):
                        member_col = c
                        break
                if member_col is None and skill_cols:
                    member_col = skill_cols[0]
                    logging.warning(
                        "skillsシート: メンバー列名は標準と一致しないため、先頭列 '%s' をメンバー列として扱いした。",
                        member_col,
                    )

                seen_eq = set()
                for c in skill_cols:
                    if c == member_col:
                        continue
                    cid = str(c).strip()
                    if not cid or cid.lower() in ("nan", "none", "null"):
                        continue
                    if cid not in seen_eq:
                        seen_eq.add(cid)
                        equipment_list.append(cid)

                for _, row in skills_df.iterrows():
                    m_name = str(row.get(member_col, "")).strip() if member_col else ""
                    if not m_name or m_name.lower() == "nan":
                        continue
                    row_skills = {}
                    for c in skill_cols:
                        if c == member_col:
                            continue
                        sval = str(row.get(c, "")).strip()
                        if not sval or sval.lower() in ("nan", "none", "null"):
                            continue
                        row_skills[c] = sval
                        if "+" in c:
                            p, m = c.split("+", 1)
                            p = p.strip()
                            m = m.strip()
                            if m and m not in row_skills:
                                row_skills[m] = sval
                            if p and p not in row_skills:
                                row_skills[p] = sval
                    skills_dict[m_name] = row_skills
                members = list(skills_dict.keys())
                logging.info(
                    "skillsシート: 1行ヘッダ形式（旧互換）で読み込みました（メンバー=%s人）。",
                    len(members),
                )

            if not members:
                logging.error("skillsシートからメンバーを読み込ゝませんでした。")
            else:
                _validate_skills_op_as_priority_numbers_unique(
                    skills_dict, equipment_list
                )
                from planning_core.core.attendance_readiness import (
                    legacy_master_attendance_sheets_required,
                )

                if legacy_master_attendance_sheets_required():
                    _validate_skills_members_have_attendance_sheets(members, mp)

            # need は header=None で読み」先頭の複数行を“見出し行”として解釈
            needs_raw = pd.read_excel(
                _master_xls, sheet_name="need", header=None
            )

        col0 = 0
        process_header_row = None
        machine_header_row = None
        base_row = None

        for r in range(needs_raw.shape[0]):
            v0 = needs_raw.iat[r, col0]
            if pd.isna(v0):
                continue
            s0 = str(v0).strip()
            if process_header_row is None and s0 in ("工程名", "工程名"):
                process_header_row = r
            elif machine_header_row is None and s0 in ("機械名", "機械名"):
                machine_header_row = r
            if base_row is None and not s0.startswith("特別指定"):
                if "必要人数" in s0 or "必須人数" in s0:
                    base_row = r
            if process_header_row is not None and machine_header_row is not None and base_row is not None:
                break

        if process_header_row is None or machine_header_row is None or base_row is None:
            raise ValueError(
                "need シートのヘッダー行が見つかりません。"
                " A列に 工程名/機械名（旧テンプレ: 工程名/機械名）と、"
                " 基本必要人数（旧: 基本必須人数 など「必要人数」または「必須人数」を含む行）が必要です。"
            )

        # 「依頼NO条件」列佝置（デフォルトは 1列目）
        cond_col_idx = 1
        for r in range(needs_raw.shape[0]):
            c1 = needs_raw.iat[r, 1] if needs_raw.shape[1] > 1 else None
            c2 = needs_raw.iat[r, 2] if needs_raw.shape[1] > 2 else None
            if pd.isna(c1) or pd.isna(c2):
                continue
            if str(c1).strip() == NEED_COL_CONDITION and str(c2).strip() == NEED_COL_NOTE:
                cond_col_idx = 1
                break

        # 工程名×機械名 の列一覧（列番坷は Excel上の実列を保挝）
        pm_cols = []
        for col_idx in range(needs_raw.shape[1]):
            if col_idx < 3:
                continue
            p = needs_raw.iat[process_header_row, col_idx]
            m = needs_raw.iat[machine_header_row, col_idx]
            if pd.isna(p) or pd.isna(m):
                continue
            p_s = _normalize_equipment_match_key(str(p).strip())
            m_s = _normalize_equipment_match_key(str(m).strip())
            if not p_s or not m_s or p_s.lower() == "nan" or m_s.lower() == "nan":
                continue
            pm_cols.append((col_idx, p_s, m_s))

        req_map = {}
        # 工程名+機械名コンボ → need シート上の列インデックス（左ろど尝さい＝配台キューで先）
        need_combo_col_index: dict[str, int] = {}
        # need_rules: [{'order': int, 'condition': str, 'overrides': {combo_key/machine/process: int}}]
        need_rules = []

        # 基本必須人数
        for col_idx, p_s, m_s in pm_cols:
            n = parse_optional_int(needs_raw.iat[base_row, col_idx])
            if n is None or n < 1:
                n = 1
            combo_key = f"{p_s}+{m_s}"
            need_combo_col_index[combo_key] = col_idx
            req_map[combo_key] = n
            # フォールバック用（機械名 or 工程名の値で引けるよごにれる）
            if p_s not in req_map:
                req_map[p_s] = n
            if m_s not in req_map:
                req_map[m_s] = n

        surplus_map: dict[str, int] = {}
        surplus_row = _find_need_surplus_add_row_index(
            needs_raw, base_row, col0, pm_cols
        )
        if surplus_row is not None:
            for col_idx, p_s, m_s in pm_cols:
                raw_ex = parse_optional_int(needs_raw.iat[surplus_row, col_idx])
                ex = int(raw_ex) if raw_ex is not None and raw_ex >= 0 else 0
                ex = max(0, min(ex, 50))
                combo_key = f"{p_s}+{m_s}"
                surplus_map[combo_key] = ex
                if p_s not in surplus_map:
                    surplus_map[p_s] = ex
                if m_s not in surplus_map:
                    surplus_map[m_s] = ex
            logging.info(
                "need シート: 配台時追加人数行を検出（Excel行≈%s）。列ととの上限を読み込みました。",
                surplus_row + 1,
            )
        else:
            logging.info(
                "need シート: 基本必須人数の直下に配台時追加人数行を検出でしませんでした（省略坯）。"
            )

        if TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW:
            logging.info(
                "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW は有効: 配台時追加人数は読み込んでも常に 0 扱い（フォームは基本必須人数のみ試行）。"
            )

        logging.info(
            "need人数マスタ: %s の need シートを読み込み（skills と同一 ExcelFile で開いた直後。need 専用ディスクキャッシュは無し・AI json とは無関係）。",
            mp,
        )
        for _ci, _ps, _ms in pm_cols:
            _ck = f"{_ps}+{_ms}"
            _bn = req_map.get(_ck)
            _sx = surplus_map.get(_ck, 0) if surplus_map else 0
            logging.info(
                "need列サマリ combo=%s 基本必須人数=%s 配台時追加人数上限=%s",
                _log_map_key_label(_ck),
                _bn,
                _sx,
            )

        # 特別指定
        for r in range(needs_raw.shape[0]):
            v0 = needs_raw.iat[r, col0]
            if pd.isna(v0):
                continue
            lab = str(v0).strip()
            m = re.match(r"特別指定\s*(\d+)", lab)
            if not m:
                continue
            order = int(m.group(1))
            if order < 1 or order > 99:
                continue

            cond_raw = needs_raw.iat[r, cond_col_idx] if needs_raw.shape[1] > cond_col_idx else None
            cond = "" if pd.isna(cond_raw) else str(cond_raw).strip()

            overrides = {}
            for col_idx, p_s, m_s in pm_cols:
                v = needs_raw.iat[r, col_idx]
                n = parse_optional_int(v)
                if n is not None and 1 <= n <= 99:
                    combo_key = f"{p_s}+{m_s}"
                    overrides[combo_key] = n
                    # フォールバック用
                    overrides[p_s] = n
                    overrides[m_s] = n

            if overrides:
                need_rules.append({"order": order, "condition": cond, "overrides": overrides})

        need_rules.sort(key=lambda rr: rr["order"])
        logging.info(f"need 特別指定ルール: {len(need_rules)} 件（工程名+機械名キー）。")

        logging.info(f"『{MASTER_FILE}」からスキルと設備覝件(need)を読み込みました。")
        return (
            skills_dict,
            members,
            equipment_list,
            req_map,
            need_rules,
            surplus_map,
            need_combo_col_index,
        )

    except PlanningValidationError:
        raise
    except Exception as e:
        logging.error(f"マスタファイル({MASTER_FILE})のスキル/need読み込みエラー: {e}")
        raise


def load_need_machine_columns() -> list[dict[str, str]]:
    """need シートの工程名×機械名列（機械カレンダー UI 列の正本）。"""
    mp = _require_master_workbook_path_exists()
    _master_xls = _cached_master_pd_excel_file(mp)
    if _master_xls is None:
        raise FileNotFoundError(f"マスタブックを開けません: {mp}")
    needs_raw = pd.read_excel(_master_xls, sheet_name="need", header=None)
    process_header_row = None
    machine_header_row = None
    for r in range(needs_raw.shape[0]):
        v0 = needs_raw.iat[r, 0]
        if pd.isna(v0):
            continue
        s0 = str(v0).strip()
        if process_header_row is None and s0 in ("工程名", "工程名"):
            process_header_row = r
        elif machine_header_row is None and s0 in ("機械名", "機械名"):
            machine_header_row = r
        if process_header_row is not None and machine_header_row is not None:
            break
    if process_header_row is None or machine_header_row is None:
        raise ValueError(
            "need シートのヘッダー行が見つかりません。"
            " A列に工程名・機械名行が必要です。"
        )
    columns: list[dict[str, str]] = []
    seen: set[str] = set()
    for col_idx in range(needs_raw.shape[1]):
        if col_idx < 3:
            continue
        p = needs_raw.iat[process_header_row, col_idx]
        m = needs_raw.iat[machine_header_row, col_idx]
        if pd.isna(p) or pd.isna(m):
            continue
        p_s = _normalize_equipment_match_key(str(p).strip())
        m_s = _normalize_equipment_match_key(str(m).strip())
        if not p_s or not m_s or p_s.lower() == "nan" or m_s.lower() == "nan":
            continue
        combo_key = f"{p_s}+{m_s}"
        if combo_key in seen:
            continue
        seen.add(combo_key)
        columns.append(
            {"equipment_key": combo_key, "process": p_s, "machine": m_s}
        )
    return columns


def _combo_preset_member_name(raw: str) -> str:
    """組み合わせ表メンバーセルから配台用の氏名を取る。『OP 山田』『AS1 佐藤』は氏名のみ。"""
    s = "" if raw is None else str(raw).strip()
    if not s:
        return ""
    stripped = re.sub(r"^(?:OP|AS)\s*\d*\s+", "", s, count=1, flags=re.IGNORECASE).strip()
    return stripped or s


def load_team_combination_presets_from_master() -> dict[
    str, list[tuple[int, int | None, tuple[str, ...], int | None]]
]:
    """
    master.xlsm「組み合わせ表」を読み」工程+機械キーごとに
    [(組み合わせ優先度, 必須人数またはNone, メンバータプル, 組み合わせ行IDまたはNone), ...] を返す。
    同一キー内は優先度昇順」坌順佝はシート上の行順。
    「必須人数」列は配台時に need 基本人数より優先れる（メンバー列人数と一致すること）。
    配台では成立したプリセットをまとめて候補に載せ」組み合わせ探索とまとめで team_start 等で最良を決める
    （シート優先度は試行順のみ。先頭プリセットの坳決はしない）。
    A 列「組み合わせ行ID」は無い＝空の旧シートでは ID は None。
    """
    if not TEAM_ASSIGN_USE_MASTER_COMBO_SHEET:
        return {}
    path = _master_workbook_path_resolved()
    if not os.path.isfile(path):
        return {}
    try:
        xls = _cached_master_pd_excel_file(path)
        if xls is None:
            return {}
        df = pd.read_excel(xls, sheet_name=MASTER_SHEET_TEAM_COMBINATIONS, header=0)
    except Exception as e:
        logging.info("組み合わせ表シートの読込をスキップしました: %s", e)
        return {}
    if df is None or df.empty:
        return {}

    def norm_cell(x) -> str:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).strip()

    colmap = {norm_cell(c): c for c in df.columns if norm_cell(c)}
    # 旧シート互換: 「組合せ」表記も許容
    id_c = (
        colmap.get("組み合わせ行ID")
        or colmap.get("組合せ行ID")
        or colmap.get("インデックス")
    )
    proc_c = colmap.get("工程名")
    mach_c = colmap.get("機械名")
    combo_c = colmap.get("工程+機械")
    prio_c = colmap.get("組み合わせ優先度") or colmap.get("組合せ優先度")
    req_c = colmap.get("必須人数") or colmap.get("必要人数")

    def mem_col_order(c) -> int:
        m = re.search(r"メンバー\s*(\d+)", norm_cell(c))
        return int(m.group(1)) if m else 9999

    mem_keys = sorted(
        [c for c in df.columns if norm_cell(str(c)).startswith("メンバー")],
        key=mem_col_order,
    )
    buckets: dict[
        str,
        list[tuple[int, int, int | None, tuple[str, ...], int | None]],
    ] = defaultdict(list)
    _cols = list(df.columns)
    _ix = {
        "proc": _cols.index(proc_c) if proc_c and proc_c in _cols else -1,
        "mach": _cols.index(mach_c) if mach_c and mach_c in _cols else -1,
        "combo": _cols.index(combo_c) if combo_c and combo_c in _cols else -1,
        "prio": _cols.index(prio_c) if prio_c and prio_c in _cols else -1,
        "req": _cols.index(req_c) if req_c and req_c in _cols else -1,
        "id": _cols.index(id_c) if id_c and id_c in _cols else -1,
        "mem": [_cols.index(mc) for mc in mem_keys if mc in _cols],
    }

    def _cell_at(row_tuple, col_index: int):
        if col_index < 0:
            return ""
        return norm_cell(row_tuple[1 + col_index])

    for row in df.itertuples(index=True, name=None):
        row_i = int(row[0])
        proc = _cell_at(row, _ix["proc"])
        mach = _cell_at(row, _ix["mach"])
        combo_cell = _cell_at(row, _ix["combo"])
        if proc and mach:
            key = f"{proc}+{mach}"
        elif combo_cell:
            key = combo_cell
        else:
            continue
        pr = parse_optional_int(_cell_at(row, _ix["prio"])) if _ix["prio"] >= 0 else None
        if pr is None:
            pr = 10**9
        sheet_req: int | None = None
        if _ix["req"] >= 0:
            sheet_req = parse_optional_int(_cell_at(row, _ix["req"]))
            if sheet_req is not None and sheet_req < 1:
                sheet_req = None
        sheet_combo_id: int | None = None
        if _ix["id"] >= 0:
            sheet_combo_id = parse_optional_int(_cell_at(row, _ix["id"]))
            if sheet_combo_id is not None and sheet_combo_id < 1:
                sheet_combo_id = None
        team: list[str] = []
        for mc_ix in _ix["mem"]:
            s = _cell_at(row, mc_ix)
            if not s or s.lower() in ("nan", "none", "null"):
                continue
            team.append(_combo_preset_member_name(s))
        if not team:
            continue
        buckets[key].append(
            (pr, row_i, sheet_req, tuple(team), sheet_combo_id)
        )

    out: dict[
        str, list[tuple[int, int | None, tuple[str, ...], int | None]]
    ] = {}
    for key, lst in buckets.items():
        lst.sort(key=lambda x: (x[0], x[1]))
        out[key] = [(t[0], t[2], t[3], t[4]) for t in lst]
    return out
def _lookup_combo_sheet_row_id_for_preset_team(
    preset_rows: list | None,
    team: tuple,
) -> int | None:
    """
    採用フォームのメンバー集合（NFKC・trim）は組み合わせ表プリセットのいうれかと一致するとし」
    しの行の組み合わせ行ID（A列）を返す。組み合わせ探索のみで決まり combo_sheet_row_id は付いでいない
    履歴行の補完に使う。複数一致時は組み合わせ優先度（数値は尝さい方）を採用。
    """
    if not preset_rows or not team:
        return None

    def _mem_key(members) -> frozenset:
        out: set[str] = set()
        for m in members:
            s = str(m).strip()
            if not s:
                continue
            out.add(unicodedata.normalize("NFKC", s))
        return frozenset(out)

    target = _mem_key(team)
    if not target:
        return None
    best_id: int | None = None
    best_prio: int | None = None
    for pr, _sheet_rs, preset_team, combo_row_id in preset_rows:
        if combo_row_id is None:
            continue
        if _mem_key(preset_team) != target:
            continue
        try:
            prio_val = int(pr)
        except (TypeError, ValueError):
            prio_val = 10**9
        if best_prio is None or prio_val < best_prio:
            best_prio = prio_val
            try:
                best_id = int(combo_row_id)
            except (TypeError, ValueError):
                best_id = None
    return best_id
def generate_default_calendar_dates(year, month):
    cal = calendar.Calendar()
    return [d for d in cal.itermonthdates(year, month) if d.year == year and d.month == month and d.weekday() < 5]
def parse_time_str(time_str, default_time):
    if time_str is None or pd.isna(time_str) or not str(time_str).strip() or str(time_str).strip().lower() == 'null':
        return default_time
    try:
        if isinstance(time_str, time): return time_str
        if isinstance(time_str, datetime): return time_str.time()
        time_str = str(time_str).strip()
        if len(time_str.split(':')) == 3:
            return datetime.strptime(time_str, "%H:%M:%S").time()
        return datetime.strptime(time_str, "%H:%M").time()
    except:
        return default_time
def _excel_scalar_to_time_optional(v) -> time | None:
    """master メインの時刻セル（datetime / time / 文字列）を time に。解釈試行は None。"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, time):
        return v
    if isinstance(v, datetime):
        return v.time()
    return parse_time_str(v, None)
def _pick_master_main_sheet_name(sheetnames: list[str]) -> str | None:
    """
    master.xlsm の「メイン」設定シート名を解決れる（VBA MasterGetMainWorksheet と同じ趣旨）。
    「〇月メインカレンダー」等を誤採用しないよご「カレンダー」を含む坝剝は除外し、
    複数候補はシート名は最短のものを優先れる。
    """
    for prefer in ("メイン", "メイン_", "Main"):
        if prefer in sheetnames:
            return prefer
    cand = [sn for sn in sheetnames if "メイン" in sn and "カレンダー" not in sn]
    if not cand:
        return None
    return min(cand, key=len)
def _read_master_main_factory_operating_times(master_path: str) -> tuple[time | None, time | None]:
    """
    master.xlsm のメインシート A12（稼働開始）・B12（稼働終了）を読む。
    いうれか欠損・正常・開始>=終了のときは (None, None)。
    """
    p = (master_path or "").strip()
    if not p or not os.path.isfile(p):
        return None, None
    if _workbook_should_skip_openpyxl_io(p):
        return None, None
    try:
        wb = load_workbook(p, data_only=True, read_only=False)
    except Exception as e:
        logging.warning("工場稼働時刻: master を openpyxl で開きませんでした（既定の日内枠を使用した）: %s", e)
        return None, None
    try:
        sn = _pick_master_main_sheet_name(list(wb.sheetnames))
        if sn is None:
            return None, None
        ws = wb[sn]
        st = _excel_scalar_to_time_optional(ws.cell(row=12, column=1).value)
        et = _excel_scalar_to_time_optional(ws.cell(row=12, column=2).value)
        if st is None or et is None:
            return None, None
        if st >= et:
            logging.warning(
                "工場稼働時刻: master メイン A12/B12 は開始>=終了 (%s >= %s) のため、既定値を使用した。",
                st,
                et,
            )
            return None, None
        return st, et
    finally:
        try:
            wb.close()
        except Exception:
            pass
def _read_master_main_regular_shift_times(master_path: str) -> tuple[time | None, time | None]:
    """
    master.xlsm のメインシート A15（定常開始）・B15（定常終了）を読む。
    いうれか欠損・正常・開始>=終了のときは (None, None)。
    """
    p = (master_path or "").strip()
    if not p or not os.path.isfile(p):
        return None, None
    if _workbook_should_skip_openpyxl_io(p):
        return None, None
    try:
        wb = load_workbook(p, data_only=True, read_only=False)
    except Exception as e:
        logging.warning(
            "定常時刻: master を openpyxl で開きませんでした（結果シートの定常外着色をスキップ）: %s",
            e,
        )
        return None, None
    try:
        sn = _pick_master_main_sheet_name(list(wb.sheetnames))
        if sn is None:
            return None, None
        ws = wb[sn]
        st = _excel_scalar_to_time_optional(ws.cell(row=15, column=1).value)
        et = _excel_scalar_to_time_optional(ws.cell(row=15, column=2).value)
        if st is None or et is None:
            return None, None
        if st >= et:
            logging.warning(
                "定常時刻: master メイン A15/B15 は開始>=終了 (%s >= %s) のため、着色・比較に使いません。",
                st,
                et,
            )
            return None, None
        return st, et
    finally:
        try:
            wb.close()
        except Exception:
            pass


@contextmanager
def _override_default_factory_hours_from_master(master_path: str):
    """段階2の間の値 DEFAULT_START_TIME / DEFAULT_END_TIME を master メイン A12/B12 で上書き。"""
    global DEFAULT_START_TIME, DEFAULT_END_TIME
    orig_s, orig_e = DEFAULT_START_TIME, DEFAULT_END_TIME
    ns, ne = _read_master_main_factory_operating_times(master_path)
    try:
        if ns is not None and ne is not None:
            DEFAULT_START_TIME = ns
            DEFAULT_END_TIME = ne
            logging.info(
                "工場稼働枠: master.xlsm メイン A12/B12 を採用 → %s ～ %s（結果_* の日内グリッド・配台枠）",
                DEFAULT_START_TIME.strftime("%H:%M"),
                DEFAULT_END_TIME.strftime("%H:%M"),
            )
        yield
    finally:
        DEFAULT_START_TIME, DEFAULT_END_TIME = orig_s, orig_e
def infer_mid_break_from_reason(reason_text, start_t, end_t, break1_start=None, break1_end=None):
    """
    備考から中抜き時間を推定するローカル補正。
    AIは中抜きを返さない場合のフェイルセーフとして使う。
    master.xlsm カレンダー由来の休暇区分: 公休=公休年休・午後のみ勤務」後休=午後年休・公休のみ勤務（出勤簿.txt と坌義）。
    公休・後休の境界はメンバー勤怠の休憩時間1_開始/終了（未指定時は DEFAULT_BREAKS[0]）に合わせる。
    """
    if reason_text is None:
        return None, None
    txt = str(reason_text).strip()
    if not txt or txt.lower() in ("nan", "none", "null", "通常"):
        return None, None

    b1_s = break1_start if break1_start is not None else DEFAULT_BREAKS[0][0]
    b1_e = break1_end if break1_end is not None else DEFAULT_BREAKS[0][1]

    noon_end = time(12, 0)
    afternoon_start = time(13, 0)
    # カレンダー記坷と一致させる（シフト時刻は誤っている場合の補完用。正しい行では区間は空になり追加されない）
    if txt == "公休":
        # 正しい行は出勤は休憩1終了以降で補完試行。全日シフトの誤入力時はしこまでを中抜き（公休年休相当）
        if start_t and start_t < b1_e:
            return start_t, b1_e
        return None, None
    if txt == "後休":
        if end_t and b1_s < end_t:
            return b1_s, end_t
        return None, None

    # 1) 明示的な時刻範囲（例: 11:00-14:00 / 11:00～14:00）
    m = re.search(r"(\d{1,2}[:：]\d{2})\s*[~〜\-＝ー]\s*(\d{1,2}[:：]\d{2})", txt)
    if m:
        s = parse_time_str(m.group(1).replace("：", ":"), None)
        e = parse_time_str(m.group(2).replace("：", ":"), None)
        if s and e and s < e:
            return s, e

    # 2) あいまい語（公休/午後/終日） + 睾場離脱・休暇系キーワード
    # 「午後休みです」等は「午後」を含むは」旧ロジックは「抜け」等のみ見でより中抜き推定に到靔しなかった
    leave_keywords = (
        "事務所", "会議", "教育", "研修", "外出", "離れ", "抜け", "中抜き", "打坈せ",
        "休み", "休暇", "欠勤",
    )
    has_leave_hint = any(k in txt for k in leave_keywords)
    if not has_leave_hint:
        return None, None

    if ("終日" in txt) or ("1日" in txt and "通常" not in txt):
        return start_t, end_t
    if ("公休中" in txt) or ("公休" in txt):
        return start_t, noon_end
    if ("午後" in txt):
        return afternoon_start, end_t

    return None, None
_AFTERNOON_OFF_DISPLAY_END = DEFAULT_BREAKS[0][0]
def _reason_is_afternoon_off(reason: str) -> bool:
    """後休（午後年休・公休のみ勤務）または備考の午後休系。"""
    r = str(reason or "")
    return ("午後" in r and ("休" in r or "休み" in r)) or ("後休" in r)
def _reason_is_morning_off(reason: str) -> bool:
    """公休（公休年休・午後のみ勤務）。カレンダー由来の略坷のみ明示扱い（事務所勤務などと混坌しない）。"""
    return "公休" in str(reason or "")
def _calendar_display_clock_out_for_calendar_sheet(entry: dict, day_date: date):
    """
    配台は breaks_dt の午後中抜きで正ししなる一方」end_dt は 17:00 のままてと結果カレンダーの退勤列の値誤る。
    後休（午後年休）または備考は午後休み系で」定時まで続し午後の中抜きはあるとしの値退勤表示を休憩時間1_開始に权ごる（end_dt 本体は変更しない）。
    """
    if not entry.get("is_working"):
        return None
    end_dt = entry.get("end_dt")
    if end_dt is None:
        return None
    reason = str(entry.get("reason") or "")
    afternoon_off = _reason_is_afternoon_off(reason)
    if not afternoon_off:
        return None
    breaks_dt = entry.get("breaks_dt") or []
    for b_s, b_e in breaks_dt:
        if b_s is None or b_e is None:
            continue
        bs = b_s.time() if isinstance(b_s, datetime) else b_s
        if isinstance(bs, datetime):
            bs = bs.time()
        if bs < DEFAULT_BREAKS[0][0]:
            continue
        if isinstance(b_e, datetime):
            be_cmp = b_e
        elif isinstance(b_e, time):
            be_cmp = datetime.combine(day_date, b_e)
        else:
            continue
        if be_cmp >= end_dt - timedelta(seconds=1):
            return datetime.combine(day_date, _AFTERNOON_OFF_DISPLAY_END)
    return None
def _member_schedule_break_cell_label(grid_mid_dt, breaks_dt, shift_end_dt, reason):
    """
    個人_* スケジュールの10分枠は休憩帯に入るとしの文言。
    昼食など通常休憩は「休憩」。後休（午後年休）で定時まで工場にいない午後帯は「休暇」。
    公休（公休年休）で公休の欠勤区間は休憩帯として入っている場合は「休暇」。
    """
    reason = str(reason or "")
    afternoon_off = _reason_is_afternoon_off(reason)
    morning_off = _reason_is_morning_off(reason)
    for b_s, b_e in breaks_dt:
        if b_s is None or b_e is None:
            continue
        if not (b_s <= grid_mid_dt < b_e):
            continue
        if isinstance(b_e, datetime) and isinstance(shift_end_dt, datetime):
            bs = b_s.time() if isinstance(b_s, datetime) else b_s
            if isinstance(bs, datetime):
                bs = bs.time()
            if afternoon_off and bs >= DEFAULT_BREAKS[0][0] and b_e >= shift_end_dt - timedelta(seconds=2):
                return "休暇"
            if morning_off and bs < DEFAULT_BREAKS[0][0]:
                be_t = b_e.time() if isinstance(b_e, datetime) else b_e
                if be_t <= time(13, 0):
                    return "休暇"
        return "休憩"
    return None
def _member_schedule_off_shift_label(
    day_date: date,
    grid_mid_dt: datetime,
    d_start_dt: datetime,
    d_end_dt: datetime,
    reason: str,
) -> str:
    """
    個人_* シートで所定出退勤の外側の10分枠。
    公休の公休（工場日の所定開始～午後出勤まで）は年休」後休の午後は年休。しれ以外のシフト外は勤務外。
    """
    r = str(reason or "")
    day_start = datetime.combine(day_date, DEFAULT_START_TIME)
    day_end = datetime.combine(day_date, DEFAULT_END_TIME)
    if grid_mid_dt < d_start_dt:
        if _reason_is_morning_off(r) and grid_mid_dt >= day_start:
            return "年休"
        return "勤務外"
    if grid_mid_dt >= d_end_dt:
        if _reason_is_afternoon_off(r) and grid_mid_dt < day_end:
            return "年休"
        return "勤務外"
    return "勤務外"
def _member_schedule_full_day_off_label(entry) -> str:
    """
    全日非勤務（is_working=False）の個人シート列の表示。
    休暇区分は年休（カレンダー *）のときは『年休」。工場休日などは『休」。
    """
    if not entry:
        return "休"
    r = str(entry.get("reason") or "").strip()
    if r == "年休" or r.startswith("年休 "):
        return "年休"
    return "休"
def _attendance_remark_text(row) -> str:
    """
    勤怠1行から「備考」列のテキストのみ取得れる。
    勤怠AIの解析リストへの投入はこの列のみ。reason 文字列は load_attendance で備考と休暇区分を読み取れる。
    """
    if row is None:
        return ""
    try:
        v = row.get(ATT_COL_REMARK)
    except Exception:
        return ""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    return s
def _attendance_leave_type_text(row) -> str:
    """勤怠1行から「休暇区分」列（カレンダー由来の 公休/後休 等）。"""
    if row is None:
        return ""
    try:
        v = row.get(ATT_COL_LEAVE_TYPE)
    except Exception:
        return ""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    return s
def _attendance_leave_type_is_full_day_paid_leave(leave_type: str) -> bool:
    """休暇区分がマスタ上の『終日年休』とみなせるとき True（前休・後休は午前/午後のみ勤務のため除外）。"""
    lt = unicodedata.normalize("NFKC", str(leave_type or "").strip())
    return lt == "年休" or lt.startswith("年休 ")
def _attendance_leave_type_is_absent(leave_type: str) -> bool:
    """休暇区分が欠勤（終日非勤務・配台不参加）のとき True。"""
    lt = unicodedata.normalize("NFKC", str(leave_type or "").strip())
    return lt == "欠勤" or lt.startswith("欠勤 ")
def _attendance_leave_type_is_calendar_no_dispatch(leave_type) -> bool:
    """
    master.xlsm カレンダー由来の休暇区分「-」（半角。NFKC で全角マイナス等も「-」に寄せる）。
    休日ではないが加工ラインへの配台（OP/AS）には載せない日。勤怠 AI や API 未設定でも確定させる。
    """
    lt = unicodedata.normalize("NFKC", str(leave_type or "").strip())
    return lt == "-"
def _attendance_leave_type_is_holiday_work(leave_type: str) -> bool:
    """休暇区分が休日出勤（配台対象の稼働日）のとき True。"""
    lt = unicodedata.normalize("NFKC", str(leave_type or "").strip())
    return lt == "休日出勤" or lt.startswith("休日出勤 ") or lt in ("午前休出", "午後休出")
def _attendance_leave_type_is_calendar_public_off(leave_type: str) -> bool:
    """休暇区分が会社カレンダー公休・所定休（終日非勤務）のとき True。"""
    lt = unicodedata.normalize("NFKC", str(leave_type or "").strip())
    return lt in ("公休", "休") or lt.startswith("公休 ")
def _attendance_preset_remark_markers() -> frozenset[str]:
    """プリセット由来の既定備考（自由記述なし＝API 不要）。"""
    return frozenset(
        {"公休", "休", "年休", "欠勤", "-", "前休", "後休", "通常", "休日出勤", "午前休出", "午後休出"}
    )
def _attendance_skip_remark_ai(remark: str, leave_type: str) -> bool:
    """コードで休暇・配台可否を確定できる行は勤怠備考 AI に載せない（トークン節約）。"""
    rem = unicodedata.normalize("NFKC", str(remark or "").strip())
    lt = unicodedata.normalize("NFKC", str(leave_type or "").strip())
    if _attendance_leave_type_is_calendar_no_dispatch(lt):
        return True
    if _attendance_leave_type_is_holiday_work(lt):
        return True
    if _attendance_leave_type_is_full_day_paid_leave(lt):
        return True
    if _attendance_leave_type_is_absent(lt):
        return True
    if _attendance_leave_type_is_calendar_public_off(lt):
        return True
    markers = _attendance_preset_remark_markers()
    if rem in markers and (not lt or lt in markers or lt == rem):
        return True
    if rem and lt and rem == lt and lt in markers:
        return True
    return False
def _ai_json_bool(v, default: bool = False) -> bool:
    """勤怠備考 AI の真偽値（bool / 数値 / 文字列の杺れを坸坎）。"""
    if v is None:
        return default
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v != 0
    if isinstance(v, float):
        if pd.isna(v):
            return default
        return v != 0.0
    s = str(v).strip().lower()
    if s in ("true", "1", "yes", "y", "はい", "真", "on"):
        return True
    if s in ("false", "0", "no", "n", "いいえ", "坽", "off", ""):
        return False
    return default
ATTENDANCE_AI_ENTRY_KEY_FIELD = "対象キー"
ATTENDANCE_REMARK_AI_RESPONSE_SCHEMA: dict = {
    "type": "OBJECT",
    "properties": {
        "entries": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    ATTENDANCE_AI_ENTRY_KEY_FIELD: {"type": "STRING"},
                    "出勤時刻": {"type": "STRING"},
                    "退勤時刻": {"type": "STRING"},
                    "中抜き開始": {"type": "STRING"},
                    "中抜き終了": {"type": "STRING"},
                    "作業効率": {"type": "NUMBER"},
                    "is_holiday": {"type": "BOOLEAN"},
                    "配台不参加": {"type": "BOOLEAN"},
                },
                "required": [ATTENDANCE_AI_ENTRY_KEY_FIELD],
            },
        }
    },
    "required": ["entries"],
}
def _attendance_remark_ai_prompt(remark_lines) -> str:
    """勤怠備考 AI のプロンプト（通常勤務との差分だけを返させる）。

    全行の全項目を書き戻させると 1 リクエストの出力が数万トークンに膨れ、応答が数分かかる。
    """
    joined = "\n".join(str(x) for x in remark_lines)
    return f"""
以下は勤怠の「備考」「休暇区分」の自由記述です。出退勤時刻の変更・中抜き・休日・配台不参加・作業効率のうち、
通常勤務から**変わる点があるものだけ**を JSON で返してください。

【出力量の制約（最重要）】
- 通常どおりで変更がない行は entries に**含めない**。
- 変更がある行でも、**読み取れる項目だけ**を書く。変更のない項目・不明な項目は**省略**する（null を書かない）。
- 説明文・マークダウン（``` 等）は出力しない。JSON のみ。

【出力形式】
{{
  "entries": [
    {{"{ATTENDANCE_AI_ENTRY_KEY_FIELD}": "YYYY-MM-DD_メンバー名", "出勤時刻": "HH:MM"}}
  ]
}}

【項目】
- {ATTENDANCE_AI_ENTRY_KEY_FIELD}（必須）: 下の一覧の「YYYY-MM-DD_メンバー名」をそのまま写す。
- 出勤時刻 / 退勤時刻: "HH:MM"。定時どおりなら省略。
- 中抜き開始 / 中抜き終了: 一時的な離脱（中抜け・事務所・会議など）があるときだけ両方を書く。
  例: 「午前中は事務所で作業」→ 中抜き開始 "08:45"・中抜き終了 "12:00"
      「午後は会議」→ 中抜き開始 "13:00"・中抜き終了 "17:00"
- is_holiday: 終日休暇・欠勤など**勤務自体がない**ときだけ true。
  午前休・午後休など部分的な休みでは書かず、出退勤時刻や中抜きで表現する。
- 配台不参加: 勤務はあるが**加工ラインへの配台（OP/AS の割当）に載せてはいけない**ときだけ true。
  表記は問わず意味で判断する（「配台不可」「配台ＮＧ」「ラインに乗らない」「月次点検のみ」「点検で一日」
  「事務のみ」「教育で現場不可」「手配なし」「アサイン不要」などの言い換えも含む）。
  休暇区分が「-」（ハイフン1文字）のみのときは is_holiday を書かず、配台不参加 を true にする。
- 作業効率: 0.0〜1.0。通常（1.0）なら省略。

【特記事項リスト】
{joined}
"""
def _attendance_ai_entries_to_map(payload) -> dict:
    """勤怠備考 AI の応答を ``{"YYYY-MM-DD_メンバー": {...}}`` に正規化する。

    新形式は ``entries`` 配列。旧形式（トップレベルが対象キーの辞書）もキャッシュ互換で受ける。
    """
    if not isinstance(payload, dict):
        return {}
    entries = payload.get("entries")
    if entries is None:
        return {
            k: v
            for k, v in payload.items()
            if isinstance(k, str) and k.strip() and isinstance(v, dict)
        }
    if not isinstance(entries, list):
        return {}
    out: dict = {}
    for ent in entries:
        if not isinstance(ent, dict):
            continue
        key = str(ent.get(ATTENDANCE_AI_ENTRY_KEY_FIELD) or "").strip()
        if not key:
            continue
        out[key] = {
            k: v
            for k, v in ent.items()
            if k != ATTENDANCE_AI_ENTRY_KEY_FIELD and v is not None
        }
    return out
def _attendance_is_empty_shift(row, *, key: str, analyzed_keys) -> bool:
    """出退勤が両方空で、かつ AI 解析にも回らなかった行（備考も休暇区分も無い）か。

    AI は変更のない行を返さないため、応答の有無ではなく解析対象だったかどうかで判定する。
    """
    if not (pd.isna(row.get("出勤時間")) and pd.isna(row.get("退勤時間"))):
        return False
    return key not in (analyzed_keys or ())
def _parse_attendance_overtime_end_optional(v) -> time | None:
    """勤怠「残業(分)」列。有効な時刻のみ。空・不正は None（_excel_scalar_to_time_optional と同趣旨）。"""
    return _excel_scalar_to_time_optional(v)
def _resolve_attendance_overtime_end(
    raw,
    *,
    base_end_t: time,
    curr_date: date,
) -> time | None:
    """
    勤怠「残業(分)」列の解釈（いずれかで成功したらその time を返す）。

    1) 時刻（文字列 HH:MM、datetime、time、Excel 0<値<1 の日内小数）
    2) 定時退勤からの延長「分」: 1〜720 の整数（Excel 数値・文字列の整数も可）
    """
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    if isinstance(raw, bool):
        return None
    t_clock = _parse_attendance_overtime_end_optional(raw)
    if t_clock is not None:
        return t_clock
    if isinstance(raw, str):
        s = raw.strip()
        if s.isdigit():
            try:
                raw = int(s)
            except ValueError:
                return None
    if isinstance(raw, (int, float)):
        x = float(raw)
        if 0 < x < 1:
            try:
                new_dt = datetime.combine(curr_date, time(0, 0)) + timedelta(days=x)
                return new_dt.time()
            except (OverflowError, ValueError):
                return None
        if x == int(x) and 1 <= int(x) <= 720:
            try:
                base_dt = datetime.combine(curr_date, base_end_t)
                new_dt = base_dt + timedelta(minutes=int(x))
                if new_dt.date() != curr_date:
                    return time(23, 59, 59)
                return new_dt.time()
            except (OverflowError, ValueError):
                return None
    return None
def load_attendance_and_analyze(members):
    attendance_data = {}
    # ※「勤怠備考」は master 坄メンバーシートの「備考」列のみ。メイン再優先・特別指定_備考は別API（generate_plan 坴で追記）。
    ai_log = {
        "（注）このシートの見方": "勤怠は「勤怠備考_*」と「勤怠備考_Geminiモデル」。メイン再優先・特別指定は JSON と「_*_AI_API」「_*_Geminiモデル」行で確認。",
        "勤怠備考_AI_API": "なし",
        "勤怠備考_AI_詳細": "解析対象の備考行なし",
        "勤怠備考_Geminiモデル": "—（解析対象の備考行なし）",
    }
    
    from planning_core.core.attendance_paths import attendance_data_json_path
    from planning_core.core.attendance_store import (
        apply_company_calendar_to_members,
        load_attendance_store,
        member_attendance_to_dataframe_records,
    )

    jp = attendance_data_json_path()
    if not jp.is_file():
        raise RuntimeError(
            "勤怠データを読み込めません。attendance-data.json が未作成です。"
            "会社カレンダー／メンバー勤怠タブでセットアップしてください。"
            " master.xlsm のレガシー勤怠シートへはフォールバックしません。"
        )

    try:
        store = load_attendance_store(jp)
        from planning_core.core.attendance_member_roster import members_for_attendance_analysis

        analysis_members = members_for_attendance_analysis(list(members), store)
        json_records = member_attendance_to_dataframe_records(store, analysis_members)
        if not json_records and analysis_members:
            y = int(store.get("company_calendar", {}).get("year") or date.today().year)
            for month in range(1, 13):
                apply_company_calendar_to_members(
                    store, list(analysis_members), y, month, only_unedited=False
                )
            json_records = member_attendance_to_dataframe_records(store, analysis_members)
        if json_records:
            df = pd.DataFrame(json_records)
            df["日付"] = pd.to_datetime(df["日付"], errors="coerce").dt.date
            df = df.dropna(subset=["日付"])
            logging.info(
                "勤怠正本 attendance-data.json を読み込みました（%s、%d 行）。",
                jp,
                len(df),
            )
        else:
            logging.info(
                "勤怠正本 attendance-data.json は存在しますがメンバー勤怠行がありません（%s）。"
                "レガシーシートへはフォールバックしません。",
                jp,
            )
            df = pd.DataFrame()
    except Exception as e:
        raise RuntimeError(
            f"勤怠正本 {jp} の読込に失敗しました。アプリで内容を確認・修復してください: {e}"
        ) from e

    # 2. AI による勤怠文脈の解析
    remarks_to_analyze = []
    analyzed_keys: set[str] = set()
    for _, row in df.iterrows():
        m = str(row.get('メンバー', '')).strip()
        if m not in members:
            continue
        rem = _attendance_remark_text(row)
        lt = _attendance_leave_type_text(row)
        d_str = row['日付'].strftime("%Y-%m-%d") if pd.notna(row['日付']) else ""
        key = f"{d_str}_{m}"
        if _attendance_skip_remark_ai(rem, lt):
            continue
        if rem:
            remarks_to_analyze.append(f"{key} の備考: {rem}")
            analyzed_keys.add(key)
        elif lt and lt not in ("通常", ""):
            remarks_to_analyze.append(f"{key} の休暇区分（備考は空）: {lt}")
            analyzed_keys.add(key)

    if remarks_to_analyze:
        remarks_blob = "\n".join(remarks_to_analyze)
        cache_key = hashlib.sha256(
            (remarks_blob + "\n" + ATTENDANCE_REMARK_AI_SCHEMA_ID).encode("utf-8")
        ).hexdigest()
        ai_cache = load_ai_cache()

        # 同一備考セットはキャッシュを優先利用し、APIコールを節約
        cached_data = get_cached_ai_result(ai_cache, cache_key)
        if cached_data is not None:
            ai_parsed = _attendance_ai_entries_to_map(cached_data)
            ai_log["勤怠備考_AI_API"] = "なし(キャッシュ使用)"
            ai_log["勤怠備考_AI_詳細"] = "キャッシュヒット"
            ai_log["勤怠備考_Geminiモデル"] = "—（キャッシュ利用・今回 API 未実行）"
        elif not API_KEY:
            ai_parsed = {}
            ai_log["勤怠備考_AI_API"] = "なし"
            ai_log["勤怠備考_AI_詳細"] = "Gemini API キー未設定のため勤怠備考AIをスキップ"
            ai_log["勤怠備考_Geminiモデル"] = "—（API キー未設定）"
            logging.info("Gemini API キーが未設定のため備考AI解析をスキップしました。")
        else:
            logging.info(
                "■ AIが複数日の特記事項を解析中...（対象 %d 件）",
                len(remarks_to_analyze),
            )
            ai_log["勤怠備考_AI_API"] = "あり"
            try:
                client = _gemini_client(API_KEY)
                ai_parsed, models_used, failed_batches = (
                    _gemini_generate_json_map_in_batches(
                        client,
                        items=remarks_to_analyze,
                        build_prompt=_attendance_remark_ai_prompt,
                        log_label="勤怠備考AI",
                        response_schema=ATTENDANCE_REMARK_AI_RESPONSE_SCHEMA,
                        parse_map=_attendance_ai_entries_to_map,
                    )
                )
                ai_log["勤怠備考_Geminiモデル"] = (
                    ", ".join(dict.fromkeys(models_used)) or "—（呼び出し失敗）"
                )
                if failed_batches:
                    # 欠けたままキャッシュすると次回以降も欠けたままになるので保存しない
                    ai_log["勤怠備考_AI_詳細"] = f"一部失敗（{failed_batches} バッチ）"
                else:
                    put_cached_ai_result(ai_cache, cache_key, ai_parsed)
                    save_ai_cache(ai_cache)
                    ai_log["勤怠備考_AI_詳細"] = "解析成功"
            except Exception as e:
                ai_parsed = {}
                logging.warning("AI通信エラー: %s", e)
                ai_log["勤怠備考_AI_詳細"] = str(e)
                ai_log["勤怠備考_Geminiモデル"] = "—（呼び出し失敗）"
    else:
        ai_parsed = {}

    # 3. 日付ととの制約辞書を構築
    for _, row in df.iterrows():
        if pd.isna(row['日付']): continue
        curr_date = row['日付']
        m = str(row.get('メンバー', '')).strip()
        if m not in members: continue

        if curr_date not in attendance_data:
            attendance_data[curr_date] = {}

        original_reason = _attendance_remark_text(row)
        leave_type = _attendance_leave_type_text(row)

        key = f"{curr_date.strftime('%Y-%m-%d')}_{m}"
        ai_info = ai_parsed.get(key, {})

        is_empty_shift = _attendance_is_empty_shift(
            row, key=key, analyzed_keys=analyzed_keys
        )
        is_holiday = _ai_json_bool(ai_info.get("is_holiday"), False) or is_empty_shift
        forced_calendar_paid_leave = _attendance_leave_type_is_full_day_paid_leave(leave_type)
        if forced_calendar_paid_leave:
            is_holiday = True
        exclude_from_line = _ai_json_bool(ai_info.get("配台不参加"), False)
        if _attendance_leave_type_is_absent(leave_type):
            is_holiday = True
            exclude_from_line = True
        if _attendance_leave_type_is_calendar_public_off(leave_type):
            is_holiday = True
        if _attendance_leave_type_is_calendar_no_dispatch(leave_type):
            exclude_from_line = True
            # 休日ではないが加工配台のみ除外（AI・空シフト推定で is_holiday になるのを防ぐ）
            is_holiday = False
        if _attendance_leave_type_is_holiday_work(leave_type):
            is_holiday = False

        ai_eff = ai_info.get("作業効率")
        excel_eff = row.get('作業効率')
        
        if ai_eff is not None:
            eff_val = ai_eff
        elif excel_eff is not None and not pd.isna(excel_eff):
            eff_val = excel_eff
        else:
            eff_val = 1.0
            
        try:
            efficiency = float(eff_val)
        except (ValueError, TypeError):
            efficiency = 1.0

        if original_reason:
            if (
                leave_type
                and leave_type not in ("通常", "")
                and leave_type not in original_reason
            ):
                reason = f"{leave_type} {original_reason}"
            else:
                reason = original_reason
        elif leave_type and leave_type not in ("通常", ""):
            reason = leave_type
        else:
            reason = '通常' if not is_empty_shift else '休日シフト'

        # マスタに出勤・退勤の両方は入っている日は」勤怠AIの出勤/退勤時刻で上書きしない（休暇区分のみの行で誤推定されごる）
        excel_s = row.get("出勤時間")
        excel_e = row.get("退勤時間")
        if not pd.isna(excel_s) and not pd.isna(excel_e):
            start_t = parse_time_str(excel_s, DEFAULT_START_TIME)
            end_t = parse_time_str(excel_e, DEFAULT_END_TIME)
        else:
            start_t = parse_time_str(ai_info.get("出勤時刻") or excel_s, DEFAULT_START_TIME)
            end_t = parse_time_str(ai_info.get("退勤時刻") or excel_e, DEFAULT_END_TIME)
        base_end_t = end_t

        b1_s = parse_time_str(row.get('休憩時間1_開始'), DEFAULT_BREAKS[0][0])
        b1_e = parse_time_str(row.get('休憩時間1_終了'), DEFAULT_BREAKS[0][1])
        b2_s = parse_time_str(row.get('休憩時間2_開始'), DEFAULT_BREAKS[1][0])
        b2_e = parse_time_str(row.get('休憩時間2_終了'), DEFAULT_BREAKS[1][1])

        # ★追加: AIから中抜き時間を取得
        mid_break_s = parse_time_str(ai_info.get("中抜き開始"), None)
        mid_break_e = parse_time_str(ai_info.get("中抜き終了"), None)
        # AIは中抜きを返さなかった場合は「備考文言からローカル推定で補完
        if not (mid_break_s and mid_break_e):
            fb_s, fb_e = infer_mid_break_from_reason(reason, start_t, end_t, b1_s, b1_e)
            if fb_s and fb_e:
                mid_break_s, mid_break_e = fb_s, fb_e

        ot_applied_flag = False
        ot_end: time | None = None
        if not is_holiday:
            ot_end = _resolve_attendance_overtime_end(
                row.get(ATT_COL_OT_END),
                base_end_t=base_end_t,
                curr_date=curr_date,
            )
            if ot_end is not None:
                end_t = ot_end
                ot_applied_flag = True

        def combine_dt(t): return datetime.combine(curr_date, t) if t else None
        
        start_dt = combine_dt(start_t)
        end_dt = combine_dt(end_t)
        if (not is_holiday) and start_dt and end_dt and end_dt <= start_dt:
            logging.warning(
                "勤怠 %s %s: %s 適用後に退勤が出勤以前となったため、%s を無視して定時退勤に戻します。",
                curr_date,
                m,
                ATT_COL_OT_END,
                ATT_COL_OT_END,
            )
            end_t = base_end_t
            end_dt = combine_dt(end_t)
        breaks_dt = []
        
        # 通常の休憩を追加
        if b1_s and b1_e: breaks_dt.append((combine_dt(b1_s), combine_dt(b1_e)))
        if b2_s and b2_e: breaks_dt.append((combine_dt(b2_s), combine_dt(b2_e)))
        
        # ★追加: 中抜き時間はある場合は」特別な「休憩」としてスケジュール計算に追加
        if mid_break_s and mid_break_e: breaks_dt.append((combine_dt(mid_break_s), combine_dt(mid_break_e)))
        
        is_working = not is_holiday
        ot_minutes = 0
        if not is_holiday:
            ot_minutes = _attendance_overtime_minutes_from_raw(
                row.get(ATT_COL_OT_END),
                base_end_t=base_end_t,
                curr_date=curr_date,
            )
        base_end_dt = combine_dt(base_end_t)
        attendance_data[curr_date][m] = {
            "is_working": is_working,
            "eligible_for_assignment": is_working and (not exclude_from_line),
            "start_dt": start_dt,
            "end_dt": end_dt,
            "base_end_dt": base_end_dt,
            "breaks_dt": merge_time_intervals(breaks_dt),
            "efficiency": efficiency,
            "reason": reason,
            "overtime_minutes": ot_minutes,
        }

    return attendance_data, ai_log
def _attendance_overtime_minutes_from_raw(
    raw,
    *,
    base_end_t: time,
    curr_date: date,
) -> int:
    """master「残業(分)」セルを延長分（1〜720）に正規化。空・時刻のみは分換算、無効は 0。"""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return 0
    if isinstance(raw, bool):
        return 0
    if isinstance(raw, str):
        s = raw.strip()
        if not s:
            return 0
        if s.isdigit():
            try:
                raw = int(s)
            except ValueError:
                return 0
    if isinstance(raw, (int, float)):
        x = float(raw)
        if x == int(x) and 1 <= int(x) <= 720:
            return int(x)
    t_clock = _parse_attendance_overtime_end_optional(raw)
    if t_clock is not None and base_end_t is not None:
        try:
            base_dt = datetime.combine(curr_date, base_end_t)
            end_dt = datetime.combine(curr_date, t_clock)
            if end_dt > base_dt:
                return min(720, int((end_dt - base_dt).total_seconds() // 60))
        except (OverflowError, ValueError):
            return 0
    return 0
def _overtime_simulation_json_path() -> "Path | None":
    from pathlib import Path

    raw = (os.environ.get(ENV_OVERTIME_SIMULATION_JSON) or "").strip()
    if not raw:
        return None
    p = Path(raw)
    return p if p.is_file() else None
def _pick_member_attendance_template(
    attendance_data: dict, member: str, target_date: date
) -> tuple[date | None, dict | None]:
    """当該メンバーの直近稼働日をテンプレとして返す。"""
    plan_dates = sorted(attendance_data.keys())
    for d in reversed([d for d in plan_dates if d <= target_date]):
        st = attendance_data.get(d, {}).get(member)
        if st and st.get("is_working"):
            return d, st
    for d in plan_dates:
        st = attendance_data.get(d, {}).get(member)
        if st and st.get("is_working"):
            return d, st
    return None, None
def _default_attendance_entry_for_date(d: date) -> dict:
    start_dt = datetime.combine(d, DEFAULT_START_TIME)
    end_dt = datetime.combine(d, DEFAULT_END_TIME)
    breaks_dt = [
        (datetime.combine(d, bs), datetime.combine(d, be)) for bs, be in DEFAULT_BREAKS
    ]
    return {
        "is_working": True,
        "eligible_for_assignment": True,
        "start_dt": start_dt,
        "end_dt": end_dt,
        "base_end_dt": end_dt,
        "breaks_dt": merge_time_intervals(breaks_dt),
        "efficiency": 1.0,
        "reason": "残業シミュレーション（休日出勤）",
        "overtime_minutes": 0,
    }
def build_attendance_overtime_preview_dict() -> dict:
    """段階3.5 ウィザード向け: load_attendance_and_analyze と同一ロジックの勤怠プレビュー。"""
    (
        _skills_dict,
        members,
        _equipment_list,
        _req_map,
        _need_rules,
        _surplus_map,
        _need_combo_col_index,
    ) = load_skills_and_needs()
    if not members:
        return {
            "format_version": 1,
            "ok": False,
            "error": "skills にメンバーが登録されていません",
            "members": [],
            "dates": [],
            "cells": {},
        }
    attendance_data, _ai_log = load_attendance_and_analyze(members)
    today = date.today()
    window_end = today + timedelta(days=30)
    sorted_dates = sorted(
        d for d in attendance_data.keys() if today <= d <= window_end
    )
    cells: dict = {}
    for d in sorted_dates:
        d_key = d.isoformat()
        cells[d_key] = {}
        weekend = d.weekday() >= 5
        for m in members:
            st = attendance_data.get(d, {}).get(m)
            if not st:
                cells[d_key][m] = {
                    "is_working": False,
                    "eligible_for_assignment": False,
                    "overtime_minutes": 0,
                    "weekend": weekend,
                }
                continue
            cells[d_key][m] = {
                "is_working": bool(st.get("is_working")),
                "eligible_for_assignment": bool(
                    st.get("eligible_for_assignment", st.get("is_working"))
                ),
                "overtime_minutes": int(st.get("overtime_minutes") or 0),
                "weekend": weekend,
            }
    return {
        "format_version": 1,
        "ok": True,
        "members": list(members),
        "dates": [d.isoformat() for d in sorted_dates],
        "cells": cells,
    }
def apply_overtime_simulation_overrides(
    attendance_data: dict, path: "Path | None" = None
) -> bool:
    """
    段階3.5: PM_AI_OVERTIME_SIMULATION_JSON の working_overrides / overtime_minutes を
    attendance_data にインプレース適用する（master は変更しない）。
    """
    from pathlib import Path

    p = path or _overtime_simulation_json_path()
    if p is None:
        return False
    try:
        payload = json.loads(Path(p).read_text(encoding="utf-8"))
    except Exception as e:
        logging.warning("残業シミュレーション JSON 読込失敗: %s", e)
        return False
    if not isinstance(payload, dict):
        return False

    working_overrides = payload.get("working_overrides") or {}
    overtime_map = payload.get("overtime_minutes") or {}
    applied = False

    if isinstance(working_overrides, dict):
        for d_str, mem_map in working_overrides.items():
            if not isinstance(mem_map, dict):
                continue
            d = parse_optional_date(d_str)
            if d is None:
                continue
            if d not in attendance_data:
                attendance_data[d] = {}
            for member, flag in mem_map.items():
                m = str(member).strip()
                if not m:
                    continue
                if flag is True:
                    tmpl_d, tmpl_st = _pick_member_attendance_template(
                        attendance_data, m, d
                    )
                    if tmpl_st and tmpl_d is not None:
                        cloned = _clone_attendance_day_shifted(
                            {m: tmpl_st}, tmpl_d, d
                        )[m]
                    else:
                        cloned = _default_attendance_entry_for_date(d)
                    cloned["is_working"] = True
                    cloned["eligible_for_assignment"] = True
                    cloned["reason"] = "残業シミュレーション（休日出勤）"
                    attendance_data[d][m] = cloned
                    applied = True
                elif flag is False:
                    ent = attendance_data[d].get(m)
                    if ent is None:
                        attendance_data[d][m] = {
                            "is_working": False,
                            "eligible_for_assignment": False,
                            "start_dt": None,
                            "end_dt": None,
                            "base_end_dt": None,
                            "breaks_dt": [],
                            "efficiency": 1.0,
                            "reason": "残業シミュレーション（休日扱い）",
                            "overtime_minutes": 0,
                        }
                    else:
                        ent = dict(ent)
                        ent["is_working"] = False
                        ent["eligible_for_assignment"] = False
                        ent["overtime_minutes"] = 0
                        attendance_data[d][m] = ent
                    applied = True

    if isinstance(overtime_map, dict):
        for d_str, mem_map in overtime_map.items():
            if not isinstance(mem_map, dict):
                continue
            d = parse_optional_date(d_str)
            if d is None:
                continue
            if d not in attendance_data:
                continue
            for member, raw_min in mem_map.items():
                m = str(member).strip()
                if not m:
                    continue
                try:
                    ot_min = int(raw_min)
                except (TypeError, ValueError):
                    continue
                if ot_min < 0 or ot_min > 720:
                    continue
                ent = attendance_data[d].get(m)
                if not ent or not ent.get("is_working"):
                    continue
                ent = dict(ent)
                base_end_dt = ent.get("base_end_dt") or ent.get("end_dt")
                if base_end_dt is None:
                    base_end_dt = datetime.combine(d, DEFAULT_END_TIME)
                base_end_t = base_end_dt.time()
                if ot_min <= 0:
                    ent["end_dt"] = base_end_dt
                    ent["overtime_minutes"] = 0
                else:
                    new_end_t = _resolve_attendance_overtime_end(
                        ot_min,
                        base_end_t=base_end_t,
                        curr_date=d,
                    )
                    if new_end_t is not None:
                        new_end_dt = datetime.combine(d, new_end_t)
                        start_dt = ent.get("start_dt")
                        if start_dt and new_end_dt <= start_dt:
                            ent["end_dt"] = base_end_dt
                            ent["overtime_minutes"] = 0
                        else:
                            ent["end_dt"] = new_end_dt
                            ent["overtime_minutes"] = ot_min
                    else:
                        ent["end_dt"] = base_end_dt
                        ent["overtime_minutes"] = 0
                attendance_data[d][m] = ent
                applied = True

    if applied:
        logging.info(
            "残業シミュレーション: %s を attendance_data に適用しました。",
            p,
        )
    return applied

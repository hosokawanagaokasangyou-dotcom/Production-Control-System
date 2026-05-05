# -*- coding: utf-8 -*-
"""
インタラクティブ配台試行: 結果_配台表.json を tasks_df にマージし、段階2と同じ _generate_plan_impl を実行。

- 段階2の配台ループは変更せず、試行専用の環境・カレンダー解釈・結果表上書きで差し替える。
- 配台試行順は入力 JSON を正とする。結果_配台表は timeline の暦日集約を基準とし、入力が依頼×機械あたり 1 行のときも潰さない（planning_core）。
- 機械カレンダーは * / ＊ / ※ のセルのみ占有。工場枠は master A12/B12 開始・同日 23:59 まで延長可。
  加工が暦日をまたぐ場合は PlanningValidationError で中止。
- 人員不足は interactive_trial_shortages_snapshot の op_shortage / as_shortage に記録。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
import traceback
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
os.chdir(str(SCRIPT_DIR))

# region agent log
_AGENT_DEBUG_SESSION = "327eec"
# 手動修正タブ調査用（スクリーンショットの依頼NO）
_WATCH_TASK_IDS = frozenset(
    {"JR260501", "Y5-4", "Y5-5", "Y5-14", "Y5-6", "Y5-3", "Y5-8"}
)
_WATCH_FOCUS_DAY = "2026-05-08"


def _agent_debug_ndjson(
    hypothesis_id: str, location: str, message: str, data: dict
) -> None:
    """Cursor debug NDJSON（ワークスペース .cursor/debug-<session>.log）。"""
    try:
        root = Path(__file__).resolve().parent.parent.parent.parent
        log_path = root / ".cursor" / f"debug-{_AGENT_DEBUG_SESSION}.log"
        payload = {
            "sessionId": _AGENT_DEBUG_SESSION,
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data,
            "timestamp": int(time.time() * 1000),
        }
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except Exception:
        pass


def _parse_contract_iso_dt(obj) -> datetime | None:
    if isinstance(obj, dict) and obj.get("__t") == "datetime":
        v = str(obj.get("v") or "").strip()
        if v:
            try:
                return datetime.fromisoformat(v)
            except ValueError:
                return None
    return None


def _is_machining_kind(ev: dict) -> bool:
    ek = str(ev.get("event_kind") or "").strip()
    return ek in ("", "machining")


def _machine_hint_sec(ev: dict) -> bool:
    m = str(ev.get("machine") or "") + str(ev.get("machine_occupancy_key") or "")
    return "SEC" in m.upper()


def _summarize_equipment_gantt_contract_for_watch(
    contract_path: Path,
) -> dict:
    """設備ガント契約 JSON の timeline_events から、監視依頼NO×SEC×指定日の連続性を要約。"""
    out: dict = {
        "contract_path": str(contract_path),
        "exists": contract_path.is_file(),
    }
    if not contract_path.is_file():
        return out
    try:
        raw = json.loads(contract_path.read_text(encoding="utf-8"))
    except Exception as e:
        out["parse_error"] = str(e)
        return out
    packed = raw.get("kwargs_packed") if isinstance(raw, dict) else None
    evs = (packed or {}).get("timeline_events") if isinstance(packed, dict) else None
    if not isinstance(evs, list):
        out["timeline_len"] = 0
        return out
    target = date.fromisoformat(_WATCH_FOCUS_DAY)
    picked = []
    for ev in evs:
        if not isinstance(ev, dict):
            continue
        if not _is_machining_kind(ev):
            continue
        tid = str(ev.get("task_id") or "").strip()
        if tid not in _WATCH_TASK_IDS:
            continue
        if not _machine_hint_sec(ev):
            continue
        st = _parse_contract_iso_dt(ev.get("start_dt"))
        ed = _parse_contract_iso_dt(ev.get("end_dt"))
        if st is None or ed is None:
            continue
        if st.date() != target:
            continue
        occ = str(ev.get("machine_occupancy_key") or ev.get("machine") or "")
        picked.append((occ, tid, st, ed))

    out["matching_event_count"] = len(picked)
    by_task_raw: dict[str, dict] = {}
    for _occ, tid, st, ed in picked:
        acc = by_task_raw.setdefault(
            tid,
            {
                "segments": 0,
                "_first": None,
                "_last": None,
                "duration_min": 0.0,
            },
        )
        acc["segments"] += 1
        acc["duration_min"] += max(0.0, (ed - st).total_seconds() / 60.0)
        if acc["_first"] is None or st < acc["_first"]:
            acc["_first"] = st
        if acc["_last"] is None or ed > acc["_last"]:
            acc["_last"] = ed
    by_task: dict[str, dict] = {}
    for tid, acc in by_task_raw.items():
        fst = acc.pop("_first", None)
        lst = acc.pop("_last", None)
        if fst is not None:
            acc["first_start"] = fst.isoformat(timespec="seconds")
        if lst is not None:
            acc["last_end"] = lst.isoformat(timespec="seconds")
        by_task[tid] = acc

    by_occ: dict[str, list] = defaultdict(list)
    for occ, tid, st, ed in picked:
        by_occ[occ or "(empty)"].append((st, ed, tid))

    gap_info: dict[str, object] = {}
    for occ, lst in by_occ.items():
        lst.sort(key=lambda x: x[0])
        gaps_min = []
        for i in range(1, len(lst)):
            prev_end = lst[i - 1][1]
            cur_start = lst[i][0]
            g = (cur_start - prev_end).total_seconds() / 60.0
            if g > 1.0:
                gaps_min.append(round(g, 3))
        gap_info[occ] = {
            "ordered_segments": len(lst),
            "gaps_over_1min_minutes": gaps_min,
            "max_gap_minutes": max(gaps_min) if gaps_min else 0.0,
        }

    out["focus_day"] = _WATCH_FOCUS_DAY
    out["by_task_id"] = by_task
    out["gaps_by_occupancy_key"] = gap_info
    return out


def _filter_shortages_for_watch(snap: dict) -> dict:
    op = [r for r in (snap.get("op_shortage") or []) if isinstance(r, dict) and str(r.get("task_id") or "") in _WATCH_TASK_IDS]
    ast = [r for r in (snap.get("as_shortage") or []) if isinstance(r, dict) and str(r.get("task_id") or "") in _WATCH_TASK_IDS]
    return {
        "op_shortage_watch_count": len(op),
        "as_shortage_watch_count": len(ast),
        "op_shortage_watch": op[:40],
        "as_shortage_watch": ast[:40],
    }


def _day_iso(dd) -> str:
    if hasattr(dd, "isoformat"):
        try:
            return dd.isoformat()
        except Exception:
            pass
    return str(dd)


def _targets_watch_sec_day(targets: dict) -> dict:
    sums: dict[str, float] = defaultdict(float)
    for (tid, mach, dd), q in (targets or {}).items():
        if str(tid) not in _WATCH_TASK_IDS:
            continue
        if "SEC" not in str(mach).upper():
            continue
        if _day_iso(dd) != _WATCH_FOCUS_DAY:
            continue
        sums[str(tid)] += float(q)
    return dict(sums)


def _resolve_equipment_gantt_contract_path(plan_out: str) -> Path | None:
    p = Path(plan_out)
    if not p.name:
        return None
    sibling = p.parent / f"{p.stem}_equipment_gantt_contract.json"
    return sibling if sibling.is_file() else None


# endregion

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def main() -> int:
    if len(sys.argv) < 2:
        print(
            "usage: dispatch_interactive_trial.py <path-to-result-dispatch.json>",
            file=sys.stderr,
        )
        return 2
    path = Path(sys.argv[1]).resolve()
    if not path.is_file():
        print(f"not a file: {path}", file=sys.stderr)
        return 1
    print("[dispatch trial] 入力JSONを読み込み中…", flush=True)
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"json read failed: {e}", file=sys.stderr)
        return 1

    rows = raw.get("rows") if isinstance(raw, dict) else None
    if rows is None:
        print("missing rows array", file=sys.stderr)
        return 1
    json_columns = raw.get("columns") if isinstance(raw, dict) else None

    os.environ["PM_AI_INTERACTIVE_DISPATCH_TRIAL"] = "1"

    shortage_path = path.with_name("dispatch_trial_shortages.json")

    try:
        import planning_core as pc
        from planning_core.bootstrap import PlanningValidationError

        print("[dispatch trial] 計画タスクを読み込み、表データをマージ中…", flush=True)
        tasks_df = pc.load_planning_tasks_df()
        merged_df, targets = pc.merge_interactive_result_dispatch_json_into_tasks_df(
            tasks_df, rows
        )
        # region agent log
        _agent_debug_ndjson(
            "H1",
            "dispatch_interactive_trial.py:merge",
            "interactive_targets_watch_SEC_day",
            {
                "targets_watch_sec_day": _targets_watch_sec_day(targets),
                "total_target_keys": len(targets or {}),
                "watch_ids": sorted(_WATCH_TASK_IDS),
                "focus_day": _WATCH_FOCUS_DAY,
            },
        )
        # endregion
        print("[dispatch trial] 段階2（配台計画）を実行中…（時間がかかる場合があります）", flush=True)
        paths = pc._generate_plan_impl(
            tasks_df_override=merged_df,
            return_output_paths=True,
            interactive_relax_intraday=False,
            interactive_dispatch_targets=targets if targets else None,
            interactive_result_dispatch_json_rows=rows,
            interactive_result_dispatch_json_columns=json_columns
            if isinstance(json_columns, list)
            else None,
        )
        snap = pc.interactive_trial_shortages_snapshot()
        # region agent log
        pp = (paths or {}).get("production_plan") if isinstance(paths, dict) else None
        contract_path = _resolve_equipment_gantt_contract_path(str(pp or "")) if pp else None
        if pp and contract_path is None:
            _miss = Path(str(pp)).parent / (Path(str(pp)).stem + "_equipment_gantt_contract.json")
            gantt_summary = {
                "contract_path": str(_miss),
                "exists": False,
                "note": "sibling_contract_not_found",
            }
        elif contract_path is not None:
            gantt_summary = _summarize_equipment_gantt_contract_for_watch(contract_path)
        else:
            gantt_summary = {"contract_path": None, "exists": False, "note": "production_plan path missing"}
        _agent_debug_ndjson(
            "H2",
            "dispatch_interactive_trial.py:after_generate",
            "stage2_paths_and_gantt_contract",
            {
                "production_plan": str(pp or ""),
                "member_schedule": str((paths or {}).get("member_schedule") or "")
                if isinstance(paths, dict)
                else "",
                "contract_resolved": str(contract_path) if contract_path else "",
                "gantt_contract_summary": gantt_summary,
                "shortages_watch": _filter_shortages_for_watch(snap),
            },
        )
        # endregion
        shortage_payload: dict = {
            "format_version": 2,
            "source_json": str(path),
            "note": "interactive trial via planning_core._generate_plan_impl",
            "op_shortage": snap["op_shortage"],
            "as_shortage": snap["as_shortage"],
        }
        if isinstance(paths, dict):
            shortage_payload["production_plan"] = str(paths.get("production_plan") or "")
            shortage_payload["member_schedule"] = str(paths.get("member_schedule") or "")
        shortage_path.write_text(
            json.dumps(shortage_payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        print("[dispatch trial] 不足情報JSONを書き出しました。", flush=True)
    except PlanningValidationError as e:
        msg = str(e).strip() or "PlanningValidationError"
        print(msg, file=sys.stderr)
        try:
            pc._write_stage2_blocking_message(msg)
        except Exception:
            pass
        try:
            snap = pc.interactive_trial_shortages_snapshot()
            shortage_path.write_text(
                json.dumps(
                    {
                        "format_version": 2,
                        "source_json": str(path),
                        "note": "validation failed before/during stage2",
                        "error": msg,
                        "op_shortage": snap["op_shortage"],
                        "as_shortage": snap["as_shortage"],
                    },
                    ensure_ascii=False,
                    indent=2,
                )
                + "\n",
                encoding="utf-8",
            )
        except Exception:
            pass
        return 3
    except Exception as e:
        print(f"dispatch trial failed: {e}", file=sys.stderr)
        traceback.print_exc()
        return 1

    export_script = SCRIPT_DIR / "export_result_dispatch_from_json.py"
    if export_script.is_file():
        print("[dispatch trial] 結果Excel(xlsx)をエクスポート中…", flush=True)
        py = sys.executable or "python3"
        try:
            subprocess.run(
                [py, str(export_script), str(path)],
                cwd=str(SCRIPT_DIR),
                check=True,
                timeout=600,
            )
            print("[dispatch trial] xlsx エクスポート完了。", flush=True)
        except Exception as e:
            print(f"xlsx export warning: {e}", file=sys.stderr)

    print(str(shortage_path), flush=True)
    return 0


if __name__ == "__main__":
    try:
        import workbook_env_bootstrap as _wbe_exit

        sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
    except ImportError:
        sys.exit(main())

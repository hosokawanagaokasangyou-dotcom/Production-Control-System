# -*- coding: utf-8 -*-
"""
設備ガント／実績明細ガントを「描画契約 JSON」から再実行するための入出力。

`_write_results_equipment_gantt_sheet` に渡す引数束を JSON 化し、
デコードして同一関数を呼び直せば同じシートを再生成できる（ExcelWriter が必要）。

段階2では計画ブック保存先と同名ベースで次を出力する想定:

- ``*_equipment_gantt_contract.json`` … 計画設備ガント
- ``*_actual_detail_gantt_contract.json`` … 加工実績明細ガント
"""

from __future__ import annotations

import json
import logging
import os
import tempfile
import unicodedata
from datetime import date, datetime, time, timedelta
from typing import Any

logger = logging.getLogger(__name__)

GANTT_CONTRACT_SCHEMA_VERSION = 1


def encode_value(obj: Any) -> Any:
    """datetime / date / time / timedelta・tuple・ネストを JSON 安全形へ。"""
    if obj is None:
        return None
    if isinstance(obj, datetime):
        return {"__t": "datetime", "v": obj.isoformat()}
    if isinstance(obj, date):
        return {"__t": "date", "v": obj.isoformat()}
    if isinstance(obj, time):
        return {"__t": "time", "v": obj.isoformat()}
    if isinstance(obj, timedelta):
        return {"__t": "timedelta", "total_seconds": obj.total_seconds()}
    if isinstance(obj, tuple):
        return {"__t": "tuple", "items": [encode_value(x) for x in obj]}
    if isinstance(obj, dict):
        out: dict[str, Any] = {}
        for k, v in obj.items():
            sk = (
                json.dumps(["__key_tuple__", encode_value(k)], ensure_ascii=False)
                if isinstance(k, tuple)
                else str(k)
            )
            out[sk] = encode_value(v)
        return out
    if isinstance(obj, list):
        return [encode_value(x) for x in obj]
    if isinstance(obj, (int, float, str, bool)):
        return obj
    try:
        import numpy as np

        if isinstance(obj, (np.integer,)):
            return int(obj)
        if isinstance(obj, (np.floating,)):
            return float(obj)
        if isinstance(obj, np.bool_):
            return bool(obj)
    except Exception:
        pass
    return str(obj)


def _decode_dict_key(k: str) -> Any:
    if isinstance(k, str) and k.startswith("["):
        try:
            parsed = json.loads(k)
            if isinstance(parsed, list) and len(parsed) == 2 and parsed[0] == "__key_tuple__":
                return decode_value(parsed[1])
        except Exception:
            pass
    return k


def decode_value(obj: Any) -> Any:
    """encode_value の逆変換（dict / list を再帰）。"""
    if obj is None:
        return None
    if isinstance(obj, dict):
        if "__t" in obj:
            t = obj["__t"]
            if t == "datetime":
                return datetime.fromisoformat(str(obj["v"]))
            if t == "date":
                return date.fromisoformat(str(obj["v"]))
            if t == "time":
                s = str(obj["v"])
                return time.fromisoformat(s) if "T" not in s else datetime.fromisoformat(s).time()
            if t == "timedelta":
                return timedelta(seconds=float(obj["total_seconds"]))
            if t == "tuple":
                return tuple(decode_value(x) for x in obj["items"])
        out_d: dict[Any, Any] = {}
        for k, v in obj.items():
            dk = _decode_dict_key(k) if isinstance(k, str) else k
            out_d[dk] = decode_value(v)
        return out_d
    if isinstance(obj, list):
        return [decode_value(x) for x in obj]
    return obj


def make_gantt_render_contract(
    *,
    timeline_events: list,
    equipment_list: list,
    sorted_dates: list,
    attendance_data: dict,
    data_extract_dt_str: str | None,
    base_now_dt: datetime | None = None,
    actual_timeline_events: list | None = None,
    regular_shift_times: tuple[time | None, time | None] | None = None,
    plan_rows: bool = True,
    chart_title: str | None = None,
    sheet_name_override: str | None = None,
    gantt_compare_shape_styling: bool = False,
    compare_aladdin_qty_by_machine_date: dict | None = None,
    kind: str = "equipment_gantt",
) -> dict:
    """`_write_results_equipment_gantt_sheet` に対応する契約 dict を構築する。"""
    inner = {
        "timeline_events": encode_value(timeline_events),
        "equipment_list": list(equipment_list),
        "sorted_dates": encode_value(sorted_dates),
        "attendance_data": encode_value(attendance_data),
        "data_extract_dt_str": data_extract_dt_str,
        "base_now_dt": encode_value(base_now_dt),
        "actual_timeline_events": encode_value(actual_timeline_events),
        "regular_shift_times": encode_value(regular_shift_times),
        "plan_rows": bool(plan_rows),
        "chart_title": chart_title,
        "sheet_name_override": sheet_name_override,
        "gantt_compare_shape_styling": bool(gantt_compare_shape_styling),
        "compare_aladdin_qty_by_machine_date": encode_value(
            compare_aladdin_qty_by_machine_date
        ),
    }
    return {
        "schema_version": GANTT_CONTRACT_SCHEMA_VERSION,
        "kind": kind,
        "fn": "_write_results_equipment_gantt_sheet",
        "kwargs_packed": inner,
    }


def unpack_gantt_contract(contract: dict) -> dict:
    """契約から `_write_results_equipment_gantt_sheet` 用キーワード引数 dict を復元。"""
    inner = dict(contract.get("kwargs_packed") or {})
    rst = decode_value(inner.get("regular_shift_times"))
    if isinstance(rst, list) and len(rst) == 2:
        rst = (rst[0], rst[1])
    out = {
        "timeline_events": decode_value(inner.get("timeline_events")) or [],
        "equipment_list": list(inner.get("equipment_list") or []),
        "sorted_dates": decode_value(inner.get("sorted_dates")) or [],
        "attendance_data": decode_value(inner.get("attendance_data")) or {},
        "data_extract_dt_str": inner.get("data_extract_dt_str"),
        "base_now_dt": decode_value(inner.get("base_now_dt")),
        "actual_timeline_events": decode_value(inner.get("actual_timeline_events")),
        "regular_shift_times": rst,
        "plan_rows": bool(inner.get("plan_rows", True)),
        "chart_title": inner.get("chart_title"),
        "sheet_name_override": inner.get("sheet_name_override"),
        "gantt_compare_shape_styling": bool(inner.get("gantt_compare_shape_styling", False)),
        "compare_aladdin_qty_by_machine_date": decode_value(
            inner.get("compare_aladdin_qty_by_machine_date")
        ),
    }
    return out


def render_gantt_sheet_from_contract(writer, contract: dict):
    """
    契約 JSON（または同等 dict）から設備ガント系シートを 1 枚書き込む。
    戻り値: ``_write_results_equipment_gantt_sheet`` と同じ ``(label_specs, day_blocks)``。
    """
    from planning_core._core import _write_results_equipment_gantt_sheet

    kw = unpack_gantt_contract(contract)
    return _write_results_equipment_gantt_sheet(
        writer,
        kw["timeline_events"],
        kw["equipment_list"],
        kw["sorted_dates"],
        kw["attendance_data"],
        kw["data_extract_dt_str"],
        kw["base_now_dt"],
        actual_timeline_events=kw["actual_timeline_events"],
        regular_shift_times=kw["regular_shift_times"],
        plan_rows=kw["plan_rows"],
        chart_title=kw["chart_title"],
        sheet_name_override=kw["sheet_name_override"],
        gantt_compare_shape_styling=kw["gantt_compare_shape_styling"],
        compare_aladdin_qty_by_machine_date=kw["compare_aladdin_qty_by_machine_date"],
    )


def gantt_contract_json_path(plan_xlsx_final: str, *, which: str) -> str:
    """
    which: ``equipment`` | ``actual_detail``
    """
    base, _ = os.path.splitext(os.path.abspath(plan_xlsx_final))
    suf = "_equipment_gantt_contract.json" if which == "equipment" else "_actual_detail_gantt_contract.json"
    try:
        return unicodedata.normalize("NFC", base + suf)
    except Exception:
        return base + suf


def write_gantt_contract_json(plan_xlsx_final: str, which: str, contract: dict) -> tuple[str | None, str]:
    """契約 JSON をブックより先に書けるよう一時ファイル経由で保存（plan_workbook と同趣旨）。"""
    out_path = gantt_contract_json_path(plan_xlsx_final, which=which)
    parent = os.path.dirname(out_path)
    try:
        os.makedirs(parent, exist_ok=True)
    except OSError:
        pass
    last_err: Exception | None = None
    for label, tmp_dir in (("same_dir_as_output", parent), ("system_temp", tempfile.gettempdir())):
        tmp_path: str | None = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w",
                encoding="utf-8",
                newline="\n",
                delete=False,
                prefix="pm_ai_gantt_contract_",
                suffix=".tmp",
                dir=tmp_dir,
            ) as tmp:
                tmp_path = tmp.name
                json.dump(contract, tmp, ensure_ascii=False, indent=2)
                tmp.write("\n")
            try:
                os.replace(tmp_path, out_path)
            except OSError:
                import shutil

                shutil.move(tmp_path, out_path)
            return out_path, label
        except Exception as e:
            last_err = e
            if tmp_path and os.path.isfile(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass
            continue
    if last_err is not None:
        logger.warning("gantt_render_contract: JSON 書込に失敗: %s", last_err)
    return None, "failed"


def load_gantt_contract_json(path: str) -> dict:
    with open(path, encoding="utf-8-sig") as f:
        return json.load(f)


def render_gantt_from_contract_json_file(writer, contract_path: str):
    """保存済み契約ファイルからシートを書く。"""
    contract = load_gantt_contract_json(contract_path)
    return render_gantt_sheet_from_contract(writer, contract)


def render_book_with_gantt_contracts_only(
    xlsx_out_path: str,
    *,
    equipment_contract_path: str | None = None,
    actual_detail_contract_path: str | None = None,
) -> None:
    """
    ガント契約のみからブックを新規作成する（表シートは含まない）。
    検証・オフライン再生成用。
    """
    import pandas as pd

    paths = [p for p in (equipment_contract_path, actual_detail_contract_path) if p]
    if not paths:
        raise ValueError("少なくとも1つの契約 JSON パスが必要です")
    with pd.ExcelWriter(xlsx_out_path, engine="openpyxl") as writer:
        if equipment_contract_path:
            render_gantt_from_contract_json_file(writer, equipment_contract_path)
        if actual_detail_contract_path:
            render_gantt_from_contract_json_file(writer, actual_detail_contract_path)

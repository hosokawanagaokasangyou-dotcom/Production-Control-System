# -*- coding: utf-8 -*-
"""
成果ブックの「表データ JSON」と Excel の対応。

**段階2・表シート（設備時間割・タスク一覧・配台表など）**
計算で DataFrame を確定したあと、**先に** ``*_tabular_source.json``（正本）を書き、
同じペイロードから ``ExcelWriter`` へ行を載せる（``build_workbook_payload_from_dataframes`` →
``write_tabular_sheets_from_payload_to_excel_writer``）。

**最終ミラー JSON**（ビューア用 ``production_plan_multi_day_*.json``）
日付列の表示整形・ガント後処理など **保存済み xlsx の最終セル値** を
``workbook_payload_from_final_xlsx_file``（pandas + ガント系 reheader）で出力する。

**論理・ビュー JSON**（``*_logical_view.json``）
結合セルを値で展開してから同じ抽出へ渡したもの（UI グラフィック向け。
``planning_core.logical_workbook_view``／``write_production_plan_logical_view_json``）。

**値のみのブックを JSON から再生成** … ``write_xlsx_from_workbook_payload_tabular``
"""

from __future__ import annotations

import json
import logging
import math
import os
from datetime import date, datetime, time
from decimal import Decimal
from typing import Any

import pandas as pd

logger = logging.getLogger(__name__)

# plan_workbook_sidecar と同一（インポート循環回避のためローカル定義）
WORKBOOK_JSON_FORMAT_VERSION = 2


def json_serializable_cell_value(v: Any) -> Any:
    """JSON 行レコード用にセル値を正規化。"""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, time):
        return v.isoformat()
    if isinstance(v, bool):
        return v
    if isinstance(v, Decimal):
        try:
            return float(v)
        except Exception:
            return str(v)
    if isinstance(v, int):
        return int(v)
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None
        return float(v)
    return str(v)


def workbook_payload_from_final_xlsx_file(
    xlsx_path: str,
    *,
    source_xlsx_basename: str | None = None,
    format_version: int = WORKBOOK_JSON_FORMAT_VERSION,
    metadata_extra: dict | None = None,
) -> dict:
    """
    保存済み xlsx（段階2の後処理まで終えたファイル）から、計画結果ビューア用ペイロードを構築する。

    ``plan_workbook_sidecar._dump_xlsx_all_sheets_to_json`` と同等の表抽出（ガント系の reheader 含む）。
    """
    from .plan_workbook_sidecar import _reheader_dataframe_if_equipment_sheet_unnamed

    if not xlsx_path or not os.path.isfile(xlsx_path):
        raise FileNotFoundError(xlsx_path)

    base = source_xlsx_basename or os.path.basename(xlsx_path)
    sheets_in = pd.read_excel(xlsx_path, sheet_name=None, engine="openpyxl")
    sheets_out: dict[str, dict] = {}

    for name, df in (sheets_in or {}).items():
        if df is None or getattr(df, "empty", True):
            sheets_out[name] = {"columns": [], "row_count": 0, "rows": []}
            continue
        try:
            df = _reheader_dataframe_if_equipment_sheet_unnamed(name, df)
            rows = json.loads(
                df.to_json(orient="records", date_format="iso", double_precision=15)
            )
        except Exception as e:
            logger.warning(
                "workbook_payload: シート %r の行データ化に失敗: %s",
                name,
                e,
            )
            sheets_out[name] = {"columns": [], "row_count": 0, "rows": []}
            continue
        sheets_out[name] = {
            "columns": list(df.columns),
            "row_count": int(len(df)),
            "rows": rows,
        }

    payload: dict = {
        "format_version": format_version,
        "source_xlsx": base,
        "sheets": sheets_out,
    }
    if metadata_extra:
        payload.update(metadata_extra)
    return payload


def sheet_payload_to_dataframe(sheet: dict) -> pd.DataFrame:
    """単一シートのペイロードから DataFrame を復元する。"""
    cols = sheet.get("columns") or []
    rows = sheet.get("rows") or []
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows, columns=cols)


def build_workbook_payload_from_dataframes(
    sheets: dict[str, pd.DataFrame],
    *,
    source_xlsx_basename: str,
    format_version: int = WORKBOOK_JSON_FORMAT_VERSION,
    metadata_extra: dict | None = None,
) -> dict:
    """
    メモリ上の DataFrame 群から、計画結果ビューアと同形の ``sheets`` ペイロードを構築する。
    段階2の表シートを Excel より先に JSON 化する際に使う。
    """
    sheets_out: dict[str, dict] = {}
    for name, df in sheets.items():
        if df is None or getattr(df, "empty", True):
            sheets_out[name] = {"columns": [], "row_count": 0, "rows": []}
            continue
        try:
            rows = json.loads(
                df.to_json(orient="records", date_format="iso", double_precision=15)
            )
        except Exception as e:
            logger.warning(
                "workbook_payload: DataFrame→JSON 行に失敗 sheet=%r: %s",
                name,
                e,
            )
            sheets_out[name] = {"columns": [], "row_count": 0, "rows": []}
            continue
        sheets_out[name] = {
            "columns": list(df.columns),
            "row_count": int(len(df)),
            "rows": rows,
        }

    payload: dict = {
        "format_version": format_version,
        "source_xlsx": source_xlsx_basename,
        "sheets": sheets_out,
    }
    if metadata_extra:
        payload.update(metadata_extra)
    return payload


def tabular_source_json_path(plan_xlsx_final: str) -> str:
    """段階2計画ブックと並べて置く「表シート正本」JSON パス（拡張子手前にサフィックス）。"""
    base, _ = os.path.splitext(os.path.abspath(plan_xlsx_final))
    return base + "_tabular_source.json"


def write_tabular_source_json_file(
    plan_xlsx_final: str,
    payload: dict,
) -> tuple[str | None, str]:
    """
    表データ正本 JSON を **Excel ブックより先** に書き出す。
    戻り値: (出力パス or None, 使用した一時書込戦略ラベル)
    """
    from .plan_workbook_sidecar import _write_workbook_json_payload

    out = tabular_source_json_path(plan_xlsx_final)
    try:
        if os.path.isfile(out):
            os.remove(out)
    except OSError:
        pass
    final, strategy = _write_workbook_json_payload(out, payload)
    return final, strategy


def write_tabular_sheets_from_payload_to_excel_writer(
    writer,
    payload: dict,
    *,
    sheet_order: list[str],
) -> None:
    """
    ペイロードの表シートだけを ``sheet_order`` の順で ``ExcelWriter`` に書く。
    （人員配台優先の 2 段 ``startrow`` シート等は含めない）
    """
    sheets = payload.get("sheets") or {}
    for sn in sheet_order:
        if sn not in sheets:
            logger.warning(
                "workbook_payload: ペイロードにシート %r が無いためスキップします。",
                sn,
            )
            continue
        df = sheet_payload_to_dataframe(sheets[sn])
        safe = str(sn)[:31]
        df.to_excel(writer, sheet_name=safe, index=False)


def write_xlsx_from_workbook_payload_tabular(
    payload: dict,
    xlsx_out_path: str,
    *,
    engine: str = "openpyxl",
) -> None:
    """
    ペイロードの **表データのみ** を 1 ブックに書き出す（書式・図形・結合なし）。

    設備ガントなど複雑レイアウトの完全再現は目的としない。検証・差分用。
    """
    sheets = payload.get("sheets") or {}
    _parent = os.path.dirname(os.path.abspath(xlsx_out_path))
    if _parent:
        os.makedirs(_parent, exist_ok=True)
    with pd.ExcelWriter(xlsx_out_path, engine=engine) as writer:
        for sheet_name, sheet in sheets.items():
            if not isinstance(sheet, dict):
                continue
            df = sheet_payload_to_dataframe(sheet)
            safe_name = str(sheet_name)[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)

# -*- coding: utf-8 -*-
"""
段階2の成果ブック用 JSON サイドカー。

全シート表形式ミラーの組み立ては ``workbook_payload.workbook_payload_from_final_xlsx_file`` に集約。
値のみの xlsx を JSON から出力するには ``write_xlsx_from_workbook_payload_tabular`` を参照。
"""

from __future__ import annotations

import json
import logging
import os
import re
import shutil
import tempfile
import unicodedata

import pandas as pd

from .workbook_payload import (
    WORKBOOK_JSON_FORMAT_VERSION,
    workbook_payload_from_final_xlsx_file,
    write_xlsx_from_workbook_payload_tabular,
)

# 0/false/no/off/none: do not read or write sidecar JSON
ENV_PLAN_RESULT_TASK_JSON = "PM_AI_PLAN_RESULT_TASK_JSON"
ENV_PLAN_RESULT_TASK_JSON_PATH = "PM_AI_PLAN_RESULT_TASK_JSON_PATH"
# 0/false/no/off/none: do not write production_plan_multi_day_*.json (full workbook mirror)
ENV_PLAN_WORKBOOK_JSON = "PM_AI_PLAN_WORKBOOK_JSON"
# 0/false/no/off/none: do not write member_schedule_*.json (mirror of member_schedule_*.xlsx)
ENV_MEMBER_SCHEDULE_JSON = "PM_AI_MEMBER_SCHEDULE_JSON"

RESULT_TASK_JSON_SUFFIX = "_\u7d50\u679c_\u30bf\u30b9\u30af\u4e00\u89a7.json"
SIDE_FORMAT_VERSION = 1

# pandas の read_excel(header=0) だと、結果_設備ガントのようにタイトル行のあとに列見出し行がある
# シートでは列名が Unnamed: 0 になる。JSON 側で HH:MM 列を復元する。
_GANTT_TIME_HEADER_RE = re.compile(r"^\s*(\d{1,2}):(\d{2})\s*$")


def _cell_str_df(v) -> str:
    if v is None:
        return ""
    try:
        if isinstance(v, float) and pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def _looks_like_hh_mm_header(s: str) -> bool:
    return bool(s and _GANTT_TIME_HEADER_RE.match(s))


def _reheader_dataframe_if_equipment_sheet_unnamed(name: str, df: pd.DataFrame) -> pd.DataFrame:
    """
    設備ガント系シートで pandas が Unnamed 列だけ返したとき、先頭数十行から
    A 列が「日付」かつ HH:MM セルが複数ある行を列見出し行として採用する。
    """
    if df is None or getattr(df, "empty", True):
        return df
    cols = [str(x) for x in df.columns]
    if not cols or not str(cols[0]).startswith("Unnamed"):
        return df
    if name and ("設備" not in name and "時間割" not in name):
        return df
    ncols = len(df.columns)
    max_scan = min(50, len(df))
    for r in range(max_scan):
        row = df.iloc[r]
        if _cell_str_df(row.iloc[0] if len(row) > 0 else None) != "日付":
            continue
        vals = [_cell_str_df(row.iloc[i]) if i < len(row) else "" for i in range(ncols)]
        first_t = -1
        for c in range(1, len(vals)):
            if _looks_like_hh_mm_header(vals[c]):
                first_t = c
                break
        if first_t < 0:
            continue
        time_hits = sum(
            1 for c in range(first_t, len(vals)) if _looks_like_hh_mm_header(vals[c])
        )
        if time_hits < 2:
            continue
        header = [_cell_str_df(row.iloc[i]) if i < len(row) else "" for i in range(ncols)]
        out = df.iloc[r + 1 :].copy()
        out.columns = header
        return out.reset_index(drop=True)
    return df


def _normalized_json_sidecar_path(xlsx_path: str) -> str:
    """xlsx と同名の .json パス（絶対・正規化・NFC）。和文パスで実体と表記揺れする場合の対策。"""
    p = os.path.splitext(os.path.abspath(os.path.normpath(xlsx_path)))[0] + ".json"
    try:
        return unicodedata.normalize("NFC", p)
    except Exception:
        return p


def normalized_workbook_json_path(xlsx_path: str) -> str:
    """計画ブック・メンバー別スケジュールと同名ベースの .json 出力パス（xlsx が無くても論理名に使える）。"""
    return _normalized_json_sidecar_path(xlsx_path)


def _write_workbook_json_payload(out_path: str, payload: dict) -> tuple[str, str]:
    """
    一時ファイルへ書き、os.replace / shutil.move で最終名へ。同一フォルダで open(w) だけが errno 2 になる
    環境（WSL/ドライブ併用・一部 Unicode 正規化差）に強い。戻り値: (out_path, temp_dir_used 説明文字列)
    """
    out_path = unicodedata.normalize("NFC", out_path)
    parent = os.path.dirname(out_path)
    try:
        os.makedirs(parent, exist_ok=True)
    except OSError:
        pass
    last_err: Exception | None = None
    for label, tmp_dir in (
        ("same_dir_as_output", parent),
        ("system_temp", tempfile.gettempdir()),
    ):
        tmp_path: str | None = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w",
                encoding="utf-8",
                newline="\n",
                delete=False,
                prefix="pm_ai_wb_json_",
                suffix=".tmp",
                dir=tmp_dir,
            ) as tmp:
                tmp_path = tmp.name
                json.dump(payload, tmp, ensure_ascii=False, indent=2)
                tmp.write("\n")
            try:
                os.replace(tmp_path, out_path)
            except OSError:
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
        raise last_err
    raise OSError("JSON sidecar: could not write via temporary file")


def _plan_result_task_json_disabled() -> bool:
    v = (os.environ.get(ENV_PLAN_RESULT_TASK_JSON) or "").strip().lower()
    return v in ("0", "false", "no", "off", "none")


def _plan_workbook_json_disabled() -> bool:
    v = (os.environ.get(ENV_PLAN_WORKBOOK_JSON) or "").strip().lower()
    return v in ("0", "false", "no", "off", "none")


def _member_schedule_json_disabled() -> bool:
    v = (os.environ.get(ENV_MEMBER_SCHEDULE_JSON) or "").strip().lower()
    return v in ("0", "false", "no", "off", "none")


def _dump_xlsx_all_sheets_to_json(
    xlsx_path: str,
    *,
    payload_extra: dict | None = None,
    json_out_path: str | None = None,
) -> str | None:
    """
    保存済み xlsx を ``workbook_payload.workbook_payload_from_final_xlsx_file`` で表形式ペイロード化し、
    同名ベースの .json に書き出す内部共通処理（ガント系の列見出し reheader 込み）。

    json_out_path を指定したときはそのパスへ出力（一時 xlsx から最終 JSON 名へ出す用）。
    """
    if not xlsx_path or not os.path.isfile(xlsx_path):
        logging.warning(
            "plan_workbook_sidecar: 同名 JSON をスキップ（xlsx 不在） path=%s",
            xlsx_path,
        )
        return None
    out_path = json_out_path if json_out_path else _normalized_json_sidecar_path(xlsx_path)
    try:
        try:
            if os.path.isfile(out_path):
                os.remove(out_path)
        except OSError:
            pass
        payload = workbook_payload_from_final_xlsx_file(
            xlsx_path,
            source_xlsx_basename=os.path.basename(xlsx_path),
            metadata_extra=payload_extra,
        )
    except Exception as e:
        logging.warning(
            "plan_workbook_sidecar: ブックのペイロード化に失敗したため同名 JSON を出せません: %s (%s)",
            xlsx_path,
            e,
        )
        return None
    try:
        _out_final, _tmp_strategy = _write_workbook_json_payload(out_path, payload)
        logging.info(
            "plan_workbook_sidecar: 同名 JSON を出力しました path=%s (%s)",
            _out_final,
            _tmp_strategy,
        )
        return _out_final
    except Exception as e:
        logging.warning(
            "plan_workbook_sidecar: JSON ファイル書き込みに失敗: %s (%s)",
            out_path,
            e,
        )
        return None


def result_task_sidecar_path(plan_xlsx: str) -> str:
    base, _ = os.path.splitext(plan_xlsx)
    return base + RESULT_TASK_JSON_SUFFIX


def read_result_task_dataframe(plan_xlsx: str) -> pd.DataFrame | None:
    """
    Returns None if sidecar is disabled, missing, or invalid (caller uses read_excel).
    """
    if not plan_xlsx or _plan_result_task_json_disabled():
        return None
    ex = (os.environ.get(ENV_PLAN_RESULT_TASK_JSON_PATH) or "").strip()
    if ex and os.path.isfile(ex):
        p = ex
    else:
        p = result_task_sidecar_path(plan_xlsx)
        if not os.path.isfile(p):
            return None
    try:
        with open(p, encoding="utf-8-sig") as f:
            data = json.load(f)
        if isinstance(data, dict) and "rows" in data:
            rows = data["rows"]
        elif isinstance(data, list):
            rows = data
        else:
            return pd.DataFrame()
        if not rows:
            return pd.DataFrame()
        return pd.DataFrame(rows)
    except Exception:
        return None


def write_result_task_json_sidecar(plan_xlsx: str, df: pd.DataFrame, *, sheet_name: str) -> str | None:
    if _plan_result_task_json_disabled():
        logging.debug(
            "plan_workbook_sidecar: skip result-task JSON (%s disables)",
            ENV_PLAN_RESULT_TASK_JSON,
        )
        return None
    _sidecar_out = result_task_sidecar_path(plan_xlsx) if plan_xlsx else ""
    try:
        if df is None or getattr(df, "empty", True):
            logging.warning(
                "plan_workbook_sidecar: 結果_タスク一覧 JSON をスキップ（DataFrame が空） plan_xlsx=%s",
                plan_xlsx,
            )
            return None
        out = result_task_sidecar_path(plan_xlsx)
        rows = json.loads(df.to_json(orient="records", date_format="iso", double_precision=15))
        payload = {
            "format_version": SIDE_FORMAT_VERSION,
            "sheet_name": sheet_name,
            "columns": list(df.columns),
            "row_count": int(len(df)),
            "rows": rows,
        }
        # 直 open(w) は同一フォルダで errno 2 になる環境があるため、全ブック JSON と同じ一時ファイル経路に統一
        out, _ = _write_workbook_json_payload(out, payload)
        return out
    except Exception as e:
        logging.warning(
            "plan_workbook_sidecar: 結果_タスク一覧 JSON 書き出し失敗: %s plan_xlsx=%r out=%r",
            e,
            plan_xlsx,
            _sidecar_out,
            exc_info=True,
        )
        return None


ENV_PLAN_LOGICAL_VIEW_JSON = "PM_AI_PLAN_LOGICAL_VIEW_JSON"


def _plan_logical_view_json_disabled() -> bool:
    v = (os.environ.get(ENV_PLAN_LOGICAL_VIEW_JSON) or "").strip().lower()
    return v in ("0", "false", "no", "off", "none")


def write_production_plan_logical_view_json(
    plan_xlsx: str,
    *,
    json_out_path: str | None = None,
) -> str | None:
    """
    結合セルを値で展開したうえで全シートを表形式化した **論理・ビュー用 JSON** を書き出す。

    出力パス: 既定は ``<plan stem>_logical_view.json``（:func:`logical_workbook_view.logical_view_json_path`）。

    無効化: 環境変数 ``PM_AI_PLAN_LOGICAL_VIEW_JSON=0`` 等。

    通常の ``production_plan_multi_day*.json``（xlsx 直読みミラー）と併用し、
    設備ガントの 10 分枠を UI で使う場合は本ファイルを正とする。
    """
    if _plan_logical_view_json_disabled():
        logging.info(
            "plan_workbook_sidecar: 論理ビュー JSON はスキップします "
            "(%s=%r)",
            ENV_PLAN_LOGICAL_VIEW_JSON,
            (os.environ.get(ENV_PLAN_LOGICAL_VIEW_JSON) or "").strip(),
        )
        return None
    if not plan_xlsx or not os.path.isfile(plan_xlsx):
        logging.warning(
            "plan_workbook_sidecar: 論理ビュー JSON（xlsx 不在） path=%s",
            plan_xlsx,
        )
        return None
    try:
        from .logical_workbook_view import (
            build_logical_view_workbook_payload,
            logical_view_json_path,
        )

        payload = build_logical_view_workbook_payload(
            plan_xlsx,
            source_xlsx_basename=os.path.basename(plan_xlsx),
        )
        out_path = json_out_path or logical_view_json_path(plan_xlsx)
        final, strategy = _write_workbook_json_payload(out_path, payload)
        logging.info(
            "plan_workbook_sidecar: 論理ビュー JSON を出力 path=%s (%s)",
            final,
            strategy,
        )
        return final
    except Exception as e:
        logging.warning(
            "plan_workbook_sidecar: 論理ビュー JSON 書き出し失敗: %s (%s)",
            plan_xlsx,
            e,
            exc_info=True,
        )
        return None


def write_production_plan_workbook_json(
    plan_xlsx: str,
    *,
    json_out_path: str | None = None,
    metadata_extra: dict | None = None,
) -> str | None:
    """
    ``production_plan_multi_day_*.xlsx`` と同名ベースの ``.json`` に、全シートを表形式で書き出す。
    図形・セル書式は含まず、セル値のみ。

    抽出ロジックは ``workbook_payload.workbook_payload_from_final_xlsx_file`` に集約。
    値のみのブックを JSON から再生成するには ``write_xlsx_from_workbook_payload_tabular`` を参照。
    """
    if _plan_workbook_json_disabled():
        logging.info(
            "plan_workbook_sidecar: production_plan 同名 JSON は無効のためスキップします "
            "(PM_AI_PLAN_WORKBOOK_JSON=%r)",
            (os.environ.get(ENV_PLAN_WORKBOOK_JSON) or "").strip(),
        )
        return None
    extra = dict(metadata_extra) if metadata_extra else None
    return _dump_xlsx_all_sheets_to_json(
        plan_xlsx, payload_extra=extra, json_out_path=json_out_path
    )


def write_member_schedule_workbook_json(
    member_xlsx: str,
    *,
    json_out_path: str | None = None,
    metadata_extra: dict | None = None,
) -> str | None:
    """
    ``member_schedule_*.xlsx`` と同名ベースの ``.json`` に全シートを書き出す（計画ブック JSON と同形式。
    シート名はメンバー名）。
    """
    if _member_schedule_json_disabled():
        logging.info(
            "plan_workbook_sidecar: member_schedule 同名 JSON は無効のためスキップします "
            "(PM_AI_MEMBER_SCHEDULE_JSON=%r)",
            (os.environ.get(ENV_MEMBER_SCHEDULE_JSON) or "").strip(),
        )
        return None
    extra: dict = {"workbook_kind": "member_schedule"}
    if metadata_extra:
        extra.update(metadata_extra)
    return _dump_xlsx_all_sheets_to_json(
        member_xlsx, payload_extra=extra, json_out_path=json_out_path
    )

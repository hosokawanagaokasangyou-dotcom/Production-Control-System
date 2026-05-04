# -*- coding: utf-8 -*-
"""JSON sidecars for stage-2 plan workbook (result task sheet) to limit Excel re-read."""

from __future__ import annotations

import json
import logging
import os
import shutil
import tempfile
import time
import unicodedata

import pandas as pd

# 0/false/no/off/none: do not read or write sidecar JSON
ENV_PLAN_RESULT_TASK_JSON = "PM_AI_PLAN_RESULT_TASK_JSON"
ENV_PLAN_RESULT_TASK_JSON_PATH = "PM_AI_PLAN_RESULT_TASK_JSON_PATH"
# 0/false/no/off/none: do not write production_plan_multi_day_*.json (full workbook mirror)
ENV_PLAN_WORKBOOK_JSON = "PM_AI_PLAN_WORKBOOK_JSON"
# 0/false/no/off/none: do not write member_schedule_*.json (mirror of member_schedule_*.xlsx)
ENV_MEMBER_SCHEDULE_JSON = "PM_AI_MEMBER_SCHEDULE_JSON"

RESULT_TASK_JSON_SUFFIX = "_\u7d50\u679c_\u30bf\u30b9\u30af\u4e00\u89a7.json"
SIDE_FORMAT_VERSION = 1
WORKBOOK_JSON_FORMAT_VERSION = 2


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
    xlsx を pandas で全シート読み、同名ベースの .json に書き出す内部共通処理。
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
        sheets_in = pd.read_excel(xlsx_path, sheet_name=None, engine="openpyxl")
    except Exception as e:
        logging.warning(
            "plan_workbook_sidecar: read_excel に失敗したため同名 JSON を出せません: %s (%s)",
            xlsx_path,
            e,
        )
        return None
    sheets_out: dict[str, dict] = {}
    for name, df in (sheets_in or {}).items():
        if df is None or getattr(df, "empty", True):
            sheets_out[name] = {"columns": [], "row_count": 0, "rows": []}
            continue
        try:
            rows = json.loads(
                df.to_json(orient="records", date_format="iso", double_precision=15)
            )
        except Exception as e:
            logging.warning(
                "plan_workbook_sidecar: シート %r の行データ化に失敗し同名 JSON を中断: %s",
                name,
                e,
            )
            return None
        sheets_out[name] = {
            "columns": list(df.columns),
            "row_count": int(len(df)),
            "rows": rows,
        }
    payload = {
        "format_version": WORKBOOK_JSON_FORMAT_VERSION,
        "source_xlsx": os.path.basename(xlsx_path),
        "sheets": sheets_out,
    }
    if payload_extra:
        payload.update(payload_extra)
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


def _append_agent_debug_ndjson(payload: dict) -> None:
    """Cursor デバッグ用 NDJSON（session 471ee7）。PM_AI_REPO_ROOT が無い場合は何もしない。"""
    try:
        rr = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
        if not rr:
            return
        log_path = os.path.join(rr, ".cursor", "debug-471ee7.log")
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        line = {"sessionId": "471ee7", "timestamp": int(time.time() * 1000), **payload}
        with open(log_path, "a", encoding="utf-8") as fp:
            fp.write(json.dumps(line, ensure_ascii=False) + "\n")
    except Exception:
        pass


def write_result_task_json_sidecar(plan_xlsx: str, df: pd.DataFrame, *, sheet_name: str) -> str | None:
    # #region agent log
    _sidecar_out = result_task_sidecar_path(plan_xlsx) if plan_xlsx else ""
    # #endregion
    if _plan_result_task_json_disabled():
        # #region agent log
        _append_agent_debug_ndjson(
            {
                "hypothesisId": "SC1",
                "location": "plan_workbook_sidecar.write_result_task_json_sidecar",
                "message": "sidecar disabled by PM_AI_PLAN_RESULT_TASK_JSON",
                "data": {
                    "planXlsx": plan_xlsx,
                    "envPlanResultTaskJson": (os.environ.get(ENV_PLAN_RESULT_TASK_JSON) or "").strip(),
                },
            }
        )
        # #endregion
        logging.debug(
            "plan_workbook_sidecar: skip result-task JSON (%s disables)",
            ENV_PLAN_RESULT_TASK_JSON,
        )
        return None
    try:
        if df is None or getattr(df, "empty", True):
            # #region agent log
            _append_agent_debug_ndjson(
                {
                    "hypothesisId": "SC1",
                    "location": "plan_workbook_sidecar.write_result_task_json_sidecar",
                    "message": "sidecar skip empty df_tasks",
                    "data": {"planXlsx": plan_xlsx, "sidecarPath": _sidecar_out},
                }
            )
            # #endregion
            logging.warning(
                "plan_workbook_sidecar: 結果_タスク一覧 JSON をスキップ（DataFrame が空） plan_xlsx=%s",
                plan_xlsx,
            )
            return None
        out = result_task_sidecar_path(plan_xlsx)
        # open(w) 前に親フォルダを必ず作成（xlsx 出力先と同階層でも、先に作られていないと Errno 2 になる）
        _out_dir = os.path.dirname(out)
        if _out_dir:
            os.makedirs(_out_dir, exist_ok=True)
        try:
            if os.path.isfile(out):
                os.remove(out)
        except OSError:
            pass
        rows = json.loads(df.to_json(orient="records", date_format="iso", double_precision=15))
        payload = {
            "format_version": SIDE_FORMAT_VERSION,
            "sheet_name": sheet_name,
            "columns": list(df.columns),
            "row_count": int(len(df)),
            "rows": rows,
        }
        with open(out, encoding="utf-8", newline="\n") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            f.write("\n")
        # #region agent log
        _append_agent_debug_ndjson(
            {
                "hypothesisId": "SC1",
                "location": "plan_workbook_sidecar.write_result_task_json_sidecar",
                "message": "sidecar write ok",
                "data": {"planXlsx": plan_xlsx, "sidecarPath": out, "rowCount": int(len(df))},
            }
        )
        # #endregion
        return out
    except Exception as e:
        # #region agent log
        _append_agent_debug_ndjson(
            {
                "hypothesisId": "SC1",
                "location": "plan_workbook_sidecar.write_result_task_json_sidecar",
                "message": "sidecar write exception",
                "data": {
                    "planXlsx": plan_xlsx,
                    "sidecarPath": _sidecar_out,
                    "error": str(e)[:500],
                },
            }
        )
        # #endregion
        logging.warning(
            "plan_workbook_sidecar: 結果_タスク一覧 JSON 書き出し失敗: %s plan_xlsx=%r out=%r",
            e,
            plan_xlsx,
            _sidecar_out,
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
    図形・セル書式は含まず、セル値のみ（Excel 再読込と同様）。
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

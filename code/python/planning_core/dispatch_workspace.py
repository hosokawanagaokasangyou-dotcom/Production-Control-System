# -*- coding: utf-8 -*-
"""
Tabular inputs, workspace, and UNC defaults (parity with AppPaths / plan/*.m).

加工計画DATA相当の列構成の参照は plan/01_加工計画DATA_単一ファイル.m の Power Query 成形後と一致させるのが前提。
工程別生産計画問合せブックは「6行目＋5行目」の複合見出しが必要なため、自動検出時は plan/工程別生産計画問合せ_成形.m と同趣旨の成形を
read_tabular_dataframe 内で適用する（無効化: PM_AI_KOUBAI_INQUIRY_SHAPING=0）。

See EnvVarDocs: PM_AI_TASK_INPUT_SOURCE_DIR, PM_AI_ACTUAL_DETAIL_SOURCE_DIR, PM_AI_RESULT_DISPATCH_TABLE_DIR,
PM_AI_PROCESSING_PLAN_HEADER_ROW.
"""

from __future__ import annotations

import json
import logging
import os
import pathlib
import re
import time
import unicodedata

import pandas as pd

ENV_PM_AI_WORKSPACE = "PM_AI_WORKSPACE"
ENV_PROCESSING_PLAN_PATH = "PM_AI_PROCESSING_PLAN_PATH"
ENV_PLAN_INPUT_PATH = "PM_AI_PLAN_INPUT_PATH"
ENV_PROCESSING_PLAN_SHEET = "PM_AI_PROCESSING_PLAN_SHEET"
ENV_TABULAR_CSV_ENCODING = "PM_AI_TABULAR_CSV_ENCODING"
# 1-based Excel row where column headers are (e.g. 6 for 工程別 layouts with 5 metadata rows).
ENV_PROCESSING_PLAN_HEADER_ROW = "PM_AI_PROCESSING_PLAN_HEADER_ROW"
# auto / 1 = koubai inquiry shape (merge 6+5, drop 加工時間/加工速度, strip 加工数量 except exact-only column, dates); 0 = off
ENV_KOUBAI_INQUIRY_SHAPING = "PM_AI_KOUBAI_INQUIRY_SHAPING"
ENV_ACTUAL_DETAIL_SOURCE_DIR = "PM_AI_ACTUAL_DETAIL_SOURCE_DIR"
ENV_ACTUAL_DETAIL_WORKBOOK = "PM_AI_ACTUAL_DETAIL_WORKBOOK"
ENV_RESULT_DISPATCH_TABLE_DIR = "PM_AI_RESULT_DISPATCH_TABLE_DIR"
ENV_PM_AI_REPO_ROOT = "PM_AI_REPO_ROOT"
ENV_TASK_INPUT_SOURCE_DIR = "PM_AI_TASK_INPUT_SOURCE_DIR"

_LOG = logging.getLogger(__name__)

# CSV / Parquet / Excel task-input files (read_tabular_dataframe).
_PROCESSING_PLAN_FILE_SUFFIXES = (
    ".csv",
    ".parquet",
    ".pq",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
)

# Same logical path as AppPaths.DEFAULT_PQ_A_SOURCE_UNC (PQ 01 / PM_AI_TASK_INPUT_SOURCE_DIR).
DEFAULT_TASK_INPUT_SOURCE_DIR = (
    "\\\\192.168.0.101\\"
    "\u5171\u6709\u30d5\u30a9\u30eb\u30c0\\"
    "\u6e56\u5357\u5de5\u5834\\"
    "\u6e56\u5357\u5171\u6709\\"
    "\u751f\u7523\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0\\"
    "\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0\\"
    "\u25cfDATA\\"
    "\u751f\u7523\u8a08\u753b\u554f\u5408\u305b"
)

# Same as plan/02 PQ Folder.Files (PM_AI_ACTUAL_DETAIL_SOURCE_DIR).
DEFAULT_ACTUAL_DETAIL_SOURCE_DIR = (
    "\\\\192.168.0.101\\"
    "\u5171\u6709\u30d5\u30a9\u30eb\u30c0\\"
    "\u6e56\u5357\u5de5\u5834\\"
    "\u6e56\u5357\u5171\u6709\\"
    "002  \u52a0\u5de5G\\"
    "\u25cf\u691c\u67fb\u8868\u4f5c\u6210\\"
    "\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30DATA"
)


def pm_ai_workspace_dir() -> str:
    """Return absolute workspace dir if valid, else empty string."""
    p = (os.environ.get(ENV_PM_AI_WORKSPACE) or "").strip()
    return os.path.abspath(p) if p and os.path.isdir(p) else ""


def plan_input_workbook_path_for_excel_ops() -> str:
    """配台計画タスク入力を置く Excel ブック（専用 UI の ``PM_AI_PLAN_INPUT_PATH``）。

    実在する ``.xlsx`` / ``.xlsm`` / テンプレ拡張子のときだけ絶対パスを返す。
    CSV / Parquet のときは空（ブック I/O 対象外）。
    """
    p = (os.environ.get(ENV_PLAN_INPUT_PATH) or "").strip()
    if not p or not os.path.isfile(p):
        return ""
    low = p.lower()
    if low.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        return os.path.normpath(os.path.abspath(p))
    return ""


def pick_newest_excel_in_dir(dir_path: str) -> str | None:
    """Newest .xlsx/.xlsm by max(mtime, atime). Skips Excel lock files (~$)."""
    best: str | None = None
    best_key = -1.0
    try:
        for name in os.listdir(dir_path):
            if name.startswith("~$"):
                continue
            low = name.lower()
            if not low.endswith((".xlsx", ".xlsm")):
                continue
            fp = os.path.join(dir_path, name)
            if not os.path.isfile(fp):
                continue
            try:
                st = os.stat(fp)
                t = max(float(st.st_mtime), float(getattr(st, "st_atime", 0.0)))
            except OSError:
                continue
            if t > best_key:
                best_key = t
                best = fp
    except OSError:
        return None
    return best


def pick_newest_processing_plan_file_in_dir(dir_path: str) -> str | None:
    """Newest tabular plan file by max(mtime, atime). Skips Excel lock files (~$)."""
    best: str | None = None
    best_key = -1.0
    try:
        for name in os.listdir(dir_path):
            if name.startswith("~$"):
                continue
            low = name.lower()
            if not low.endswith(_PROCESSING_PLAN_FILE_SUFFIXES):
                continue
            fp = os.path.join(dir_path, name)
            if not os.path.isfile(fp):
                continue
            try:
                st = os.stat(fp)
                t = max(float(st.st_mtime), float(getattr(st, "st_atime", 0.0)))
            except OSError:
                continue
            if t > best_key:
                best_key = t
                best = fp
    except OSError:
        return None
    return best


def resolve_processing_plan_path_from_env() -> str | None:
    """
    Effective PM_AI_PROCESSING_PLAN_PATH:

    1. If PM_AI_PROCESSING_PLAN_PATH points to an existing file -> normalize to absolute and set env.
    2. Else if PM_AI_TASK_INPUT_SOURCE_DIR is a directory -> pick newest tabular file there,
       set PM_AI_PROCESSING_PLAN_PATH to that path (stage1 / load_tasks_df use it).

    Returns the resolved absolute path, or None if no file was resolved.
    """
    explicit = (os.environ.get(ENV_PROCESSING_PLAN_PATH) or "").strip()
    if explicit:
        if os.path.isfile(explicit):
            ap = os.path.normpath(os.path.abspath(explicit))
            os.environ[ENV_PROCESSING_PLAN_PATH] = ap
            return ap
        _LOG.warning(
            "PM_AI_PROCESSING_PLAN_PATH is not a file (%r). Trying PM_AI_TASK_INPUT_SOURCE_DIR.",
            explicit,
        )

    src = (os.environ.get(ENV_TASK_INPUT_SOURCE_DIR) or "").strip()
    if not src:
        return None
    src_abs = os.path.normpath(os.path.abspath(src))
    if not os.path.isdir(src_abs):
        _LOG.warning("PM_AI_TASK_INPUT_SOURCE_DIR is not a directory: %r", src)
        return None

    picked = pick_newest_processing_plan_file_in_dir(src_abs)
    if not picked:
        _LOG.warning(
            "No CSV/Parquet/Excel task-input files under PM_AI_TASK_INPUT_SOURCE_DIR: %s",
            src_abs,
        )
        return None

    ap = os.path.normpath(os.path.abspath(picked))
    os.environ[ENV_PROCESSING_PLAN_PATH] = ap
    _LOG.info(
        "PM_AI_PROCESSING_PLAN_PATH set to newest file under PM_AI_TASK_INPUT_SOURCE_DIR: %s",
        ap,
    )
    return ap


def resolve_actual_detail_workbook_path(task_input_workbook: str) -> str | None:
    """
    Actual-detail workbook path. load_machining_actuals_df reuses this resolution when
    PM_AI_ACTUALS_DATA_WORKBOOK is unset (input_resolution.resolve_actuals_workbook_path).

    Single-file override, then newest workbook under ACTUAL_DETAIL_SOURCE_DIR (or default UNC),
    then dirname(PM_AI_PLAN_INPUT_PATH) when that path is an Excel file.
    """
    wb_explicit = (os.environ.get(ENV_ACTUAL_DETAIL_WORKBOOK) or "").strip()
    if wb_explicit and os.path.isfile(wb_explicit):
        return wb_explicit
    d = (os.environ.get(ENV_ACTUAL_DETAIL_SOURCE_DIR) or "").strip()
    if not d:
        d = DEFAULT_ACTUAL_DETAIL_SOURCE_DIR
    if d and os.path.isdir(d):
        picked = pick_newest_excel_in_dir(d)
        if picked:
            return picked
    tw = (task_input_workbook or "").strip()
    if tw and os.path.isfile(tw):
        return tw
    return None


def resolve_result_dispatch_table_output_dir(task_input_workbook: str) -> str:
    """Output folder for standalone result dispatch table xlsx (fixed filename).

    Priority: PM_AI_RESULT_DISPATCH_TABLE_DIR, dirname(PM_AI_PLAN_INPUT_PATH),
    dirname(task_input_workbook argument), PM_AI_REPO_ROOT/code.

    ``plan_input_workbook_path_for_excel_ops`` が空でも、環境変数の計画タスク入力パスがあれば
    その親フォルダ（通常 ``.../output``）へ出す（専用 UI・Power Query 用ファイルの置き場所を揃える）。
    リポジトリ既定は ``{repo}/code/output``（フォルダが無くてもパスを返し、書き込み側が作成する）。
    """
    o = (os.environ.get(ENV_RESULT_DISPATCH_TABLE_DIR) or "").strip()
    if o and os.path.isdir(o):
        return os.path.abspath(o)
    pip = (os.environ.get(ENV_PLAN_INPUT_PATH) or "").strip()
    if pip and os.path.isfile(pip):
        low = pip.lower()
        if low.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
            return os.path.dirname(os.path.abspath(pip))
    wb = (task_input_workbook or "").strip()
    if wb and os.path.isfile(wb):
        return os.path.dirname(os.path.abspath(wb))
    repo = (os.environ.get(ENV_PM_AI_REPO_ROOT) or "").strip()
    if repo:
        code_root = os.path.join(os.path.abspath(repo), "code")
        if os.path.isdir(code_root):
            return os.path.join(code_root, "output")
    return ""


def _norm_sheet_key(name: str) -> str:
    return unicodedata.normalize("NFKC", (name or "").strip())


def _nfkc_header_token(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return unicodedata.normalize("NFKC", str(v).strip())


def _probe_excel_no_header(
    path: str, sheet_name: str | int, *, nrows: int = 60
) -> pd.DataFrame | None:
    """Top rows without interpreting row 0 as header (for header-row detection)."""
    try:
        return pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=None,
            nrows=nrows,
            engine="calamine",
        )
    except ImportError:
        pass
    except Exception as ex:  # noqa: BLE001
        _LOG.debug("probe calamine failed: %s", ex)
    try:
        return pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=None,
            nrows=nrows,
            engine="openpyxl",
            engine_kwargs={"read_only": True},
        )
    except Exception as ex_ro:  # noqa: BLE001
        _LOG.debug("probe openpyxl read_only failed: %s", ex_ro)
    try:
        return pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=None,
            nrows=nrows,
            engine="openpyxl",
        )
    except Exception as ex:  # noqa: BLE001
        _LOG.warning("probe excel failed (no header): path=%r sheet=%r err=%s", path, sheet_name, ex)
        return None


def _infer_tabular_excel_header_row_0based(path: str, sheet_name: str | int) -> int:
    """
    Find the row index (pandas ``header=``) where tabular headers live.

    工程別生産計画問合せ xlsx は先頭数行がメタデータで、表ヘッダーが 6 行目（Excel）のため
    既定の header=0 では列名がずれる。同一行に「依頼NO」「工程名」がある最初の行を採用する。
    """
    probe = _probe_excel_no_header(path, sheet_name, nrows=80)
    if probe is None or getattr(probe, "empty", True):
        return 0
    want_a = _nfkc_header_token("依頼NO")
    want_b = _nfkc_header_token("工程名")
    if not want_a or not want_b:
        return 0
    n = len(probe)
    for i in range(n):
        row = probe.iloc[i]
        tokens = {_nfkc_header_token(v) for v in row}
        tokens.discard("")
        if want_a in tokens and want_b in tokens:
            if i != 0:
                _LOG.info(
                    "Excel: header row auto-detected at Excel row %s (pandas header=%s). file=%r sheet=%r",
                    i + 1,
                    i,
                    path,
                    sheet_name,
                )
            return i
    return 0


def _resolve_tabular_excel_header_row_0based(path: str, sheet_name: str | int) -> int:
    raw = (os.environ.get(ENV_PROCESSING_PLAN_HEADER_ROW) or "").strip()
    if raw:
        try:
            excel_row = int(raw, 10)
        except ValueError:
            _LOG.warning(
                "%s=%r is not an integer; ignoring.",
                ENV_PROCESSING_PLAN_HEADER_ROW,
                raw,
            )
        else:
            if excel_row >= 1:
                return excel_row - 1
            _LOG.warning(
                "%s must be >= 1 (Excel row of headers); got %s",
                ENV_PROCESSING_PLAN_HEADER_ROW,
                raw,
            )
    return _infer_tabular_excel_header_row_0based(path, sheet_name)


def _resolve_tabular_sheet_name_calamine(path: str, sheet_name: str | int) -> str | int:
    """
    Calamine matches sheet names literally; Excel labels may differ by spaces / NFKC.
    List sheets via calamine only (do not use openpyxl ExcelFile — may hit corrupt styles).
    """
    if isinstance(sheet_name, int):
        return sheet_name
    req = (sheet_name or "").strip()
    if not req:
        return 0
    try:
        xf = pd.ExcelFile(path, engine="calamine")
    except Exception as ex:  # noqa: BLE001 — probe only
        _LOG.debug("ExcelFile(calamine) probe failed: %s", ex)
        return sheet_name
    names = list(xf.sheet_names)
    if req in names:
        return req
    nreq = _norm_sheet_key(req)
    n_matches = [n for n in names if _norm_sheet_key(n) == nreq]
    if len(n_matches) == 1:
        if n_matches[0] != req:
            _LOG.info(
                "Excel: sheet name matched after NFKC/strip: requested=%r -> using=%r",
                req,
                n_matches[0],
            )
        return n_matches[0]
    if len(n_matches) > 1:
        _LOG.warning(
            "Multiple sheets match %r after normalize; using %r. All: %s",
            req,
            n_matches[0],
            n_matches,
        )
        return n_matches[0]
    if len(names) == 1:
        _LOG.warning(
            "Sheet %r not in workbook; only one sheet %r — using it.",
            req,
            names[0],
        )
        return names[0]
    return sheet_name


def _read_excel_pandas_openpyxl(
    path: str, sheet_name: str | int, *, header: int = 0
) -> pd.DataFrame:
    """
    Fallback when calamine is unavailable or fails. Prefer openpyxl read_only first: it skips
    full stylesheet parse and avoids openpyxl IndexError on corrupt style XML. Full workbook
    parse is only a last resort (e.g. read_only limitations on some sheets).
    """
    try:
        return pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=header,
            engine="openpyxl",
            engine_kwargs={"read_only": True},
        )
    except Exception as err_ro:
        _LOG.warning(
            "pd.read_excel(openpyxl read_only) failed; retry full parse: path=%r sheet=%r err=%s",
            path,
            sheet_name,
            err_ro,
        )
        try:
            return pd.read_excel(
                path, sheet_name=sheet_name, header=header, engine="openpyxl"
            )
        except Exception as err_full:
            _LOG.error(
                "pd.read_excel(openpyxl full) also failed: path=%r sheet=%r err=%s",
                path,
                sheet_name,
                err_full,
            )
            raise


def _read_excel_tabular(
    path: str, sheet_name: str | int, *, header: int = 0
) -> pd.DataFrame:
    """
    Prefer calamine (Rust; ignores broken Excel styles, fast). Requires python-calamine
    and pandas >= 2.2. Resolve sheet names for calamine literal matching before read.
    On ImportError or read failure, fall back to openpyxl (same resolved sheet key).
    """
    resolved = _resolve_tabular_sheet_name_calamine(path, sheet_name)
    try:
        return pd.read_excel(
            path, sheet_name=resolved, header=header, engine="calamine"
        )
    except ImportError as err:
        _LOG.warning(
            "pd.read_excel(calamine) unavailable (%s); install python-calamine. Using openpyxl.",
            err,
        )
        return _read_excel_pandas_openpyxl(path, resolved, header=header)
    except Exception as err:
        _LOG.warning(
            "pd.read_excel(calamine) failed; fallback openpyxl: path=%r sheet=%r err=%s",
            path,
            resolved,
            err,
        )
        return _read_excel_pandas_openpyxl(path, resolved, header=header)


_DATE_FRAG_RE = re.compile(r"\d{1,2}/\d{1,2}")


def _should_merge_koubai_header_row(probe: pd.DataFrame, hdr: int) -> bool:
    """工程別ブック: 見出し行のひとつ上の行に M/D など日付断片があれば複合見出し対象。"""
    if hdr < 1 or probe is None or getattr(probe, "empty", True):
        return False
    prev = probe.iloc[hdr - 1]
    for v in prev:
        if pd.isna(v):
            continue
        if _DATE_FRAG_RE.search(str(v)):
            return True
    return False


def _combine_koubai_header_cells(label_cell, supplement_cell) -> str:
    """Excel 6行目相当 + 5行目相当（plan 工程別… と同じ順: ラベル & 上段）。"""
    a = "" if pd.isna(label_cell) else str(label_cell).strip()
    b = "" if pd.isna(supplement_cell) else str(supplement_cell).strip()
    return f"{a}{b}"


def _dedupe_column_names(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for i, raw in enumerate(names):
        base = raw.strip() if str(raw).strip() else f"Column{i}"
        if base not in seen:
            seen[base] = 0
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}.{seen[base]}")
    return out


def _read_excel_full_no_header(path: str, sheet_name: str | int) -> pd.DataFrame:
    resolved = _resolve_tabular_sheet_name_calamine(path, sheet_name)
    try:
        return pd.read_excel(path, sheet_name=resolved, header=None, engine="calamine")
    except ImportError as err:
        _LOG.warning(
            "pd.read_excel(calamine) unavailable (%s); openpyxl no-header read: path=%r",
            err,
            path,
        )
        try:
            return pd.read_excel(
                path,
                sheet_name=resolved,
                header=None,
                engine="openpyxl",
                engine_kwargs={"read_only": True},
            )
        except Exception as err_ro:
            _LOG.warning("openpyxl read_only no-header failed; full parse: %s", err_ro)
            return pd.read_excel(path, sheet_name=resolved, header=None, engine="openpyxl")
    except Exception as err:
        _LOG.warning(
            "pd.read_excel(calamine) no-header failed; fallback openpyxl: path=%r err=%s",
            path,
            err,
        )
        try:
            return pd.read_excel(
                path,
                sheet_name=resolved,
                header=None,
                engine="openpyxl",
                engine_kwargs={"read_only": True},
            )
        except Exception:
            return pd.read_excel(path, sheet_name=resolved, header=None, engine="openpyxl")


def _parse_cell_to_date(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        ts = pd.Timestamp(val)
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        pass
    try:
        ts = pd.to_datetime(str(val).strip(), errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        return None


def _base_year_month_from_order_date(df: pd.DataFrame) -> tuple[int, int]:
    """受注日列の先頭の有効な値から年・月（M/D 見出しの年付与に使用）。"""
    if df is None or getattr(df, "empty", True) or "受注日" not in df.columns:
        from datetime import date

        t = date.today()
        return t.year, t.month
    for val in df["受注日"]:
        d = _parse_cell_to_date(val)
        if d is not None:
            return d.year, d.month
    from datetime import date

    t = date.today()
    return t.year, t.month


def _normalize_koubai_date_header(name: str, base_year: int, base_month: int) -> str:
    """括弧前を対象に M/D または Y/M/D を YYYY/MM/DD に（曜日は列名から除去済み想定）。"""
    if not name or not isinstance(name, str):
        return name
    h = name.strip()
    core = h.split("(", 1)[0].strip() if "(" in h else h
    parts = [p for p in core.split("/") if p != ""]
    if len(parts) == 3:
        try:
            y, mo, d = int(parts[0]), int(parts[1]), int(parts[2])
            if y >= 1900:
                return f"{y:04d}/{mo:02d}/{d:02d}"
        except ValueError:
            pass
    if len(parts) == 2:
        try:
            mo, dy = int(parts[0]), int(parts[1])
        except ValueError:
            return h
        yy = base_year
        if base_month >= 11 and mo <= 2:
            yy = base_year + 1
        elif base_month <= 2 and mo >= 11:
            yy = base_year - 1
        else:
            yy = base_year
        return f"{yy:04d}/{mo:02d}/{dy:02d}"
    return h


def _shape_koubai_inquiry_dataframe(df_raw: pd.DataFrame, hdr: int) -> pd.DataFrame:
    """
    plan/工程別生産計画問合せ_成形.m と同趣旨:
    複合見出し、加工時間/加工速度列削除、加工数量の部分除去（見出しが「加工数量」のみの列は列名維持）、日付列名を YYYY/MM/DD。
    """
    if df_raw is None or getattr(df_raw, "empty", True):
        return df_raw
    nrows, ncols = df_raw.shape
    if hdr >= nrows or hdr < 1:
        return df_raw
    names: list[str] = []
    for j in range(ncols):
        lab = df_raw.iloc[hdr, j]
        sup = df_raw.iloc[hdr - 1, j]
        names.append(_combine_koubai_header_cells(lab, sup))
    names = _dedupe_column_names(names)
    body = df_raw.iloc[hdr + 1 :].copy()
    if body.shape[1] != len(names):
        names = names[: body.shape[1]]
    body.columns = names
    # ③ 加工時間・加工速度を含む列を削除
    keep = [
        c
        for c in body.columns
        if "加工時間" not in str(c) and "加工速度" not in str(c)
    ]
    body = body.loc[:, keep]
    # ④ 「加工数量」の部分文字列のみ除去。見出しが「加工数量」単独の列は列名を変えない
    def _strip_koubai_qty_token(col: object) -> str:
        s = str(col).strip() if col is not None else ""
        if s == "加工数量":
            return s
        return s.replace("加工数量", "")

    body = body.rename(columns=_strip_koubai_qty_token)
    by, bm = _base_year_month_from_order_date(body)
    rename_map = {
        c: _normalize_koubai_date_header(str(c), by, bm) for c in body.columns
    }
    body = body.rename(columns=rename_map)
    # 正規化で列名が衝突する場合は列を落とさず接尾辞で区別
    _final: list[str] = []
    _seen: dict[str, int] = {}
    for c in body.columns:
        s = str(c)
        if s not in _seen:
            _seen[s] = 0
            _final.append(s)
        else:
            _seen[s] += 1
            _final.append(f"{s}__{_seen[s]}")
    body.columns = _final
    return body


def _try_read_koubai_inquiry_shaped(
    path: str, sheet_arg: str | int, hdr: int
) -> pd.DataFrame | None:
    env_raw = (os.environ.get(ENV_KOUBAI_INQUIRY_SHAPING) or "").strip().lower()
    if env_raw in ("0", "false", "no", "off", "disable", "disabled"):
        return None
    probe = _probe_excel_no_header(path, sheet_arg, nrows=max(80, hdr + 5))
    force = env_raw in ("1", "true", "yes", "on", "force", "always")
    use_merge = False
    if hdr >= 1:
        if force:
            use_merge = True
        elif _should_merge_koubai_header_row(probe, hdr):
            use_merge = True
    if not use_merge:
        return None
    df_raw = _read_excel_full_no_header(path, sheet_arg)
    shaped = _shape_koubai_inquiry_dataframe(df_raw, hdr)
    _LOG.info(
        "Excel: 工程別問合せ風の複合見出し成形を適用しました（header Excel row=%s）。",
        hdr + 1,
    )
    return shaped


def read_tabular_dataframe(
    path: str,
    *,
    sheet_name: str | None = None,
    csv_encoding: str | None = None,
) -> pd.DataFrame:
    """Load CSV / Parquet / Excel by extension."""
    low = path.lower()
    if low.endswith(".csv"):
        enc = csv_encoding
        if enc is None:
            enc = (os.environ.get(ENV_TABULAR_CSV_ENCODING) or "").strip() or "utf-8-sig"
        return pd.read_csv(path, encoding=enc)
    if low.endswith((".parquet", ".pq")):
        return pd.read_parquet(path)
    sn: str | int = sheet_name if sheet_name else 0
    resolved = _resolve_tabular_sheet_name_calamine(path, sn)
    hdr = _resolve_tabular_excel_header_row_0based(path, resolved)
    shaped = _try_read_koubai_inquiry_shaped(path, sn, hdr)
    if shaped is not None:
        return shaped
    return _read_excel_tabular(path, sn, header=hdr)

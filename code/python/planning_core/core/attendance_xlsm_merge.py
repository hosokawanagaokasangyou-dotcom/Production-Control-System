# -*- coding: utf-8 -*-
"""Merge openpyxl-edited xlsm with original to preserve drawings/buttons on untouched sheets."""
from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path

_PRESERVE_PART_PREFIXES = (
    "xl/drawings/",
    "xl/media/",
    "xl/ctrlProps/",
    "xl/printerSettings/",
)


def _normalize_xl_zip_path(target: str) -> str:
    t = (target or "").strip().replace("\\", "/")
    if t.startswith("/"):
        t = t[1:]
    if not t.startswith("xl/"):
        t = "xl/" + t.lstrip("/")
    return t


def _sheet_name_to_worksheet_path(zf: zipfile.ZipFile) -> dict[str, str]:
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    id_to_target: dict[str, str] = {}
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels_xml):
        tag = m.group(0)
        id_m = re.search(r"Id=\"([^\"]*)\"", tag)
        target_m = re.search(r"Target=\"([^\"]*)\"", tag)
        if id_m and target_m:
            id_to_target[id_m.group(1)] = target_m.group(1)
    out: dict[str, str] = {}
    for m in re.finditer(r"<sheet\b[^>]*/>", wb_xml):
        tag = m.group(0)
        name_m = re.search(r"name=\"([^\"]*)\"", tag)
        id_m = re.search(r"r:id=\"([^\"]*)\"", tag)
        if not name_m or not id_m:
            continue
        target = id_to_target.get(id_m.group(1), "")
        if target:
            out[name_m.group(1)] = _normalize_xl_zip_path(target)
    return out


def _worksheet_rels_path(worksheet_path: str) -> str:
    return worksheet_path.replace("worksheets/", "worksheets/_rels/") + ".rels"


def _should_preserve_zip_entry(name: str) -> bool:
    if name.startswith(_PRESERVE_PART_PREFIXES):
        return True
    if name.startswith("xl/comments") and name.endswith(".xml"):
        return True
    return False


def merge_xlsm_preserving_unmodified_sheets(
    original_path: Path,
    edited_path: Path,
    output_path: Path,
    *,
    replaced_sheet_names: set[str],
) -> None:
    """
    openpyxl 保存後の xlsm に、未変更シートの worksheet XML と図形関連パーツを原本から復元する。

    replaced_sheet_names: openpyxl で新規作成・上書きしたシート名（APP_* 出力先）。
    """
    original_path = original_path.resolve()
    edited_path = edited_path.resolve()
    output_path = output_path.resolve()
    replaced = {n.strip() for n in replaced_sheet_names if n and n.strip()}

    with zipfile.ZipFile(original_path, "r") as z_orig, zipfile.ZipFile(edited_path, "r") as z_edit:
        orig_sheets = _sheet_name_to_worksheet_path(z_orig)
        edit_sheets = _sheet_name_to_worksheet_path(z_edit)
        preserved_names = [n for n in edit_sheets if n in orig_sheets and n not in replaced]

        preserve_worksheet_paths: set[str] = set()
        preserve_rels_paths: set[str] = set()
        for name in preserved_names:
            ws = orig_sheets[name]
            preserve_worksheet_paths.add(ws)
            rels = _worksheet_rels_path(ws)
            if rels in z_orig.namelist():
                preserve_rels_paths.add(rels)

        orig_entries = {info.filename: info for info in z_orig.infolist()}
        edit_entries = {info.filename: info for info in z_edit.infolist()}

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
            tmp_path = Path(tmp.name)

        try:
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
                written: set[str] = set()

                def write_bytes(name: str, data: bytes, info: zipfile.ZipInfo | None = None) -> None:
                    if name in written:
                        return
                    if info is not None:
                        z_out.writestr(info, data)
                    else:
                        z_out.writestr(name, data)
                    written.add(name)

                for name, info in edit_entries.items():
                    if name in preserve_worksheet_paths or name in preserve_rels_paths:
                        continue
                    if _should_preserve_zip_entry(name):
                        continue
                    write_bytes(name, z_edit.read(name), info)

                for ws in preserve_worksheet_paths:
                    if ws in orig_entries:
                        write_bytes(ws, z_orig.read(ws), orig_entries[ws])
                for rels in preserve_rels_paths:
                    if rels in orig_entries:
                        write_bytes(rels, z_orig.read(rels), orig_entries[rels])

                for name, info in orig_entries.items():
                    if _should_preserve_zip_entry(name):
                        write_bytes(name, z_orig.read(name), info)

            shutil.move(str(tmp_path), str(output_path))
        finally:
            if tmp_path.exists() and tmp_path.resolve() != output_path.resolve():
                tmp_path.unlink(missing_ok=True)

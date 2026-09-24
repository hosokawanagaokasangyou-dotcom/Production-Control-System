from __future__ import annotations

import re
import unicodedata
from pathlib import Path

from python_calamine import CalamineWorkbook

import common
from common import KONAN_SHEETS, month_range, ym_from_filename, ym_from_gatsudo

_NENDO_DIR_PAT = re.compile(r'^(\d{4})年度試算')


def list_kokubu_files(
    folders: Path | list[Path] | tuple[Path, ...] | None = None,
) -> dict[tuple[int, int], Path]:
    """国分②を年度フォルダ群から列挙。同一年月は更新日時が新しいファイルを採用。"""
    if folders is None:
        dirs = [Path(p) for p in common.KOKUBU_MEISAI_DIRS]
    elif isinstance(folders, Path):
        dirs = [folders]
    else:
        dirs = list(folders)

    best: dict[tuple[int, int], Path] = {}
    for d in dirs:
        if not d.is_dir():
            continue
        for f in d.glob('*後加工工賃明細*.xls*'):
            if f.name.startswith('~$'):
                continue
            ym = ym_from_filename(f)
            if not ym:
                continue
            cur = best.get(ym)
            if cur is None or f.stat().st_mtime > cur.stat().st_mtime:
                best[ym] = f
    return best


def pick_period(
    available: dict[tuple[int, int], Path],
    n: int,
    end: tuple[int, int] | None = None,
) -> tuple[list[tuple[int, int]], dict[tuple[int, int], Path]]:
    if not available and end is None:
        return [], {}
    end_ym = end or max(available)
    months = month_range(end_ym, n)
    files = {ym: available[ym] for ym in months if ym in available}
    return months, files


def shisan_ym(path: Path) -> tuple[int, int] | None:
    try:
        wb = CalamineWorkbook.from_path(str(path))
    except Exception:
        return None
    for name in KONAN_SHEETS:
        if name not in wb.sheet_names:
            continue
        rows = wb.get_sheet_by_name(name).to_python()
        for r in rows[:15]:
            for v in r:
                if isinstance(v, str):
                    ym = ym_from_gatsudo(v)
                    if ym:
                        return ym
    return None


def list_konan_files(root: Path) -> dict[tuple[int, int], Path]:
    best: dict[tuple[int, int], Path] = {}
    if not root.is_dir():
        return best
    cands: list[Path] = []
    for f in root.glob('*加工賃試算*.xls*'):
        if not f.name.startswith('~$'):
            cands.append(f)
    for d in root.iterdir():
        if not d.is_dir():
            continue
        name = d.name
        if name in ('temp', 'バックアップ') or 'バックアップ' in name:
            continue
        if not _NENDO_DIR_PAT.match(unicodedata.normalize('NFKC', name)):
            continue
        for f in d.glob('*加工賃試算*.xls*'):
            if not f.name.startswith('~$'):
                cands.append(f)
    for f in cands:
        ym = shisan_ym(f)
        if not ym:
            continue
        cur = best.get(ym)
        if cur is None or f.stat().st_mtime > cur.stat().st_mtime:
            best[ym] = f
    return best


def merge_available(*dicts: dict) -> dict:
    out: dict = {}
    for d in dicts:
        out.update(d)
    return out


def resolve_end_ym(*availables: dict) -> tuple[int, int] | None:
    yms = [ym for d in availables for ym in d]
    return max(yms) if yms else None

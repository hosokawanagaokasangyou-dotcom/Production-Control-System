from __future__ import annotations

import unicodedata
from collections import defaultdict
from pathlib import Path

from python_calamine import CalamineWorkbook

from common import (
    KOKUBU_MATOME,
    KOKUBU_SRC_SHEETS,
    TOTAL_LABELS,
    num,
    ym_from_filename,
)

# 0-based: H..X 工程名 / AB..AR 金額 / AT..BJ 数量（17列対応）
_COL_E = 4
_COL_G = 6
_COL_H = 7
_COL_X = 23
_COL_AA = 26
_COL_AB = 27
_COL_AT = 45
_KIND_COUNT = _COL_X - _COL_H + 1


def _cell(row, idx):
    return row[idx] if len(row) > idx else None


def _disp(v) -> str:
    """工程名表示用（NFKC、改行→空白）"""
    if v is None:
        return ''
    s = unicodedata.normalize('NFKC', str(v)).replace('\n', ' ').strip()
    return s


def _process_names(header_row) -> list[str | None]:
    names: list[str | None] = []
    for i in range(_KIND_COUNT):
        v = _cell(header_row, _COL_H + i)
        s = _disp(v)
        names.append(s if s and s != '合計' else None)
    return names


def _find_kind_header_row(rows) -> int | None:
    """加工依頼No.行の次行、または G列=合計 かつ Hに工程名がある行"""
    for i, row in enumerate(rows[:40]):
        a = _disp(_cell(row, 0))
        if '加工依頼' in a:
            if i + 1 < len(rows) and any(_process_names(rows[i + 1])):
                return i + 1
    for i, row in enumerate(rows[:40]):
        if _disp(_cell(row, _COL_G)) == '合計' and any(_process_names(row)):
            return i
    return None


def _is_data_row(row) -> bool:
    a = _cell(row, 0)
    b = _cell(row, 1)
    c = _cell(row, 2)
    c_s = _disp(c)
    if c_s in TOTAL_LABELS:
        return False
    # 依頼接頭辞または契約NOらしきもの
    a_s = _disp(a)
    if a_s and a_s not in ('0', '0.0') and not a_s.replace('.', '', 1).isdigit():
        return True
    if b is not None and str(b).strip() not in ('', '0', '0.0'):
        # 連番のみの行は契約とセットで見る
        if c_s and c_s not in TOTAL_LABELS and c_s not in ('0', '0.0'):
            return True
    if c_s and c_s not in TOTAL_LABELS and c_s not in ('0', '0.0'):
        if a_s or (b is not None and str(b).strip() not in ('', '0', '0.0')):
            return True
    return False


def read_kokubu_month(path: Path) -> dict:
    warnings: list[str] = []
    kinds: dict[str, dict[str, float]] = defaultdict(lambda: {'wage': 0.0, 'qty': 0.0})
    sheets: dict[str, dict[str, float]] = {
        s: {'wage': 0.0, 'qty': 0.0} for s in KOKUBU_SRC_SHEETS
    }
    ym = ym_from_filename(path)

    wb = CalamineWorkbook.from_path(str(path))
    if KOKUBU_MATOME not in wb.sheet_names:
        warnings.append(f'{path.name}: シート「{KOKUBU_MATOME}」がありません')
        return {
            'ym': ym,
            'path': path,
            'kinds': {},
            'sheets': sheets,
            'warnings': warnings,
        }

    matome = wb.get_sheet_by_name(KOKUBU_MATOME).to_python()
    hdr_i = _find_kind_header_row(matome)
    if hdr_i is None:
        warnings.append(f'{path.name}: 東レまとめに工程名ヘッダーが見つかりません')
        names: list[str | None] = []
        data_start = 0
    else:
        names = _process_names(matome[hdr_i])
        data_start = hdr_i + 1

    for row in matome[data_start:]:
        if not row or not _is_data_row(row):
            continue
        for i, name in enumerate(names):
            if not name:
                continue
            kinds[name]['wage'] += num(_cell(row, _COL_AB + i))
            kinds[name]['qty'] += num(_cell(row, _COL_AT + i))

    for sname in KOKUBU_SRC_SHEETS:
        if sname not in wb.sheet_names:
            warnings.append(f'{path.name}: シート「{sname}」がありません')
            continue
        rows = wb.get_sheet_by_name(sname).to_python()
        for row in rows:
            if not row or not _is_data_row(row):
                continue
            sheets[sname]['wage'] += num(_cell(row, _COL_AA))
            sheets[sname]['qty'] += num(_cell(row, _COL_E))

    return {
        'ym': ym,
        'path': path,
        'kinds': dict(kinds),
        'sheets': sheets,
        'warnings': warnings,
    }

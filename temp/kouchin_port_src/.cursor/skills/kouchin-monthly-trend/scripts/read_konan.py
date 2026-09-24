from __future__ import annotations

import unicodedata
from collections import defaultdict
from pathlib import Path

from python_calamine import CalamineWorkbook

from common import KONAN_SHEETS, num, ym_from_gatsudo


def _n(v) -> str:
    return unicodedata.normalize('NFKC', str(v or '')).strip()


def _disp(v) -> str:
    return _n(v).replace('\n', ' ').strip()


def _find_ym(rows) -> tuple[int, int] | None:
    for r in rows[:15]:
        for v in r:
            if isinstance(v, str):
                ym = ym_from_gatsudo(v)
                if ym:
                    return ym
    return None


def _find_header(rows) -> int | None:
    for i, r in enumerate(rows[:40]):
        if r and '加工依頼' in _n(r[0] if len(r) > 0 else ''):
            return i
    return None


def read_konan_month(path: Path) -> dict:
    warnings: list[str] = []
    kinds: dict[str, dict[str, float]] = defaultdict(lambda: {'wage': 0.0, 'qty': 0.0})
    sheets: dict[str, dict[str, float]] = {
        s: {'wage': 0.0, 'qty': 0.0} for s in KONAN_SHEETS
    }
    ym = None

    wb = CalamineWorkbook.from_path(str(path))
    found = 0
    for sname in KONAN_SHEETS:
        if sname not in wb.sheet_names:
            warnings.append(f'{path.name}: シート「{sname}」がありません')
            continue
        found += 1
        rows = wb.get_sheet_by_name(sname).to_python()
        if ym is None:
            ym = _find_ym(rows)
        hdr = _find_header(rows)
        if hdr is None:
            warnings.append(f'{path.name}「{sname}」: ヘッダー行が見つかりません')
            continue
        header = rows[hdr]
        name_row = rows[hdr + 1] if len(rows) > hdr + 1 else []
        amt_col = next(
            (j for j, v in enumerate(header)
             if '加工賃' in _n(v) and '加工種類別' in _n(v)),
            None,
        )
        qty_block = next(
            (j for j, v in enumerate(header)
             if '加工量' in _n(v) and '加工種類別' in _n(v)),
            None,
        )
        if amt_col is None:
            warnings.append(f'{path.name}「{sname}」: 加工種類別 加工賃列がありません')
            continue
        qcol = next((j for j, v in enumerate(header) if '入庫' in _n(v) and '数量' in _n(v)), None)
        if qcol is None:
            qcol = next((j for j, v in enumerate(header) if '入庫' in _n(v)), None)

        # 種類列: 合計列の右〜加工量ブロック手前（賃）、加工量合計の右（量）
        wage_kind_cols: list[tuple[int, str]] = []
        end_w = qty_block if qty_block is not None else len(name_row)
        for j in range(amt_col + 1, end_w):
            name = _disp(_cell(name_row, j) if name_row else None)
            if name and name != '合計':
                wage_kind_cols.append((j, name))

        qty_kind_cols: list[tuple[int, str]] = []
        if qty_block is not None:
            for j in range(qty_block + 1, max(len(name_row), len(header))):
                name = _disp(_cell(name_row, j) if name_row else None)
                if name and name != '合計':
                    qty_kind_cols.append((j, name))

        data_start = hdr + 4
        for row in rows[data_start:]:
            if not row:
                continue
            irai = _n(row[0]) if len(row) > 0 else ''
            keiyaku = _n(row[2]).replace('-', '') if len(row) > 2 else ''
            if not irai and not keiyaku:
                continue
            # 空白セルは金額0
            sheets[sname]['wage'] += num(_soft_num(row, amt_col))
            if qcol is not None:
                sheets[sname]['qty'] += num(_soft_num(row, qcol))
            for j, name in wage_kind_cols:
                kinds[name]['wage'] += num(_soft_num(row, j))
            for j, name in qty_kind_cols:
                kinds[name]['qty'] += num(_soft_num(row, j))

    if not found:
        warnings.append(f'{path.name}: 東レシートが1つもありません')

    return {
        'ym': ym,
        'path': path,
        'kinds': dict(kinds),
        'sheets': sheets,
        'warnings': warnings,
    }


def _cell(row, idx):
    return row[idx] if row is not None and len(row) > idx else None


def _soft_num(row, idx):
    v = _cell(row, idx)
    if isinstance(v, str) and not v.strip():
        return 0
    return v

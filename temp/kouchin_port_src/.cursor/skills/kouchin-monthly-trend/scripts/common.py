from __future__ import annotations

import re
import unicodedata
from pathlib import Path

KONAN_SHISAN_ROOT = (
    r'\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\002  加工G\000  後加工業務\2 後加工試算'
)
# 国分② 後加工工賃明細（年度フォルダを直接参照。.lnk は使わない）
KOKUBU_MEISAI_DIRS = (
    r'\\192.168.0.101\共有フォルダ\国分工場\国分管理\後加工\後加工工賃明細\工賃明細2026年度（令和8年度）',
    r'\\192.168.0.101\共有フォルダ\国分工場\国分管理\後加工\後加工工賃明細\工賃明細２０２５年度　(R7年度)',
)
GATSUDO_PAT = re.compile(r'(\d{4})\s*年\s*(\d{1,2})\s*月度')
KOKUBU_SRC_SHEETS = ('東レT', '東レV.C', '東レY', '東レW.E')
KOKUBU_MATOME = '東レまとめ'
KONAN_SHEETS = ('東レT.V.C', '東レY.S', '東レW.E..')
TOTAL_LABELS = frozenset({'加工量合計', '加工賃合計', '営業入庫量', '東レ合計', '総合計'})


def norm(v) -> str:
    if v is None:
        return ''
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return unicodedata.normalize('NFKC', str(v)).strip().upper()


def ym_from_gatsudo(text) -> tuple[int, int] | None:
    m = GATSUDO_PAT.search(unicodedata.normalize('NFKC', str(text or '')))
    return (int(m.group(1)), int(m.group(2))) if m else None


def ym_from_filename(path: Path) -> tuple[int, int] | None:
    return ym_from_gatsudo(path.name)


def add_months(ym: tuple[int, int], delta: int) -> tuple[int, int]:
    y, m = ym
    idx = y * 12 + (m - 1) + delta
    return idx // 12, idx % 12 + 1


def month_range(end: tuple[int, int], n: int) -> list[tuple[int, int]]:
    return [add_months(end, i - (n - 1)) for i in range(n)]


def resolve_base(script_file: str | Path | None = None) -> Path:
    """●自動検証ルート = .../.cursor/skills/kouchin-monthly-trend/scripts から 5 つ上 (parents[4])"""
    here = Path(script_file or __file__).resolve()
    return here.parents[4]


def num(v) -> float:
    return float(v) if isinstance(v, (int, float)) else 0.0

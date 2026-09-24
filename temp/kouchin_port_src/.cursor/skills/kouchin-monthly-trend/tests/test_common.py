import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))

from common import norm, ym_from_gatsudo, ym_from_filename, month_range, add_months


def test_norm_nfkc_upper():
    assert norm('Ｔ7') == 'T7'
    assert norm(72.0) == '72'
    assert norm(None) == ''


def test_ym_from_gatsudo():
    assert ym_from_gatsudo('2026年8月度') == (2026, 8)
    assert ym_from_gatsudo('対象 2026年 8 月度') == (2026, 8)
    assert ym_from_gatsudo('なし') is None


def test_ym_from_filename():
    p = Path('後加工工賃明細（2026年8月度)V01.xlsx')
    assert ym_from_filename(p) == (2026, 8)


def test_add_months_and_range():
    assert add_months((2026, 1), -1) == (2025, 12)
    assert month_range((2026, 8), 3) == [(2026, 6), (2026, 7), (2026, 8)]


def test_resolve_base_points_to_auto_verify_root():
    from common import resolve_base
    # .../●自動検証/.cursor/skills/kouchin-monthly-trend/scripts/common.py → parents[4]
    scripts = Path(__file__).resolve().parents[1] / 'scripts' / 'common.py'
    base = resolve_base(scripts)
    assert base.name == '●自動検証' or (base / '国分工場').exists() or (base / '.cursor').exists()
    assert (base / '.cursor' / 'skills' / 'kouchin-monthly-trend').is_dir()

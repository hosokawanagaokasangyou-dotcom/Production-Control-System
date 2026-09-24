import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))

from aggregate import build_cross_kind, build_grouped_cross, mom_share, shift_badge, rank_shift_suspects


def test_cross_kind_matrix():
    matrix = build_cross_kind(
        kokubu=[{'ym': (2026, 8), 'kinds': {'スリット': {'wage': 100, 'qty': 1}}}],
        konan=[{
            'ym': (2026, 8),
            'kinds': {
                'スリット': {'wage': 80, 'qty': 1},
                'カット': {'wage': 40, 'qty': 2},
            },
        }],
        months=[(2026, 8)],
        metric='wage',
    )
    assert matrix['processes'] == ['カット', 'スリット']
    assert matrix['values']['国分']['スリット'][(2026, 8)] == 100
    assert matrix['values']['湖南']['カット'][(2026, 8)] == 40


def test_mom_and_share():
    r = mom_share(
        kinds_by_ym={
            (2026, 7): {'A': {'wage': 100}},
            (2026, 8): {'A': {'wage': 120}, 'B': {'wage': 80}},
        },
        months=[(2026, 7), (2026, 8)],
        metric='wage',
    )
    assert r['mom']['A'] == 0.2
    assert abs(r['share']['A'] - 0.6) < 1e-9


def test_slice_aliases_group():
    """国分 スライス1+3 と 湖南 スライス を同一グループに合算"""
    g = build_grouped_cross(
        kokubu=[{
            'ym': (2026, 8),
            'kinds': {
                'スライス1': {'wage': 100, 'qty': 10},
                'スライス3': {'wage': 50, 'qty': 5},
            },
        }],
        konan=[{'ym': (2026, 8), 'kinds': {'スライス': {'wage': 80, 'qty': 8}}}],
        months=[(2026, 8)],
        metric='wage',
    )
    assert 'スライス' in g['processes']
    assert g['values']['国分']['スライス'][(2026, 8)] == 150
    assert g['values']['湖南']['スライス'][(2026, 8)] == 80


def test_shift_badge_strong_offset():
    assert shift_badge(100, -80) == '強い相殺'
    assert shift_badge(100, 50) == '同方向'
    assert shift_badge(100, None) is None


def test_rank_shift_suspects_orders_by_offset():
    rows = rank_shift_suspects(
        grouped={
            'processes': ['スライス', 'スリット'],
            'months': [(2026, 7), (2026, 8)],
            'values': {
                '国分': {
                    'スライス': {(2026, 7): 100, (2026, 8): 180},
                    'スリット': {(2026, 7): 50, (2026, 8): 55},
                },
                '湖南': {
                    'スライス': {(2026, 7): 200, (2026, 8): 120},
                    'スリット': {(2026, 7): 40, (2026, 8): 42},
                },
            },
        },
        metric_label='wage',
    )
    assert rows[0]['process'] == 'スライス'
    assert rows[0]['badge'] == '強い相殺'

import sys

from conftest import make_aladdin, make_nagaoka, write_toray_csv
import verify_kouchin as vk


def test_main_kokubu_smoke(tmp_path, monkeypatch):
    csv_dir = tmp_path / '東レ送付CSV'
    nagaoka_dir = tmp_path / '国分工場' / '長岡後加工賃明細'
    aladdin_dir = tmp_path / '国分工場' / 'アラジン'
    write_toray_csv(csv_dir / 'RVSHEET202608.csv', [
        ('A010', '191-352R', '260801', '', 1000),
    ], t_amount=1000)
    make_nagaoka(nagaoka_dir / '後加工工賃明細（2026年8月度)V01.xlsx', [
        ('T8', 1, '191352R', 1000),
    ])
    make_aladdin(aladdin_dir / '依頼NO別問合せ_20260831_235959.xlsx', '2026年08月', [
        ('T8-1', 1000),
    ])
    out = tmp_path / '検証結果_国分工場_20260831_120000.xlsx'
    monkeypatch.setattr(sys, 'argv', [
        'verify_kouchin.py', '--base', str(tmp_path), '--out', str(out),
    ])
    vk.main()
    assert out.exists()
    assert out.with_suffix('.html').exists()


def test_main_accepts_argv_without_sys_argv(tmp_path):
    csv_dir = tmp_path / '東レ送付CSV'
    nagaoka_dir = tmp_path / '国分工場' / '長岡後加工賃明細'
    aladdin_dir = tmp_path / '国分工場' / 'アラジン'
    write_toray_csv(csv_dir / 'RVSHEET202608.csv', [
        ('A010', '191-352R', '260801', '', 1000),
    ], t_amount=1000)
    make_nagaoka(nagaoka_dir / '後加工工賃明細（2026年8月度)V01.xlsx', [
        ('T8', 1, '191352R', 1000),
    ])
    make_aladdin(aladdin_dir / '依頼NO別問合せ_20260831_235959.xlsx', '2026年08月', [
        ('T8-1', 1000),
    ])
    out = tmp_path / 'out.xlsx'
    vk.main(['--base', str(tmp_path), '--out', str(out)])
    assert out.exists()

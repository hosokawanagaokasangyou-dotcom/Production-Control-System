import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))

from conftest import make_konan_shisan
from scan import list_kokubu_files, list_konan_files, pick_period


def test_list_kokubu_picks_latest_per_ym(tmp_path):
    d = tmp_path / '長岡'
    d.mkdir()
    (d / '後加工工賃明細（2026年7月度)V01.xlsx').write_bytes(b'PK')
    newer = d / '後加工工賃明細（2026年7月度)V02.xlsx'
    newer.write_bytes(b'PK')
    os.utime(newer, (time.time(), time.time() + 10))
    (d / '後加工工賃明細（2026年8月度)V01.xlsx').write_bytes(b'PK')
    (d / '~$後加工工賃明細（2026年8月度)V01.xlsx').write_bytes(b'PK')
    m = list_kokubu_files(d)
    assert set(m) == {(2026, 7), (2026, 8)}
    assert m[(2026, 7)].name.endswith('V02.xlsx')


def test_list_kokubu_merges_dirs_picks_newer(tmp_path):
    """複数年度フォルダを走査し、同月は更新日時が新しい方を採用"""
    d2025 = tmp_path / '工賃明細２０２５年度'
    d2026 = tmp_path / '工賃明細2026年度'
    d2025.mkdir()
    d2026.mkdir()
    older = d2025 / '後加工工賃明細（2026年2月度).xlsx'
    newer = d2026 / '後加工工賃明細（2026年2月度)V02.xlsx'
    older.write_bytes(b'PK')
    newer.write_bytes(b'PK')
    os.utime(older, (time.time(), time.time() - 100))
    os.utime(newer, (time.time(), time.time() + 10))
    (d2026 / '後加工工賃明細（2026年8月度)V01.xlsx').write_bytes(b'PK')
    m = list_kokubu_files([d2025, d2026])
    assert m[(2026, 2)] == newer
    assert (2026, 8) in m


def test_pick_period_end_is_max_ym():
    available = {(2026, 6): Path('a'), (2026, 8): Path('b'), (2026, 7): Path('c')}
    months, files = pick_period(available, n=3)
    assert months == [(2026, 6), (2026, 7), (2026, 8)]
    assert files[(2026, 8)] == Path('b')
    assert (2026, 5) not in files


def test_list_konan_files_from_nendo_and_root(tmp_path):
    root = tmp_path / '2 後加工試算'
    root.mkdir()
    make_konan_shisan(root / '月度加工賃試算.xlsx', '2026年9月度')
    nendo = root / '2026年度試算　湖南'
    make_konan_shisan(nendo / '8月度加工賃試算.xlsx', '2026年8月度')
    make_konan_shisan(nendo / '7月度加工賃試算.xlsx', '2026年7月度')
    (root / 'バックアップ').mkdir()
    make_konan_shisan(root / 'バックアップ' / 'x.xlsx', '2026年6月度')
    m = list_konan_files(root)
    assert set(m) == {(2026, 7), (2026, 8), (2026, 9)}

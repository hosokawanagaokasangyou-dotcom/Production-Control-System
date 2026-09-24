from conftest import make_nagaoka, make_aladdin
import verify_kouchin as vk


def test_read_nagaoka_aggregates_and_skips_memo(tmp_path):
    path = tmp_path / '後加工工賃明細（2026年8月度)V01.xlsx'
    make_nagaoka(path, [
        ('T8', 1, '191352R', 1000),
        ('T8', 1, '191352R', 200),
        ('T8', 2, '9月', 50),
    ])
    by_irai, by_keiyaku, keiyaku_to_irai, bad, warnings = vk.read_nagaoka(path)
    assert by_irai['T8-1'] == 1200
    assert by_keiyaku['191352R'] == 1200
    assert keiyaku_to_irai['191352R'] == {'T8-1'}
    assert bad == [(8, '9月', 50)]
    assert warnings == []


def test_read_aladdin_filters_customer_when_given(tmp_path):
    path = tmp_path / '依頼NO別問合せ_20260831_235959.xlsx'
    make_aladdin(path, '2026年08月', [
        ('T8-1', 1000, '049006'),
        ('C8-9', 777, '049052'),
    ], with_customer=True)
    all_rows, taisho, ym = vk.read_aladdin(path)
    assert all_rows['T8-1'] == 1000
    assert all_rows['C8-9'] == 777
    assert ym == (2026, 8)
    filtered, _, _ = vk.read_aladdin(path, customer='049006')
    assert filtered == {'T8-1': 1000}

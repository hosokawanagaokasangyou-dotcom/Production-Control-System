from pathlib import Path

import verify_kouchin as vk


def test_missing_manual_csv_is_empty(tmp_path):
    result, warnings = vk.load_manual_judgments(tmp_path / '手動判定.csv', '国分工場', (2026, 8))
    assert result == {}
    assert warnings == []


def test_manual_applies_only_matching_month_and_factory(tmp_path):
    p = tmp_path / '手動判定.csv'
    p.write_text(
        '工場,対象月,契約NO,正とする側,理由\n'
        '国分工場,2026年8月度,189759V,①,1円差\n'
        '国分工場,2026年7月度,188000A,②,他月\n'
        '湖南工場,2026年8月度,189111B,②,他工場\n',
        encoding='utf-8-sig',
    )
    result, warnings = vk.load_manual_judgments(p, '国分工場', (2026, 8))
    assert warnings == []
    assert result == {'189759V': (1, '1円差')}


def test_missing_prior_csv_is_empty(tmp_path):
    result, warnings = vk.load_prior_shortfalls(tmp_path / '前月過不足.csv', '国分工場', (2026, 8))
    assert result == []
    assert warnings == []

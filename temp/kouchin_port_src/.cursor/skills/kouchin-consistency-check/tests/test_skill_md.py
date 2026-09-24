from conftest import SKILL_MD, SCRIPT

SCRIPT_REL = '.cursor/skills/kouchin-consistency-check/scripts/verify_kouchin.py'


def _skill_text():
    return SKILL_MD.read_text(encoding='utf-8')


def test_skill_documents_actual_script_path():
    text = _skill_text()
    assert SCRIPT_REL in text
    assert 'python scripts/verify_kouchin.py' not in text
    assert SCRIPT.is_file()


def test_skill_uses_factories_function_not_constant():
    text = _skill_text()
    assert '_factories()' in text
    assert 'FACTORIES' not in text


def test_skill_mail_matches_peer_import_not_placeholder():
    text = _skill_text()
    assert '別途ご確認ください' not in text
    assert '0埋め' in text
    assert '自工場の行のみ' in text


def test_skill_html_saved_for_both_factories():
    text = _skill_text()
    assert '検証結果_<工場>_*.html' in text or '検証結果_<工場>_' in text and '.html' in text


def test_skill_prev_only_replaces_past_months():
    text = _skill_text()
    assert '翌月' in text
    assert '自動検出を無効化し1ファイルのみ' not in text


def test_skill_gray_fill_is_keishiki_fusei():
    text = _skill_text()
    assert '灰=対象外' not in text
    assert '対象外=E7E6E6' not in text
    assert '形式不正=E7E6E6' in text or '形式不正' in text and 'E7E6E6' in text


def test_skill_aladdin_does_not_claim_warehouse_filter():
    text = _skill_text()
    assert '倉庫511101' not in text
    assert '倉庫520201' not in text

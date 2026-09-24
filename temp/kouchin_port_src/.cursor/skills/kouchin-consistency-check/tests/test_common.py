from pathlib import Path

import verify_kouchin as vk


def test_norm_nfkc_upper_and_int():
    assert vk.norm('Ｔ7') == 'T7'
    assert vk.norm(72.0) == '72'
    assert vk.norm(None) == ''
    assert vk.norm('191-352R').replace('-', '') == '191352R'


def test_parse_ym_formats():
    assert vk._parse_ym('2026年8月度') == (2026, 8)
    assert vk._parse_ym('2026-08') == (2026, 8)
    assert vk._parse_ym('2026/8') == (2026, 8)
    assert vk._parse_ym('202608') == (2026, 8)
    assert vk._parse_ym('読めない') is None


def test_default_base_is_auto_verify_folder():
    assert vk.DEFAULT_BASE.name == '●自動検証'
    assert (vk.DEFAULT_BASE / '.cursor' / 'skills' / 'kouchin-consistency-check').is_dir()


def test_factories_not_module_constant():
    assert not hasattr(vk, 'FACTORIES')
    fac = vk._factories()
    assert fac['kokubu']['basho'] == 'A010'
    assert fac['kokubu']['check_d'] is vk.check_matome
    assert fac['kokubu']['customer3'] is None
    assert fac['konan']['basho'] == 'A010P'
    assert fac['konan']['check_d'] is None
    assert fac['konan']['customer3'] == '049006'


def test_gray_fill_is_keishiki_fusei():
    fill = vk._judge_fill('形式不正')
    assert fill.fgColor.rgb[-6:] == 'E7E6E6'
    assert vk._judge_fill('対象外') is None


def test_script_source_has_no_mail_placeholder():
    src = Path(vk.__file__).read_text(encoding='utf-8')
    assert '別途ご確認ください' not in src

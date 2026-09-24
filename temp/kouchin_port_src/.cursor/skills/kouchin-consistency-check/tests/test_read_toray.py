from conftest import write_toray_csv
import verify_kouchin as vk


def test_read_toray_filters_basho_and_flips_minus(tmp_path):
    path = tmp_path / 'RVSHEET202608.csv'
    write_toray_csv(path, [
        ('A010', '191-352R', '260801', '', 1000),
        ('A010', '191-352R', '260802', '-', 1000),
        ('A010P', '192-100A', '260803', '', 500),
    ], t_amount=500)
    result, by_basho, dates, warnings, ct_errors, minus = vk.read_toray_csv(path, 'A010')
    assert result['191352R'] == 0
    assert '192100A' not in result
    assert by_basho['A010'] == 0
    assert by_basho['A010P'] == 500
    assert minus['191352R'][0][2] == -1000
    assert warnings == []
    assert ct_errors == []

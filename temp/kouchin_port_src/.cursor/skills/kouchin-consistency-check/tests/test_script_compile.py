import warnings

from conftest import SCRIPT


def test_script_compiles_without_syntax_warning():
    src = SCRIPT.read_text(encoding='utf-8')
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter('always', SyntaxWarning)
        compile(src, str(SCRIPT), 'exec')
    syntax = [w for w in caught if issubclass(w.category, SyntaxWarning)]
    assert syntax == [], [f'{w.message} ({w.filename}:{w.lineno})' for w in syntax]

# -*- coding: utf-8 -*-
"""
工程管理AI テストコード用の環境セットアップ（単独実行可）。

  py -3 setup_environment.py

- pip を更新し、同フォルダの requirements.txt から依存をインストール（無い場合は既定リスト）
- Windows: xlwings の Excel アドインを配置

VBA の「環境構築」マクロからも同じスクリプトを呼び出します。
"""
from __future__ import annotations

import os
import subprocess
import sys
import sysconfig
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
_REQUIREMENTS = _ROOT / "requirements.txt"
# requirements.txt が無い環境向け（planning_core / マクロと揃える）
_FALLBACK_PKGS = [
    "pandas>=2.0",
    "openpyxl>=3.1",
    "xlwings>=0.30",
    "google-genai>=1.0",
    "cryptography>=42.0",
]


def _run(cmd: list[str], cwd: Path | None = None) -> int:
    print("+", " ".join(cmd), flush=True)
    r = subprocess.run(cmd, cwd=str(cwd) if cwd else None, check=False)
    return int(r.returncode)


def _xlwings_addin_install() -> int:
    if sys.platform != "win32":
        print("（Windows 以外では xlwings アドインはスキップ）", flush=True)
        return 0
    scripts = Path(sysconfig.get_path("scripts"))
    xw = scripts / "xlwings.exe"
    if not xw.is_file():
        print(f"xlwings.exe が見つかりません: {xw}", file=sys.stderr, flush=True)
        return 94
    return _run([str(xw), "addin", "install"])


def main() -> int:
    os.chdir(_ROOT)
    py = sys.executable

    code = _run([py, "-m", "pip", "install", "--upgrade", "pip"])
    if code != 0:
        print("pip の更新に失敗しました。", file=sys.stderr, flush=True)
        return code

    if _REQUIREMENTS.is_file():
        code = _run(
            [py, "-m", "pip", "install", "--upgrade", "-r", str(_REQUIREMENTS)],
            cwd=_ROOT,
        )
    else:
        print(
            f"requirements.txt が無いため既定パッケージを入れます: {_REQUIREMENTS}",
            flush=True,
        )
        code = _run(
            [py, "-m", "pip", "install", "--upgrade"] + _FALLBACK_PKGS,
            cwd=_ROOT,
        )
    if code != 0:
        print("依存パッケージのインストールに失敗しました。", file=sys.stderr, flush=True)
        return code

    code = _xlwings_addin_install()
    if code != 0:
        print("xlwings addin install に失敗しました。", file=sys.stderr, flush=True)
        return code

    print("環境セットアップが完了しました。", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# -*- coding: utf-8 -*-
"""
工程管理AI テストコード用の環境セットアップ（単独実行可）。

  py -3 setup_environment.py

- pip を更新し、同フォルダの requirements.txt から依存をインストール（無い場合は既定リスト）
- Windows: xlwings の Excel アドインを配置

VBA の「環境構築」マクロはブック直下から python\\setup_environment.py を実行します。
"""
from __future__ import annotations

import os
import re
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


def _merge_xlwings_user_conf_show_console() -> None:
    """
    %USERPROFILE%\\.xlwings\\xlwings.conf に SHOW CONSOLE を True でマージする。
    xlwings アドインの GetConfig / Python の read_user_config と同じ「"KEY","VALUE"」1行形式。
    """
    if sys.platform != "win32":
        return
    conf_path = Path.home() / ".xlwings" / "xlwings.conf"
    key_keep = "SHOW CONSOLE"
    key_upper = key_keep.upper()
    rows: list[tuple[str, str]] = []
    if conf_path.is_file():
        text = conf_path.read_text(encoding="utf-8", errors="replace")
        for line in text.splitlines():
            parts = re.findall(r'"[^"]*"', line)
            if len(parts) >= 2:
                k = parts[0].strip('"')
                v = parts[1].strip('"')
                if k.upper() == key_upper:
                    continue
                rows.append((k, v))
    rows.append((key_keep, "True"))
    conf_path.parent.mkdir(parents=True, exist_ok=True)
    data = "".join(f'"{k}","{v}"\n' for k, v in rows)
    conf_path.write_text(data, encoding="utf-8", newline="\n")
    print(f"xlwings: ユーザー設定に SHOW CONSOLE=True を書き込みました: {conf_path}", flush=True)


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

    try:
        _merge_xlwings_user_conf_show_console()
    except OSError as ex:
        print(f"xlwings.conf の更新をスキップしました: {ex}", flush=True)

    print("環境セットアップが完了しました。", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

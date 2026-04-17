# -*- coding: utf-8 -*-
"""
工程管理AI テストコード用の環境セットアップ（単独実行可）。

  py -3 setup_environment.py

- pip を更新し、同フォルダの requirements.txt から依存をインストール（無い場合は既定リスト）
- Windows: xlwings の Excel アドインを配置

VBA の「環境構築」マクロはブック直下から python\\setup_environment.py を実行します。

cmd / PowerShell 経由で失敗したときウィンドウが閉じないよう、非 0 終了時は
workbook_env_bootstrap.pause_cmd_window_on_cli_error（既定で pause）を挟む。
無効化: 環境変数 PM_AI_CMD_PAUSE_ON_ERROR=0 / false / no / off。
"""
from __future__ import annotations

import importlib.util
import os
import re
import shutil
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


def _run_capture_stdout_stderr(cmd: list[str], cwd: Path | None = None) -> int:
    """標準出力・標準エラーをコンソールに流さず終了コードだけ返す（xlwings.exe の長い Traceback 抑止用）。"""
    print("+", " ".join(cmd), flush=True)
    r = subprocess.run(
        cmd,
        cwd=str(cwd) if cwd else None,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return int(r.returncode)


def _excel_instances_visible_to_xlwings() -> bool:
    """マクロ実行中など、COM 上で Excel が既に動いているか（おおよそ）。"""
    try:
        import xlwings as xw

        return len(xw.apps) > 0
    except Exception:
        return False


def _xlwings_xlam_source_path() -> Path | None:
    spec = importlib.util.find_spec("xlwings")
    if spec is None or not spec.origin:
        return None
    p = Path(spec.origin).resolve().parent / "addin" / "xlwings.xlam"
    return p if p.is_file() else None


def _xlwings_addin_install_fallback_copy() -> int:
    """
    Excel 起動中は ``xlwings addin install`` がアドインの Close で COM エラーになることがある。
    Close を行わず、パッケージ同梱の xlwings.xlam をユーザ XLSTART にコピーする。
    """
    try:
        from xlwings.cli import get_addin_dir
    except ImportError as ex:
        print(f"xlwings の読み込みに失敗しました: {ex}", file=sys.stderr, flush=True)
        return 94

    src = _xlwings_xlam_source_path()
    if src is None:
        print("パッケージ内の xlwings.xlam が見つかりません。", file=sys.stderr, flush=True)
        return 94
    try:
        dest_dir = Path(get_addin_dir(False))
    except Exception as ex:
        print(f"XLSTART フォルダの取得に失敗しました: {ex}", file=sys.stderr, flush=True)
        return 94
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / "xlwings.xlam"
    try:
        shutil.copyfile(src, dest)
    except OSError as ex:
        errno = getattr(ex, "errno", None)
        winerror = getattr(ex, "winerror", None)
        # Excel が xlwings.xlam を掴んでいると上書きできない。
        # Windows では errno 13 だけでなく WinError 32（他プロセスが使用中）などになる。
        locked = errno == 13 or winerror in (32, 33)
        if locked and dest.is_file():
            locked_size = dest.stat().st_size
            if locked_size > 0:
                print(
                    "xlwings アドイン（xlwings.xlam）は既に XLSTART にありますが、"
                    "Excel がファイルを使用中のため上書きできませんでした。\n"
                    "pip と Python パッケージの更新は完了しています。\n"
                    "アドイン本体を最新ファイルに差し替えるには、すべての Excel を終了してから\n"
                    "「環境構築」を再実行するか、`xlwings addin install` を実行してください。",
                    flush=True,
                )
                return 0
        print(
            "xlwings アドインのコピーに失敗しました（ファイルが使用中の可能性があります）。\n"
            "対処: すべての Excel を終了してから「環境構築」を再実行するか、\n"
            "コマンドプロンプトで次を実行してください。\n"
            f"  xlwings addin install\n\n詳細: {ex}",
            file=sys.stderr,
            flush=True,
        )
        return 93
    print(
        "xlwings.xlam をユーザーの XLSTART にコピーしました。\n"
        "※ 実行中の Excel がアドインを閉じられなかったため、公式の addin install の代わりにコピーで配置しました。\n"
        "反映を確実にするには、Excel をすべて終了してから起動し直してください。",
        flush=True,
    )
    return 0


def _xlwings_addin_install() -> int:
    if sys.platform != "win32":
        print("（Windows 以外では xlwings アドインはスキップ）", flush=True)
        return 0
    scripts = Path(sysconfig.get_path("scripts"))
    xw_exe = scripts / "xlwings.exe"
    if not xw_exe.is_file():
        print(f"xlwings.exe が見つかりません: {xw_exe}", file=sys.stderr, flush=True)
        return 94
    # Excel 起動中に xlwings.exe addin install を走らせると Close で COM 例外となり、
    # 子プロセスが Traceback を大量に出すだけなので避ける。
    if _excel_instances_visible_to_xlwings():
        print(
            "Excel が起動中のため、xlwings.exe addin install はスキップし、"
            "XLSTART へ xlwings.xlam をコピーする手順に進みます。",
            flush=True,
        )
        return _xlwings_addin_install_fallback_copy()
    code = _run_capture_stdout_stderr([str(xw_exe), "addin", "install"])
    if code == 0:
        return 0
    print(
        "xlwings addin install が失敗したため、XLSTART への xlwings.xlam 直接コピーを試します。",
        flush=True,
    )
    return _xlwings_addin_install_fallback_copy()


def _strip_show_console_from_xlwings_user_conf() -> None:
    """
    以前の環境構築で書いた SHOW CONSOLE を %USERPROFILE%\\.xlwings\\xlwings.conf から外す。
    段階1/2 は cmd.exe 経由を既定とする（xlwings のコンソールを自動では有効にしない）。
    """
    if sys.platform != "win32":
        return
    conf_path = Path.home() / ".xlwings" / "xlwings.conf"
    if not conf_path.is_file():
        return
    key_upper = "SHOW CONSOLE".upper()
    rows: list[tuple[str, str]] = []
    text = conf_path.read_text(encoding="utf-8", errors="replace")
    for line in text.splitlines():
        parts = re.findall(r'"[^"]*"', line)
        if len(parts) >= 2:
            k = parts[0].strip('"')
            v = parts[1].strip('"')
            if k.upper() == key_upper:
                continue
            rows.append((k, v))
    try:
        if not rows:
            conf_path.unlink()
            print(
                f"xlwings: SHOW CONSOLE を外し、他キーが無いため {conf_path} を削除しました。",
                flush=True,
            )
        else:
            data = "".join(f'"{k}","{v}"\n' for k, v in rows)
            conf_path.write_text(data, encoding="utf-8", newline="\n")
            print(
                f"xlwings: ユーザー設定から SHOW CONSOLE を削除しました: {conf_path}",
                flush=True,
            )
    except OSError as ex:
        print(f"xlwings.conf の整理をスキップしました: {ex}", flush=True)


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

    _strip_show_console_from_xlwings_user_conf()

    print("環境セットアップが完了しました。", flush=True)
    return 0


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    raise SystemExit(_wbe_exit.run_cli_with_optional_pause_on_error(main))

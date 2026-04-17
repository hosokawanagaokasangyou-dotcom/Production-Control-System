# -*- coding: utf-8 -*-
"""
工程管理AI テストコード用の環境セットアップ（単独実行可）。

  ブック所在フォルダで:
    py -3 -X utf8 -u python\\setup_environment.py
  本ファイルと同じフォルダで:
    py -3 -X utf8 -u setup_environment.py

- pip を更新し、本スクリプトと同じフォルダの requirements.txt から依存をインストール（無い場合は既定リスト）
- Windows: xlwings の Excel アドインを配置

VBA の「環境構築」マクロはブック直下から python\\setup_environment.py を実行します。

cmd / PowerShell 経由で失敗したときウィンドウが閉じないよう、非 0 終了時は
workbook_env_bootstrap.pause_cmd_window_on_cli_error（既定で pause）を挟む。
無効化: 環境変数 PM_AI_CMD_PAUSE_ON_ERROR=0 / false / no / off、またはマクロブックの
「設定_環境変数」シートで A 列=PM_AI_CMD_PAUSE_ON_ERROR・B 列に同様の値（マクロから
TASK_INPUT_WORKBOOK が渡され、pip で依存インストール成功後にシートを読み込みます）。
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

# 本ファイルのあるフォルダ（通常は python\\）。os.chdir に依存しない。
SCRIPT_PATH = Path(__file__).resolve()
PYTHON_DIR = SCRIPT_PATH.parent
REQ_FILE = PYTHON_DIR / "requirements.txt"

_FALLBACK_PKGS = [
    "pandas>=2.0",
    "openpyxl>=3.1",
    "xlwings>=0.30",
    "google-genai>=1.0",
    "cryptography>=42.0",
]


def _log(msg: str) -> None:
    print(msg, flush=True)


def _apply_workbook_env_from_macro() -> None:
    """
    TASK_INPUT_WORKBOOK がプロセスにあるとき、マクロブックの「設定_環境変数」を os.environ に反映する。
    （openpyxl が import できる前提。pip install 成功後に呼ぶ。）
    """
    try:
        from workbook_env_bootstrap import apply_from_task_input_workbook
    except ImportError:
        return
    apply_from_task_input_workbook()


def _child_env() -> dict[str, str]:
    env = dict(os.environ)
    if sys.platform == "win32":
        env.setdefault("PYTHONUTF8", "1")
        env.setdefault("PYTHONIOENCODING", "utf-8")
    return env


def _run_streaming(cmd: list[str], *, cwd: Path | None) -> int:
    _log("+ " + " ".join(cmd))
    r = subprocess.run(
        cmd,
        cwd=str(cwd) if cwd else None,
        check=False,
        env=_child_env(),
    )
    return int(r.returncode)


def _run_capture(cmd: list[str], *, cwd: Path | None) -> tuple[int, str]:
    _log("+ " + " ".join(cmd))
    r = subprocess.run(
        cmd,
        cwd=str(cwd) if cwd else None,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=_child_env(),
    )
    blob = ((r.stdout or "") + "\n" + (r.stderr or "")).strip()
    return int(r.returncode), blob


def _is_excel_exe_running_win32() -> bool | None:
    """
    tasklist で EXCEL.EXE を検出。True=起動中、False=未起動、None=判定不能。
    None のときは保守的に「addin install を避けコピーへ」の経路に寄せる。
    """
    if sys.platform != "win32":
        return False
    try:
        r = subprocess.run(
            [
                "tasklist",
                "/FI",
                "IMAGENAME eq EXCEL.EXE",
                "/FO",
                "CSV",
                "/NH",
            ],
            capture_output=True,
            text=True,
            encoding=sys.getfilesystemencoding() or "utf-8",
            errors="replace",
            timeout=60,
            env=_child_env(),
        )
    except (OSError, subprocess.TimeoutExpired):
        return None
    out = r.stdout or ""
    for line in out.splitlines():
        cell0 = line.strip().strip('"').split(",")[0].strip('"')
        if cell0.upper() == "EXCEL.EXE":
            return True
    return False


def _xlwings_cli_exe() -> Path | None:
    scripts = Path(sysconfig.get_path("scripts"))
    for name in ("xlwings.exe", "xlwings.cmd"):
        p = scripts / name
        if p.is_file():
            return p
    return None


def _xlwings_xlam_source_path() -> Path | None:
    spec = importlib.util.find_spec("xlwings")
    if spec is None or not spec.origin:
        return None
    p = Path(spec.origin).resolve().parent / "addin" / "xlwings.xlam"
    return p if p.is_file() else None


def _user_xlstart_dir_win32() -> Path | None:
    """pywin32 無しでも使えるユーザ XLSTART（Excel 既定の起動フォルダ）。"""
    if sys.platform != "win32":
        return None
    appdata = (os.environ.get("APPDATA") or "").strip()
    if not appdata:
        return None
    return Path(appdata) / "Microsoft" / "Excel" / "XLSTART"


def _xlwings_addin_install_fallback_copy() -> int:
    try:
        from xlwings.cli import get_addin_dir
    except ImportError as ex:
        print(f"xlwings の読み込みに失敗しました: {ex}", file=sys.stderr, flush=True)
        return 94

    src = _xlwings_xlam_source_path()
    if src is None:
        print("パッケージ内の xlwings.xlam が見つかりません。", file=sys.stderr, flush=True)
        return 94
    dest_dir: Path | None = None
    try:
        dest_dir = Path(get_addin_dir(False))
    except Exception as ex:
        alt = _user_xlstart_dir_win32()
        if alt is not None:
            dest_dir = alt
            _log(
                "※ xlwings の get_addin_dir が使えないため、"
                f"ユーザ XLSTART（%APPDATA%）へコピーします: {dest_dir}"
            )
        else:
            print(
                f"XLSTART フォルダの取得に失敗しました: {ex}\n"
                "APPDATA からの代替パスも取得できませんでした。pywin32 を入れるか、"
                "手動で xlwings.xlam を Excel の XLSTART に配置してください。",
                file=sys.stderr,
                flush=True,
            )
            return 94
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / "xlwings.xlam"
    try:
        shutil.copyfile(src, dest)
    except OSError as ex:
        errno = getattr(ex, "errno", None)
        winerror = getattr(ex, "winerror", None)
        locked = errno == 13 or winerror in (32, 33)
        if locked and dest.is_file() and dest.stat().st_size > 0:
            print(
                "xlwings アドイン（xlwings.xlam）は既に XLSTART にありますが、\n"
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
            "  xlwings addin install\n\n"
            f"詳細: {ex}",
            file=sys.stderr,
            flush=True,
        )
        return 93
    print(
        "xlwings.xlam をユーザーの XLSTART にコピーしました。\n"
        "※ 実行中の Excel がアドインを閉じられなかったため、公式の addin install の代わりにコピーで配置した場合があります。\n"
        "反映を確実にするには、Excel をすべて終了してから起動し直してください。",
        flush=True,
    )
    return 0


def _xlwings_addin_install() -> int:
    if sys.platform != "win32":
        print("（Windows 以外では xlwings アドインはスキップ）", flush=True)
        return 0
    xw_exe = _xlwings_cli_exe()
    if xw_exe is None:
        scripts = Path(sysconfig.get_path("scripts"))
        print(
            f"xlwings の CLI が見つかりません（検索先: {scripts}）。\n"
            "直前の pip install が、この Python に xlwings を入れたか確認してください。",
            file=sys.stderr,
            flush=True,
        )
        return 94

    running = _is_excel_exe_running_win32()
    if running is True:
        print(
            "EXCEL.EXE が起動中のため、xlwings.exe addin install はスキップし、\n"
            "XLSTART へ xlwings.xlam をコピーする手順に進みます。",
            flush=True,
        )
        return _xlwings_addin_install_fallback_copy()
    if running is None:
        print(
            "Excel の起動有無を tasklist で判定できませんでした。\n"
            "COM トラブルを避けるため、XLSTART への直接コピーを試みます。",
            flush=True,
        )
        return _xlwings_addin_install_fallback_copy()

    code, blob = _run_capture([str(xw_exe), "addin", "install"], cwd=PYTHON_DIR)
    if code == 0:
        return 0
    print(
        f"xlwings addin install が終了コード {code} で失敗したため、XLSTART への直接コピーを試します。",
        flush=True,
    )
    if blob:
        tail = "\n".join(blob.splitlines()[-40:])
        print(
            "--- xlwings CLI 出力（末尾） ---\n" + tail + "\n--- end ---",
            file=sys.stderr,
            flush=True,
        )
    return _xlwings_addin_install_fallback_copy()


def _strip_show_console_from_xlwings_user_conf() -> None:
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
    py = Path(sys.executable)

    _log("==== 工程管理AI 環境セットアップ ====")
    _log(f"Python: {py}")
    _log(f"スクリプト: {SCRIPT_PATH}")
    _log(f"pip の cwd: {PYTHON_DIR}")
    _log(f"requirements: {REQ_FILE} （存在: {'はい' if REQ_FILE.is_file() else 'いいえ'}）")
    _log("")

    code = _run_streaming([str(py), "-m", "pip", "install", "--upgrade", "pip"], cwd=PYTHON_DIR)
    if code != 0:
        print("pip の更新に失敗しました。上記ログを確認するか、手動で次を実行してください。", file=sys.stderr, flush=True)
        print(f"  {py} -m pip install --upgrade pip", file=sys.stderr, flush=True)
        return code

    if REQ_FILE.is_file():
        code, pip_blob = _run_capture(
            [str(py), "-m", "pip", "install", "--upgrade", "-r", str(REQ_FILE)],
            cwd=PYTHON_DIR,
        )
    else:
        _log(f"requirements.txt が無いため既定パッケージを入れます: {REQ_FILE}")
        code, pip_blob = _run_capture(
            [str(py), "-m", "pip", "install", "--upgrade"] + _FALLBACK_PKGS,
            cwd=PYTHON_DIR,
        )
    if code != 0:
        print("依存パッケージのインストールに失敗しました。", file=sys.stderr, flush=True)
        if pip_blob:
            print("--- pip 出力（末尾 80 行） ---", file=sys.stderr, flush=True)
            print("\n".join(pip_blob.splitlines()[-80:]), file=sys.stderr, flush=True)
        return code

    _apply_workbook_env_from_macro()

    code = _xlwings_addin_install()
    if code != 0:
        print("xlwings アドイン手順で失敗しました。", file=sys.stderr, flush=True)
        return code

    _strip_show_console_from_xlwings_user_conf()

    print("環境セットアップが完了しました。", flush=True)
    return 0


def _pause_cli_error_standalone(exit_code: int | None) -> None:
    """workbook_env_bootstrap 不在時。pause_cmd_window_on_cli_error と同趣旨（PM_AI_CMD_PAUSE_ON_ERROR は os.environ のみ）。"""
    if os.name != "nt":
        return
    try:
        code = int(exit_code)
    except (TypeError, ValueError):
        code = 1
    if code == 0:
        return
    raw = (os.environ.get("PM_AI_CMD_PAUSE_ON_ERROR") or "1").strip().lower()
    if raw in ("0", "false", "no", "off"):
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.flush()
        except Exception:
            pass
    try:
        if getattr(sys.stdin, "isatty", lambda: False)():
            print(
                "\n[PM_AI] エラー終了です（終了コード "
                + str(code)
                + "）。ログを確認してから Enter キーを押してください…",
                flush=True,
            )
            try:
                input()
            except EOFError:
                os.system("pause")
        else:
            os.system("pause")
    except Exception:
        try:
            os.system("pause")
        except Exception:
            pass


def _run_main_with_optional_pause() -> int:
    try:
        from workbook_env_bootstrap import (
            run_cli_with_optional_pause_on_error as _run_cli_guarded,
        )
    except ImportError:
        import traceback

        _log(
            "注意: workbook_env_bootstrap を import できませんでした。"
            "エラー時 pause のみ簡易実装で続行します。"
        )

        def _run_cli_guarded(main_fn):
            code = 0
            try:
                try:
                    result = main_fn()
                    if result is not None:
                        code = int(result)
                except SystemExit as e:
                    c = e.code
                    if isinstance(c, int):
                        code = c
                    elif c:
                        code = 1
                    else:
                        code = 0
                except BaseException:
                    traceback.print_exc()
                    code = 1
            finally:
                try:
                    _pause_cli_error_standalone(code)
                except Exception:
                    pass
            return code

    return int(_run_cli_guarded(main))


if __name__ == "__main__":
    raise SystemExit(_run_main_with_optional_pause())

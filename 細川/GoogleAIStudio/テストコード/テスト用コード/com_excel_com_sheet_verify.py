# -*- coding: utf-8 -*-
"""
起動中の Excel に COM（pywin32）で接続し、シートごとに読み取り・書き込みが成功するか検証する。

前提:
  - 対象の .xlsm / .xlsx を先に Excel で開いておく
  - pip install pywin32

使い方（このフォルダで）:
  py com_excel_com_sheet_verify.py
  py com_excel_com_sheet_verify.py --book "C:\\path\\生産管理_AI配台テスト.xlsm"

結果:
  - 標準出力に表形式
  - log/com_excel_com_verify.txt にも UTF-8 で追記（--no-file で抑止）

接続:
  - GetActiveObject → GetObject(ブックのフルパス)（短縮パスも試行）→ GetObject(Class="Excel.Application")
  - VB の GetObject(, "Excel.Application") は pywin32 では Class= のみ指定（第1引数に "" を渡すとエラー）。
  - すべて失敗するとき: --dispatch-open で新規 Excel を起動して Open（検証用・二重起動に注意）。
  - RPC / OLE無効ヘッダー: 管理者の不一致、Googleドライブ等の仮想パス、32/64bit 不一致を疑う。
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime

SKIP_SHEET_NAMES = frozenset({"COM操作テストログ"})
TEST_ZZ = "ZZ1000"
TEST_ZZ_VAL = "__COM_TEST__"
TEST_A99 = "A99"
TEST_A99_VAL = "A666"


def _norm_path(p: str) -> str:
    p = os.path.normpath(str(p).strip().replace("/", "\\"))
    return os.path.normcase(os.path.abspath(p))


def workbook_matches_path(wb_com, disk_path: str) -> bool:
    try:
        fn = str(wb_com.FullName)
    except Exception:
        return False
    try:
        if _norm_path(disk_path) == _norm_path(fn):
            return True
    except Exception:
        pass
    try:
        return os.path.samefile(disk_path, fn)
    except Exception:
        return False


def find_workbook(xl_app, disk_path: str):
    try:
        n = int(xl_app.Workbooks.Count)
    except Exception:
        n = 0
    for i in range(1, n + 1):
        try:
            wb = xl_app.Workbooks(i)
        except Exception:
            continue
        if workbook_matches_path(wb, disk_path):
            return wb
    return None


def _book_path_variants(disk_path: str) -> list[str]:
    """GetObject(パス) 用。クラウド同期パスで失敗するとき 8.3 短縮名を試す。"""
    p = os.path.abspath(disk_path)
    seen: set[str] = set()
    out: list[str] = []
    for cand in (p, os.path.normpath(p)):
        k = cand.lower()
        if k not in seen:
            seen.add(k)
            out.append(cand)
    try:
        import win32api  # type: ignore

        sh = win32api.GetShortPathName(p)
        if sh:
            k = os.path.normcase(os.path.abspath(sh))
            if k not in seen:
                seen.add(k)
                out.append(sh)
    except Exception:
        pass
    return out


def attach_excel_and_workbook(
    book_path: str, *, allow_dispatch_open: bool = False
) -> tuple[object | None, object | None, str]:
    """
    Excel.Application と対象 Workbook を取得する。
    戻り値: (xl_app, wb, 接続方法またはエラー詳細)
    """
    from win32com.client import Dispatch, GetActiveObject, GetObject  # type: ignore

    book_path = os.path.abspath(book_path)
    errs: list[str] = []

    try:
        xl = GetActiveObject("Excel.Application")
        wb = find_workbook(xl, book_path)
        if wb is not None:
            return xl, wb, "GetActiveObject(Excel.Application)"
        errs.append(
            "GetActiveObject は成功したが、指定パスと一致するブックが Workbooks に無い"
        )
    except Exception as e:
        errs.append(f"GetActiveObject: {e}")

    for pvar in _book_path_variants(book_path):
        try:
            wb = GetObject(pvar)
            xl = wb.Application
            if workbook_matches_path(wb, book_path):
                label = "GetObject(ブックパス)"
                if pvar != book_path:
                    label += " [短縮パス]"
                return xl, wb, label
            errs.append(
                f"GetObject({pvar!r}) の FullName が指定パスと一致しない"
            )
        except Exception as e:
            errs.append(f"GetObject({pvar!r}): {e}")

    # VB の GetObject(, "Excel.Application") に相当: Pathname は渡さず Class のみ（"" は pywin32 でエラーになる）
    try:
        xl = GetObject(Class="Excel.Application")
        wb = find_workbook(xl, book_path)
        if wb is not None:
            return xl, wb, 'GetObject(Class="Excel.Application")'
        errs.append(
            'GetObject(Class="Excel.Application") は成功したが指定ブックが開かれていない'
        )
    except Exception as e:
        errs.append(f'GetObject(Class="Excel.Application"): {e}')

    if allow_dispatch_open:
        try:
            xl = Dispatch("Excel.Application")
            xl.Visible = True
            wb = xl.Workbooks.Open(book_path, UpdateLinks=0, ReadOnly=False)
            if wb is not None and workbook_matches_path(wb, book_path):
                return (
                    xl,
                    wb,
                    "Dispatch + Workbooks.Open（--dispatch-open・既存の Excel とは別プロセスの可能性）",
                )
            errs.append("Dispatch+Open で得たブックのパスが一致しない")
        except Exception as e:
            errs.append(f"Dispatch+Workbooks.Open: {e}")

    return None, None, " | ".join(errs)


def visible_label(v) -> str:
    try:
        iv = int(v)
    except Exception:
        return str(v)
    if iv == -1:
        return "表示"
    if iv == 0:
        return "非表示"
    if iv == 2:
        return "VeryHidden"
    return str(iv)


def safe_str(x, max_len: int = 80) -> str:
    s = repr(x) if x is not None else ""
    if len(s) > max_len:
        return s[: max_len - 3] + "..."
    return s


def try_read_a1(ws) -> tuple[str, str]:
    try:
        v = ws.Range("A1").Value
        return "OK", safe_str(v)
    except Exception as e:
        return "NG", str(e)


def try_used_range(ws) -> tuple[str, str]:
    try:
        ur = ws.UsedRange
        addr = str(ur.Address).replace("$", "")
        return "OK", addr
    except Exception as e:
        return "NG", str(e)


def try_zz_write(ws) -> tuple[str, str]:
    try:
        ws.Range(TEST_ZZ).Value = TEST_ZZ_VAL
        ws.Range(TEST_ZZ).ClearContents()
        return "OK", ""
    except Exception as e:
        return "NG", str(e)


def try_a99_roundtrip(ws) -> tuple[str, str]:
    r = ws.Range(TEST_A99)
    try:
        old = r.Value
    except Exception as e:
        return "NG(退避)", str(e)
    try:
        r.Value = TEST_A99_VAL
        back = r.Value
        if str(back) != TEST_A99_VAL:
            try:
                r.Value = old
            except Exception:
                pass
            return "NG(不一致)", safe_str(back)
        r.Value = old
        return "OK", ""
    except Exception as e:
        try:
            r.Value = old
        except Exception:
            pass
        return "NG", str(e)


def verify_sheet(ws) -> dict:
    name = ""
    try:
        name = str(ws.Name)
    except Exception:
        name = "?"

    out: dict = {"name": name}
    try:
        out["protect"] = bool(ws.ProtectContents)
    except Exception as e:
        out["protect"] = f"? {e}"

    try:
        out["visible"] = visible_label(ws.Visible)
    except Exception as e:
        out["visible"] = str(e)

    st, msg = try_read_a1(ws)
    out["a1"] = st
    out["a1_note"] = msg

    st, msg = try_used_range(ws)
    out["used"] = st
    out["used_note"] = msg

    st, msg = try_zz_write(ws)
    out["zz"] = st
    out["zz_note"] = msg

    st, msg = try_a99_roundtrip(ws)
    out["a99"] = st
    out["a99_note"] = msg

    return out


def _planning_repo_root() -> str:
    """テスト用コード/ に置いたときは親フォルダ（planning_core がある場所）を返す。"""
    here = os.path.dirname(os.path.abspath(__file__))
    parent = os.path.dirname(here)
    if os.path.isfile(os.path.join(parent, "planning_core.py")):
        return parent
    return here


def default_book_path(repo_root: str) -> str:
    cand = os.path.join(repo_root, "生産管理_AI配台テスト.xlsm")
    if os.path.isfile(cand):
        return cand
    return cand


def main() -> int:
    repo_root = _planning_repo_root()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parser = argparse.ArgumentParser(description="Excel COM シート書き込み検証（起動中のブック）")
    parser.add_argument(
        "--book",
        default=default_book_path(repo_root),
        help="検証するブックのパス（既定: リポジトリ直下の 生産管理_AI配台テスト.xlsm）",
    )
    parser.add_argument("--no-file", action="store_true", help="log ファイルへ書かない")
    parser.add_argument(
        "--dispatch-open",
        action="store_true",
        help="接続に失敗したときだけ Excel を新規起動し Workbooks.Open する（二重起動・別ウィンドウになり得る）",
    )
    args = parser.parse_args()
    book_path = os.path.abspath(args.book)

    try:
        import win32com.client  # type: ignore  # noqa: F401
    except ImportError:
        print("pywin32 が未インストールです: pip install pywin32", file=sys.stderr)
        return 2

    xl, wb, attach_info = attach_excel_and_workbook(
        book_path, allow_dispatch_open=bool(args.dispatch_open)
    )
    if xl is None or wb is None:
        print("Excel COM に接続できませんでした。", file=sys.stderr)
        print(attach_info, file=sys.stderr)
        print("", file=sys.stderr)
        print("確認:", file=sys.stderr)
        print("  ・--book は「名前を付けて保存」で見えるフルパスと一致させる。", file=sys.stderr)
        print("  ・PowerShell/Cursor と Excel の「管理者として実行」の有無を揃える。", file=sys.stderr)
        print("  ・RPC / OLE無効ヘッダーは、権限の食い違い・ドライブ仮想パス・PythonとOfficeの32/64bit不一致で起きがち。", file=sys.stderr)
        print("  ・位数確認: py -0p", file=sys.stderr)
        if not args.dispatch_open:
            print("  ・試行: py com_excel_com_sheet_verify.py --dispatch-open", file=sys.stderr)
        return 3

    rows: list[dict] = []
    try:
        n = int(wb.Worksheets.Count)
    except Exception:
        n = 0

    for i in range(1, n + 1):
        try:
            ws = wb.Worksheets(i)
        except Exception as ex:
            rows.append(
                {
                    "name": f"(Worksheets({i})取得失敗)",
                    "protect": "",
                    "visible": "",
                    "a1": "NG",
                    "a1_note": str(ex),
                    "used": "",
                    "used_note": "",
                    "zz": "",
                    "zz_note": "",
                    "a99": "",
                    "a99_note": "",
                }
            )
            continue
        try:
            sn = str(ws.Name)
        except Exception:
            sn = ""
        if sn in SKIP_SHEET_NAMES:
            continue
        rows.append(verify_sheet(ws))

    # --- 出力 ---
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines: list[str] = []
    lines.append(f"=== Excel COM 検証 {ts} ===")
    lines.append(f"ブック: {book_path}")
    lines.append(f"接続: {attach_info}")
    lines.append("")
    hdr = (
        "シート名\t保護\t表示\tA1\tUsedRange\tZZ書込\tA99<->A666"
    )
    lines.append(hdr)
    for r in rows:
        prot = r.get("protect")
        if isinstance(prot, bool):
            prot_s = "あり" if prot else "なし"
        else:
            prot_s = str(prot)
        line = "\t".join(
            [
                str(r.get("name", "")),
                prot_s,
                str(r.get("visible", "")),
                f"{r.get('a1','')}" + (f" ({r.get('a1_note','')})" if r.get("a1_note") else ""),
                f"{r.get('used','')}" + (f" ({r.get('used_note','')})" if r.get("used_note") else ""),
                f"{r.get('zz','')}" + (f" ({r.get('zz_note','')})" if r.get("zz_note") else ""),
                f"{r.get('a99','')}" + (f" ({r.get('a99_note','')})" if r.get("a99_note") else ""),
            ]
        )
        lines.append(line)

    text = "\n".join(lines) + "\n"
    try:
        sys.stdout.write(text)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(text.encode("utf-8", errors="replace"))

    if not args.no_file:
        log_dir = os.path.join(repo_root, "log")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, "com_excel_com_verify.txt")
        with open(log_path, "a", encoding="utf-8", newline="\n") as f:
            f.write(text)
        print(f"(ログ追記: {log_path})", file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

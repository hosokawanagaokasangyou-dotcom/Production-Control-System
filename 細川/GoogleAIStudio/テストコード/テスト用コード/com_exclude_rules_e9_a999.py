# -*- coding: utf-8 -*-
"""
起動中の Excel ブック（COM）で「設定_配台不要工程」シートの E9 に文字列を書き込むだけの最小テスト。
元のセル値へは戻しません（上書きしたまま終了します）。

【重要】画面の Excel に反映されないとき
  planning_core の _com_attach は「見えている Excel に繋がらない」と非表示の別プロセスで開き、
  終了時に保存せず閉じるため、編集が破棄され画面は変わりません。
  本スクリプトの既定は com_excel_com_sheet_verify と同じ接続（起動中ブック優先）です。

既定: 値は A999、Save はしない（メモリ上のブックのみ更新）。
  py com_exclude_rules_e9_a999.py

開いていないときだけ新しい Excel で開く（表示・別プロセスの可能性）:
  py com_exclude_rules_e9_a999.py --dispatch-open

旧来の planning_core 接続（非推奨・画面と別インスタンスになり得る）:
  py com_exclude_rules_e9_a999.py --planning-attach

ディスクに保存も試す:
  py com_exclude_rules_e9_a999.py --save
"""

from __future__ import annotations

import argparse
import logging
import os
import sys


def _script_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _planning_repo_root() -> str:
    here = _script_dir()
    parent = os.path.dirname(here)
    if os.path.isfile(os.path.join(parent, "planning_core.py")):
        return parent
    return here


def _refresh_excel_ui(xl, wb, ws) -> None:
    """編集が画面に出るよう、可能ならブック・シートを前面にする。"""
    try:
        xl.ScreenUpdating = True
    except Exception:
        pass
    try:
        if not bool(xl.Visible):
            xl.Visible = True
    except Exception:
        pass
    try:
        wb.Activate()
    except Exception:
        pass
    try:
        ws.Activate()
    except Exception:
        pass


def main() -> int:
    sd = _script_dir()
    repo = _planning_repo_root()
    for d in (repo, sd):
        if d not in sys.path:
            sys.path.insert(0, d)

    logging.disable(logging.WARNING)
    import planning_core as pc
    from com_excel_com_sheet_verify import attach_excel_and_workbook

    logging.disable(logging.NOTSET)
    logging.getLogger().setLevel(logging.CRITICAL)

    default_book = os.path.join(repo, "生産管理_AI配台テスト.xlsm")
    p = argparse.ArgumentParser(
        description="設定_配台不要工程 E9 へ A999（既定）を COM 書込（元値の復元はしない）"
    )
    p.add_argument("--book", default=default_book, help="マクロブックのフルパス")
    p.add_argument("--row", type=int, default=9, help="書き込み行（既定 9）")
    p.add_argument("--col", type=int, default=5, help="列番号（E=5。既定 5）")
    p.add_argument("--text", default="A999", help="セルに書く文字列（既定 A999）")
    p.add_argument(
        "--save",
        action="store_true",
        help="wb.Save() も試す（既定はメモリのみ。R/O や二重起動で失敗し得る）",
    )
    p.add_argument(
        "--dispatch-open",
        action="store_true",
        help="起動中 Excel にブックが無いときだけ、新規 Excel(表示)で Workbooks.Open する",
    )
    p.add_argument(
        "--planning-attach",
        action="store_true",
        help="planning_core._com_attach を使う（非表示の別プロセスになり得る。画面とズレやすい）",
    )
    args = p.parse_args()

    if args.dispatch_open and args.planning_attach:
        print("--dispatch-open と --planning-attach は同時に指定しないでください。", file=sys.stderr)
        return 2

    book = os.path.abspath(args.book)
    if not os.path.isfile(book):
        print(f"ブックがありません: {book}", file=sys.stderr)
        return 2

    xl = wb = None
    info: dict
    attach_label = ""

    if args.planning_attach:
        attached = pc._com_attach_open_macro_workbook(book, "COM_E9_A999")
        if attached is None:
            print(
                "COM でブックに接続できません（--planning-attach）。",
                file=sys.stderr,
            )
            return 3
        xl, wb, info = attached
        attach_label = "planning_core._com_attach_open_macro_workbook"
        if info.get("mode") == "quit_excel":
            print(
                "[警告] 非表示の別 Excel でブックを開いています。"
                " 終了時に Save しないと変更は破棄され、いま画面で見ている Excel には出ません。",
                file=sys.stderr,
            )
    else:
        xl, wb, attach_label = attach_excel_and_workbook(
            book, allow_dispatch_open=bool(args.dispatch_open)
        )
        if wb is None or xl is None:
            print("COM で接続できませんでした。", file=sys.stderr)
            print(attach_label, file=sys.stderr)
            print(
                "ヒント: 対象を Excel で開き、--book を「名前を付けて保存」のフルパスと一致させてください。",
                file=sys.stderr,
            )
            print(
                "  試す: py com_exclude_rules_e9_a999.py --dispatch-open",
                file=sys.stderr,
            )
            return 3
        opened_here = bool(args.dispatch_open) and (
            "Dispatch" in attach_label or "Open" in attach_label
        )
        info = {"mode": "keep", "opened_wb_here": opened_here}

    ok = False
    try:
        try:
            xl.DisplayAlerts = False
        except Exception:
            pass

        try:
            print(f"接続: {attach_label}")
            print(f"ブック FullName: {wb.FullName}")
            print(f"Excel.Visible: {xl.Visible}")
        except Exception:
            pass

        ws = wb.Worksheets(pc.EXCLUDE_RULES_SHEET_NAME)
        r, c = int(args.row), int(args.col)
        before = ws.Cells(r, c).Value
        ws.Cells(r, c).Value = args.text
        after = ws.Cells(r, c).Value

        _refresh_excel_ui(xl, wb, ws)

        addr = f"{pc.EXCLUDE_RULES_SHEET_NAME} R{r}C{c}（列{c}がE列）"
        print(f"変更前: {before!r}")
        print(f"書込:   {args.text!r}")
        print(f"読取:   {after!r}")
        print(f"一致:   {str(after) == str(args.text)}  （{addr}）")

        if not args.save:
            print("メモリ上のみ更新しました（--save なしのため Save しません）。")
            ok = True
            return 0

        try:
            if bool(wb.ReadOnly):
                print(
                    "[save] 読み取り専用のため Save しません。Excel を1つにするか --save を外してください。",
                    file=sys.stderr,
                )
                ok = True
                return 5
        except Exception:
            pass
        try:
            wb.Save()
            print("[save] Save しました。")
            ok = True
            return 0
        except Exception as ex:
            print(f"[save] 失敗: {ex}", file=sys.stderr)
            ok = True
            return 5

    except Exception as ex:
        print(f"エラー: {ex}", file=sys.stderr)
        import traceback

        traceback.print_exc()
        return 4
    finally:
        if xl is not None and wb is not None:
            pc._com_release_workbook_after_mutation(xl, wb, info, ok)


if __name__ == "__main__":
    raise SystemExit(main())

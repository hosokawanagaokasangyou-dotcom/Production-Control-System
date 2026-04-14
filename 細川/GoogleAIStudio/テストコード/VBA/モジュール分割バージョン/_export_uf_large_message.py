# -*- coding: utf-8 -*-
"""Excel VBE で UF_LargeMessage を組み立てて .frm/.frx をエクスポートする補助スクリプト。
   信頼センター「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が必要。
   出力は ASCII パス（例: C:\\Temp\\vba_export）を推奨。"""
from __future__ import annotations

import argparse
import os
import shutil
import sys


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--out-dir", required=True, help=r"ASCII path e.g. C:\Temp\vba_export")
    args = p.parse_args()
    out_dir = os.path.abspath(args.out_dir)
    os.makedirs(out_dir, exist_ok=True)

    try:
        import win32com.client as win32  # type: ignore
    except ImportError:
        print("pywin32 が必要です。", file=sys.stderr)
        return 1

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Add()
    try:
        vb = wb.VBProject
        # vbext_ct_MSForm = 3
        comp = vb.VBComponents.Add(3)
        comp.Name = "UF_LargeMessage"

        ctl = comp.Designer.Controls
        # 左ストライプ
        lbl = ctl.Add("Forms.Label.1", "lblStripe", True)
        lbl.Left = 0
        lbl.Top = 0
        lbl.Width = 8
        lbl.Height = 324
        lbl.BackColor = 0xC0C0C0
        lbl.BorderStyle = 0

        tb = ctl.Add("Forms.TextBox.1", "txtBody", True)
        tb.Left = 16
        tb.Top = 16
        tb.Width = 520
        tb.Height = 264
        tb.MultiLine = True
        tb.WordWrap = True
        tb.ScrollBars = 2  # fmScrollBarsVertical
        tb.Locked = True
        tb.Font.Name = "Meiryo UI"
        tb.Font.Size = 14

        def add_cmd(name: str, caption: str, left: int, top: int, width: int, height: int):
            b = ctl.Add("Forms.CommandButton.1", name, True)
            b.Caption = caption
            b.Left = left
            b.Top = top
            b.Width = width
            b.Height = height
            b.Font.Name = "Meiryo UI"
            b.Font.Size = 12
            return b

        add_cmd("cmdOK", "OK", 224, 288, 104, 28)
        add_cmd("cmdCancel", "キャンセル", 344, 288, 112, 28)
        add_cmd("cmdYes", "はい", 184, 288, 104, 28)
        add_cmd("cmdNo", "いいえ", 304, 288, 104, 28)

        comp.Designer.Controls("cmdOK").Default = True
        comp.Designer.Controls("cmdCancel").Cancel = True

        base = os.path.join(out_dir, "UF_LargeMessage")
        comp.Export(base + ".frm")
        print("Exported:", base + ".frm", sep="\n")
        print(base + ".frx")
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()

    print(
        "Merge: append Option Explicit and event code from repo UF_LargeMessage.frm "
        "after the Attribute lines of the exported .frm, then copy .frm + .frx together."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

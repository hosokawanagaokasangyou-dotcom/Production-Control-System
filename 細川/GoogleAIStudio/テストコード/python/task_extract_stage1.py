# -*- coding: utf-8 -*-
"""
段階1: 加工計画DATA から未完了タスクを抽出し output/plan_input_tasks.xlsx を生成する。
マクロで当ファイルをブックへ取り込み、「配台計画_タスク入力」で特別指定を編集した後、段階2を実行する。

環境: planning_core と同様、openpyxl（閉じたブックの I/O）と、保存ロック時の xlwings 同期を使用する。
初回はリポジトリの requirements.txt を pip し、Excel デスクトップ版を用意すること。
"""
import os
import sys
import ctypes

os.chdir(os.path.dirname(os.path.abspath(__file__)))

if os.name == "nt":
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 3)

import planning_core as pc


def main():
    if not pc.TASKS_INPUT_WORKBOOK:
        print("TASK_INPUT_WORKBOOK が未設定です。VBA からマクロ実行してください。", file=sys.stderr)
        sys.exit(2)
    try:
        ok = pc.run_stage1_extract()
    except pc.PlanningValidationError as e:
        msg = str(e).strip() or "マスタ skills の検証で中断しました。"
        if not os.path.isfile(pc.stage2_blocking_message_path):
            pc._write_stage2_blocking_message(msg)
        print(msg, file=sys.stderr)
        sys.exit(3)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()

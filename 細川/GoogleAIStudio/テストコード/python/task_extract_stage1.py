# -*- coding: utf-8 -*-
"""
段階1: 加工計画DATA から未完了タスクを抽出し output/plan_input_tasks.xlsx を生成する。
マクロで当ファイルをブックへ取り込み、「配台計画_タスク入力」で特別指定を編集した後、段階2を実行する。

環境: planning_core と同様、openpyxl（閉じたブックの I/O）と、保存ロック時の xlwings 同期を使用する。
初回はリポジトリの requirements.txt を pip し、Excel デスクトップ版を用意すること。
"""
import ctypes
import os
import sys
import traceback
from datetime import datetime

# planning_core は import 途中（FileHandler より前）で落ちると execution_log が作られない。
# VBA は「ログ無し」を判定するため、読み込み前に必ず log/execution_log.txt を用意する。
def _repo_root_for_stage1() -> str:
    d = os.path.dirname(os.path.abspath(__file__))
    return os.path.dirname(d) if os.path.basename(d).lower() == "python" else d


def _append_execution_log_line(level: str, msg: str) -> str:
    root = _repo_root_for_stage1()
    log_dir = os.path.join(root, "log")
    path = os.path.join(log_dir, "execution_log.txt")
    os.makedirs(log_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts} - {level} - {msg}\n"
    with open(path, "a", encoding="utf-8-sig", newline="\n") as f:
        f.write(line)
    return path


os.chdir(os.path.dirname(os.path.abspath(__file__)))

if os.name == "nt":
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 3)

try:
    _append_execution_log_line(
        "INFO", "段階1: task_extract_stage1 起動（planning_core 読み込み前）"
    )
except OSError as ex:
    print(f"log/execution_log.txt を開けません: {ex}", file=sys.stderr)

try:
    import planning_core as pc
except Exception:
    try:
        path = os.path.join(_repo_root_for_stage1(), "log", "execution_log.txt")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "a", encoding="utf-8-sig", newline="\n") as f:
            f.write(
                f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - ERROR - "
                "planning_core の import に失敗しました\n"
            )
            f.write(traceback.format_exc())
    except OSError:
        pass
    traceback.print_exc()
    sys.exit(1)


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

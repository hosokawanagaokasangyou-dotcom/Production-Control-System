# -*- coding: utf-8 -*-
"""Apply trial-order exclude patch to UTF-8 _core.py from git parent."""
from __future__ import annotations

import subprocess
from pathlib import Path

repo = Path(__file__).resolve().parent
rel = next(
    p
    for p in subprocess.check_output(
        ["git", "-C", str(repo), "ls-files", "*/planning_core/_core.py"],
        text=True,
    ).splitlines()
    if p.strip()
)
parent = subprocess.check_output(
    ["git", "-C", str(repo), "rev-parse", "HEAD^"],
    text=True,
).strip()
raw = subprocess.check_output(["git", "-C", str(repo), "show", f"{parent}:{rel}"])
text = raw.decode("utf-8")

old_sig = 'def _apply_planning_sheet_post_load_mutations(\n    df: "pd.DataFrame", wb_path: str, log_prefix: str\n) -> None:'
new_sig = '''def _apply_planning_sheet_post_load_mutations(
    df: "pd.DataFrame",
    wb_path: str,
    log_prefix: str,
    *,
    apply_exclude_rules_from_config: bool = True,
) -> None:'''
if old_sig not in text:
    raise SystemExit("signature not found")
text = text.replace(old_sig, new_sig, 1)

old_doc = (
    '    """\n'
    "    配台計画_タスク入力を DataFrame 化した直後の共通処理。\n"
    "    段階2の ``load_planning_tasks_df`` と同じ（設定シート・分割行・配台不要ルール）。\n"
    '    """\n'
)
new_doc = (
    '    """\n'
    "    配台計画_タスク入力を DataFrame 化した直後の共通処理。\n"
    "    段階2の ``load_planning_tasks_df`` と同じ（設定シートの行同期・分割行の自動配台不要）。\n"
    "\n"
    "    ``apply_exclude_rules_from_config=False`` のときは「設定_配台不要工程」の C/E に基づく\n"
    "    計画シートへの「配台不要」の上書きを行わない（試行順のみ再計算する経路でシート上の手動クリアを残す）。\n"
    '    """\n'
)
if old_doc not in text:
    raise SystemExit("old mutation doc not found")
text = text.replace(old_doc, new_doc, 1)

old_try = """    try:
        apply_exclude_rules_config_to_plan_df(df, wb_path, log_prefix)
    except Exception as ex:
        logging.warning(
            \"%s: 設定シートによる配台不要適用で例外（続行）: %s\",
            log_prefix,
            ex,
        )
"""
new_try = """    if apply_exclude_rules_from_config:
        try:
            apply_exclude_rules_config_to_plan_df(df, wb_path, log_prefix)
        except Exception as ex:
            logging.warning(
                \"%s: 設定シートによる配台不要適用で例外（続行）: %s\",
                log_prefix,
                ex,
            )
"""
if old_try not in text:
    raise SystemExit("try block not found")
text = text.replace(old_try, new_try, 1)

old_call = '_apply_planning_sheet_post_load_mutations(df, path, "配台試行順番更新")'
new_call = '''_apply_planning_sheet_post_load_mutations(
        df, path, "配台試行順番更新", apply_exclude_rules_from_config=False
    )'''
if old_call not in text:
    raise SystemExit("refresh call not found")
text = text.replace(old_call, new_call, 1)

old_rdoc = """    直前に ``_apply_planning_sheet_post_load_mutations`` を実行するため、
    ``load_planning_tasks_df``（段階2 読込）と同じ経路で「設定_配台不要工程」の同期反映・
    分割行の自動配台不要・ルール適用が行われる。「配台不要」を手動でクリアした行は、
    ルールで再び yes にならない限り配台対象に戻り、試行順に含まれる。"""
new_rdoc = """    事前処理は ``_apply_planning_sheet_post_load_mutations``（設定シートの行同期・分割行の自動配台不要）。
    ただし本経路では **「設定_配台不要工程」の C/E による計画シートへの配台不要の強制上書きは行わない**。
    シート上で消した「配台不要」は復活しない。段階2の ``load_planning_tasks_df`` では C/E 上書きが有効。"""
if old_rdoc not in text:
    raise SystemExit("refresh doc not found")
text = text.replace(old_rdoc, new_rdoc, 1)

out = repo / rel
out.write_text(text, encoding="utf-8", newline="\n")
print("wrote", out)

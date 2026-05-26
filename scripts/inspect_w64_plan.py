# -*- coding: utf-8 -*-
"""W6-4 計画行の簡易診断（段階2前）。"""
import json
import os
import sys

_py = os.path.dirname(os.path.abspath(__file__))
_repo = os.path.dirname(_py)
_code_py = os.path.join(_repo, "code", "python")
if _code_py not in sys.path:
    sys.path.insert(0, _code_py)
os.chdir(_code_py)

os.environ.setdefault("PM_AI_REPO_ROOT", _repo)
os.environ.setdefault("PM_AI_PLAN_INPUT_PATH", os.path.join(_repo, "output", "plan_input_tasks.xlsx"))
os.environ.setdefault("PM_AI_MASTER_WORKBOOK", os.path.join(_repo, "master.xlsm"))
os.environ.setdefault("PM_AI_SKIP_WORKBOOK_ENV_SHEET", "1")

import planning_core as pc  # noqa: E402

df = pc.load_planning_tasks_df()
mask = df.astype(str).apply(lambda c: c.str.contains("W6-4", na=False)).any(axis=1)
rows = df[mask]
print(json.dumps({"w64_row_count": int(len(rows))}, ensure_ascii=False))
if len(rows) == 0:
    sys.exit(0)
pref = [c for c in rows.columns if any(k in str(c) for k in ("依頼", "工程", "機械", "開始", "納期", "試行", "未加工", "換算", "OP"))]
cols = pref if pref else list(rows.columns[:12])
for i, (_, r) in enumerate(rows.iterrows()):
    if i >= 12:
        break
    item = {str(c): str(r.get(c, "")) for c in cols}
    print(json.dumps(item, ensure_ascii=False))

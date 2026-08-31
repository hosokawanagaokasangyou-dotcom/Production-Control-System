# -*- coding: utf-8 -*-
"""
特別ルール L14: 工程名+機械名の配台優先（同一実機械上の連続選好）。

グローバルな配台試行順番は書き換えない。同一占有キー上で、設定された工程を
複数依頼にわたり連続して取る。既定はエンボス・優先度通常。
機械名が空ならその工程の全機械に適用する。
"""

from __future__ import annotations

import json
import logging
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any

PRIORITY_HIGHEST = 0
PRIORITY_HIGH = 1
PRIORITY_NORMAL = 2
PRIORITY_LOW = 3

_PRIORITY_LABELS = {
    "最優先": PRIORITY_HIGHEST,
    "優先": PRIORITY_HIGH,
    "通常": PRIORITY_NORMAL,
    "優先度低": PRIORITY_LOW,
}

_JSON_KEY = "processMachinePriorities"


def process_from_eq_line(eq_line: str) -> str:
    """設備ラインキー（工程+機械）から工程名を取り出す。工程名そのものならそのまま返す。"""
    ek = str(eq_line or "").strip()
    if not ek:
        return ""
    if "+" in ek:
        ek = ek.split("+", 1)[0]
    return str(ek).strip()


def last_process_for_occupancy(machine_handoff: dict | None, occupancy: str) -> str:
    """占有キー上の直前工程名。last_process を優先し、無ければ last_eq から抽出する。"""
    occ = (occupancy or "").strip()
    if not machine_handoff or not occ:
        return ""
    stored = (machine_handoff.get("last_process") or {}).get(occ)
    if stored is not None and str(stored).strip() != "":
        return str(stored)
    last_eq = str((machine_handoff.get("last_eq") or {}).get(occ) or "")
    return process_from_eq_line(last_eq)


def remember_last_process(
    machine_handoff: dict | None,
    occupancy: str,
    *,
    process_name: str = "",
    eq_line: str = "",
) -> None:
    if not machine_handoff:
        return
    occ = (occupancy or "").strip()
    if not occ:
        return
    machine_handoff.setdefault("last_process", {})
    proc = str(process_name or "").strip() or process_from_eq_line(eq_line)
    machine_handoff["last_process"][occ] = proc


def _norm_process(raw: Any) -> str:
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    return re.sub(r"[\s　]+", "", t)


def _norm_machine(raw: Any) -> str:
    t = unicodedata.normalize("NFKC", str(raw or ""))
    t = t.replace("\u00a0", " ").replace("\u3000", " ")
    t = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", t)
    return re.sub(r"\s+", " ", t).strip()


def parse_priority_label(raw: Any) -> int:
    s = unicodedata.normalize("NFKC", str(raw or "").strip())
    if not s:
        return PRIORITY_NORMAL
    if s in _PRIORITY_LABELS:
        return _PRIORITY_LABELS[s]
    low = s.casefold()
    if low in ("highest", "max"):
        return PRIORITY_HIGHEST
    if low in ("high",):
        return PRIORITY_HIGH
    if low in ("low",):
        return PRIORITY_LOW
    return PRIORITY_NORMAL


@dataclass(frozen=True)
class ProcessMachinePriorityRule:
    process_name_key: str
    machine_name_key: str
    priority: int
    consecutive: bool
    enabled: bool = True


def default_priority_rules() -> list[ProcessMachinePriorityRule]:
    return [
        ProcessMachinePriorityRule(
            process_name_key=_norm_process("エンボス"),
            machine_name_key="",
            priority=PRIORITY_NORMAL,
            consecutive=True,
            enabled=True,
        )
    ]


def _rule_from_raw(raw: dict[str, Any]) -> ProcessMachinePriorityRule | None:
    proc = _norm_process(raw.get("processName") or raw.get("工程名") or "")
    if not proc:
        return None
    enabled = raw.get("enabled", True)
    if isinstance(enabled, str):
        enabled = enabled.strip().lower() not in ("0", "false", "no", "off")
    if not enabled:
        return None
    cons = raw.get("consecutive", raw.get("連続配置", True))
    if isinstance(cons, str):
        cons = cons.strip().lower() not in ("0", "false", "no", "off")
    return ProcessMachinePriorityRule(
        process_name_key=proc,
        machine_name_key=_norm_machine(raw.get("machineName") or raw.get("機械名") or ""),
        priority=parse_priority_label(raw.get("priority") or raw.get("優先度")),
        consecutive=bool(cons),
        enabled=True,
    )


def rules_from_document(doc: dict[str, Any] | None) -> list[ProcessMachinePriorityRule]:
    if not isinstance(doc, dict) or _JSON_KEY not in doc:
        return default_priority_rules()
    raw_list = doc.get(_JSON_KEY)
    if not isinstance(raw_list, list):
        return default_priority_rules()
    out: list[ProcessMachinePriorityRule] = []
    for item in raw_list:
        if not isinstance(item, dict):
            continue
        rule = _rule_from_raw(item)
        if rule is not None:
            out.append(rule)
    return out


def load_priority_rules() -> list[ProcessMachinePriorityRule]:
    try:
        from planning_core.dispatch_rules.paths import (
            bundled_template_path,
            resolve_dispatch_special_rules_json,
        )
    except Exception:
        return default_priority_rules()
    paths: list[Path] = []
    resolved = resolve_dispatch_special_rules_json()
    if resolved:
        paths.append(Path(resolved))
    bundled = bundled_template_path()
    if bundled.is_file():
        paths.append(bundled)
    for p in paths:
        try:
            raw = json.loads(p.read_text(encoding="utf-8"))
        except Exception as ex:
            logging.warning("工程機械優先 JSON を読めません: %s (%s)", p, ex)
            continue
        if isinstance(raw, dict):
            if _JSON_KEY in raw:
                return rules_from_document(raw)
            return default_priority_rules()
    return default_priority_rules()


def match_rule(
    task: dict,
    rules: list[ProcessMachinePriorityRule] | None = None,
) -> ProcessMachinePriorityRule | None:
    if rules is None:
        rules = load_priority_rules()
    proc = _norm_process(task.get("machine"))
    mach = _norm_machine(task.get("machine_name"))
    if not proc:
        return None
    generic: ProcessMachinePriorityRule | None = None
    for rule in rules:
        if not rule.enabled:
            continue
        if rule.process_name_key != proc:
            continue
        if rule.machine_name_key:
            if rule.machine_name_key == mach:
                return rule
            continue
        if generic is None:
            generic = rule
    return generic


def priority_rank_for_task(
    task: dict,
    rules: list[ProcessMachinePriorityRule] | None = None,
) -> int:
    rule = match_rule(task, rules)
    if rule is None:
        return PRIORITY_NORMAL
    return rule.priority


def _request_id(task: dict) -> str:
    """同一依頼NO（枝番は rule_task_id、通常は task_id）。§A-1 と揃える。"""
    if not isinstance(task, dict):
        return ""
    rid = str(task.get("rule_task_id") or "").strip()
    if rid:
        return rid
    return str(task.get("task_id") or "").strip()


def consecutive_prefer_sort_penalty(
    task: dict,
    machine_handoff: dict | None,
    *,
    occupancy: str = "",
    rules: list[ProcessMachinePriorityRule] | None = None,
) -> int:
    """0=直前と同じ連続対象工程。グローバル試行順より後のタイブレーク専用。"""
    rule = match_rule(task, rules)
    if rule is None or not rule.consecutive:
        return 1
    occ = (occupancy or "").strip()
    if not occ:
        eqt = str(task.get("equipment_line_key") or task.get("machine") or "").strip()
        occ = eqt
    last = last_process_for_occupancy(machine_handoff, occ)
    if _norm_process(last) == rule.process_name_key:
        return 0
    return 1


def may_skip_lower_trial_blocker(
    candidate: dict,
    blocker: dict,
    last_process: str,
    *,
    rules: list[ProcessMachinePriorityRule] | None = None,
) -> bool:
    """
    同一実機械上で、候補の連続配置のために低試行順の別工程ブロッカーを飛ばしてよいか。
    他設備のグローバル試行順ゲートは対象外（呼び出し側が同一占有に限定すること）。
    同一依頼NO（§A-1）の先行工程は飛ばさない。
    """
    cand_id = _request_id(candidate)
    block_id = _request_id(blocker)
    if cand_id and cand_id == block_id:
        return False
    rule = match_rule(candidate, rules)
    if rule is None or not rule.consecutive:
        return False
    if rule.priority == PRIORITY_LOW:
        return False
    cand_proc = _norm_process(candidate.get("machine"))
    block_proc = _norm_process(blocker.get("machine"))
    if not cand_proc or cand_proc == block_proc:
        return False
    if rule.priority in (PRIORITY_HIGHEST, PRIORITY_HIGH):
        return True
    return _norm_process(last_process) == cand_proc


def should_defer_blocker_behind_consecutive_peers(
    blocker: dict,
    peers: list,
    *,
    occupancy: str,
    last_process: str,
    occupancy_of,
    rules: list[ProcessMachinePriorityRule] | None = None,
) -> bool:
    """同一占有の連続対象がこのブロッカーを飛ばせるなら、ソートで後回しにする。"""
    occ = (occupancy or "").strip()
    if not occ:
        return False
    for peer in peers:
        if peer is blocker:
            continue
        if occupancy_of(peer) != occ:
            continue
        if may_skip_lower_trial_blocker(peer, blocker, last_process, rules=rules):
            return True
    return False

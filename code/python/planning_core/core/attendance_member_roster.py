# -*- coding: utf-8 -*-
"""メンバー勤怠グリッド用の名簿（スキル有無に依存しない並び・主担当）。"""

from __future__ import annotations

from typing import Any

PRIMARY_ROLE_POST = "後加工"
PRIMARY_ROLE_LOGISTICS = "物流"
PRIMARY_ROLE_CHOICES = (PRIMARY_ROLE_POST, PRIMARY_ROLE_LOGISTICS)

# 既定名簿（並び順＝メンバー勤怠グリッドの行順）
DEFAULT_MEMBER_ROSTER: list[dict[str, str]] = [
    {"name": "細川　守", "primary_role": PRIMARY_ROLE_POST},
    {"name": "砂田　　奈美", "primary_role": PRIMARY_ROLE_LOGISTICS},
    {"name": "古家　　淳子", "primary_role": PRIMARY_ROLE_POST},
    {"name": "宮島　　剛", "primary_role": PRIMARY_ROLE_POST},
    {"name": "図司　智子", "primary_role": PRIMARY_ROLE_POST},
    {"name": "冨田　裕子", "primary_role": PRIMARY_ROLE_POST},
    {"name": "森岡　真由美", "primary_role": PRIMARY_ROLE_POST},
    {"name": "竹内　正美", "primary_role": PRIMARY_ROLE_POST},
    {"name": "菅沼　めぐみ", "primary_role": PRIMARY_ROLE_POST},
    {"name": "森下　　誠", "primary_role": PRIMARY_ROLE_POST},
    {"name": "小川　達也", "primary_role": PRIMARY_ROLE_LOGISTICS},
    {"name": "西田   憲史", "primary_role": PRIMARY_ROLE_LOGISTICS},
    {"name": "近藤　清高", "primary_role": PRIMARY_ROLE_LOGISTICS},
    {"name": "東出　繫利", "primary_role": PRIMARY_ROLE_LOGISTICS},
]


def _normalize_role(role: str) -> str:
    r = str(role or "").strip()
    if r in PRIMARY_ROLE_CHOICES:
        return r
    return PRIMARY_ROLE_POST


def _normalize_entry(raw: dict[str, Any]) -> dict[str, str] | None:
    if not isinstance(raw, dict):
        return None
    name = str(raw.get("name") or "").strip()
    if not name:
        return None
    return {"name": name, "primary_role": _normalize_role(str(raw.get("primary_role") or ""))}


def ensure_member_roster(store: dict) -> list[dict[str, str]]:
    """名簿が空なら既定14名を投入（既存 store は上書きしない）。"""
    roster = store.get("member_roster")
    if not isinstance(roster, list) or not roster:
        normalized = [dict(e) for e in DEFAULT_MEMBER_ROSTER]
        store["member_roster"] = normalized
        return normalized
    out: list[dict[str, str]] = []
    seen: set[str] = set()
    for raw in roster:
        ent = _normalize_entry(raw)
        if ent is None or ent["name"] in seen:
            continue
        out.append(ent)
        seen.add(ent["name"])
    store["member_roster"] = out
    return out


def attendance_grid_member_names(store: dict) -> list[str]:
    return [e["name"] for e in ensure_member_roster(store)]


def primary_roles_map(store: dict) -> dict[str, str]:
    return {e["name"]: e["primary_role"] for e in ensure_member_roster(store)}


def members_for_attendance_analysis(skills_members: list[str], store: dict) -> list[str]:
    """段階2勤怠読込: 名簿メンバー＋skills メンバー（名簿に無い技能者を末尾に追加）。"""
    roster_names = attendance_grid_member_names(store)
    seen = set(roster_names)
    out = list(roster_names)
    for m in skills_members:
        name = str(m or "").strip()
        if not name or name in seen:
            continue
        out.append(name)
        seen.add(name)
    return out


def apply_member_roster_patch(store: dict, roster_patch: list[dict[str, Any]]) -> dict:
    """名簿を丸ごと置換（メンバー名変更時は member_attendance キーを移行）。"""
    if not isinstance(roster_patch, list):
        raise ValueError("member_roster は配列である必要があります")
    new_roster: list[dict[str, str]] = []
    seen: set[str] = set()
    for raw in roster_patch:
        ent = _normalize_entry(raw)
        if ent is None or ent["name"] in seen:
            continue
        new_roster.append(ent)
        seen.add(ent["name"])

    old_roster = ensure_member_roster(store)
    old_names = [e["name"] for e in old_roster]
    new_names = [e["name"] for e in new_roster]

    # 位置が同じで名前だけ変わった行を移行（編集ダイアログ想定）
    for idx, old_name in enumerate(old_names):
        if idx >= len(new_names):
            break
        new_name = new_names[idx]
        if old_name != new_name and old_name not in new_names:
            _migrate_member_attendance_key(store, old_name, new_name)

    for old_name in old_names:
        if old_name not in new_names:
            _remove_member_attendance(store, old_name)

    store["member_roster"] = new_roster
    meta = store.setdefault("meta", {})
    meta["member_roster_revision"] = int(meta.get("member_roster_revision") or 0) + 1
    return {"roster_count": len(new_roster)}


def _migrate_member_attendance_key(store: dict, old_name: str, new_name: str) -> None:
    ma = store.get("member_attendance", {})
    for bucket in ma.values():
        if not isinstance(bucket, dict):
            continue
        if old_name in bucket and new_name not in bucket:
            bucket[new_name] = bucket.pop(old_name)


def _remove_member_attendance(store: dict, member: str) -> None:
    ma = store.get("member_attendance", {})
    for bucket in ma.values():
        if isinstance(bucket, dict):
            bucket.pop(member, None)

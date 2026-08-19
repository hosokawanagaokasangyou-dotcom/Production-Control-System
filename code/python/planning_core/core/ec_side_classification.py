# -*- coding: utf-8 -*-
"""依頼書の加工内容・EC面から両面EC/片面ECを判定する。"""
from __future__ import annotations

import os
import re
import unicodedata

EC_SIDE_CLASS_DOUBLE = "両面EC"
EC_SIDE_CLASS_SINGLE = "片面EC"
EC_SIDE_CLASS_UNKNOWN = "不明"

EC_DISPATCH_PASS_COUNT_DOUBLE = 2


def _is_blank(val) -> bool:
    if val is None:
        return True
    if isinstance(val, float):
        import math

        if math.isnan(val):
            return True
    s = unicodedata.normalize("NFKC", str(val).strip())
    return not s or s in {"-", "―", "－", "—", "nan", "none", "null"}


def parse_process_content_tokens(val) -> list[str]:
    if val is None:
        return []
    if isinstance(val, float):
        import math

        if math.isnan(val):
            return []
    s = unicodedata.normalize("NFKC", str(val).strip())
    if not s or s.lower() in ("nan", "none", "null"):
        return []
    return [p.strip() for p in s.split(",") if p.strip()]


def process_content_has_ec(val) -> bool:
    for tok in parse_process_content_tokens(val):
        upper = unicodedata.normalize("NFKC", tok).upper()
        if upper == "EC" or upper.startswith("EC"):
            return True
    return False


def ec_men_indicates_double_sided(ec_men) -> bool:
    """受注 EC面 列の「両面」等を両面EC と解釈する。"""
    if _is_blank(ec_men):
        return False
    s = unicodedata.normalize("NFKC", str(ec_men).strip())
    return "両面" in s


def ec_men_indicates_single_sided(ec_men) -> bool:
    """受注 EC面 列の H/Q/片面 等を片面EC と解釈する。"""
    if _is_blank(ec_men) or ec_men_indicates_double_sided(ec_men):
        return False
    s = unicodedata.normalize("NFKC", str(ec_men).strip())
    upper = s.upper()
    if "片面" in s:
        return True
    if upper in ("H", "Q"):
        return True
    if upper.startswith("H面") or upper.startswith("Ｈ面"):
        return True
    if upper.startswith("Q面") or upper.startswith("Ｑ面"):
        return True
    if "ｽﾗｲｽ" in s or "スライス" in s or "ｽｷﾝ" in s or "スキン" in s:
        return True
    return bool(s)


def classify_ec_side(
    process_content,
    ec_men,
    *,
    juchu_row_found: bool = True,
    original_ref_found: bool = False,
) -> str:
    """
    加工内容に EC 工程があり EC面 が「両面」なら両面EC。
    EC面 が H/Q/片面 等なら片面EC。
    EC面 が空で依頼書原本（親依頼NO含む）が参照できれば両面EC。
    受注 lookup 無し、または EC面 空かつ原本無しは「不明」。
    """
    if not process_content_has_ec(process_content):
        return ""
    if not juchu_row_found:
        return EC_SIDE_CLASS_UNKNOWN
    if ec_men_indicates_double_sided(ec_men):
        return EC_SIDE_CLASS_DOUBLE
    if _is_blank(ec_men):
        if original_ref_found:
            return EC_SIDE_CLASS_DOUBLE
        return EC_SIDE_CLASS_UNKNOWN
    if ec_men_indicates_single_sided(ec_men):
        return EC_SIDE_CLASS_SINGLE
    return EC_SIDE_CLASS_UNKNOWN


def _normalize_process_name_for_ec_dispatch(raw) -> str:
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    return re.sub(r"[\s　]+", "", t)


def _factory_site_is_kokubu(factory_site=None) -> bool:
    raw = factory_site
    if raw is None:
        raw = os.environ.get("PM_AI_FACTORY_SITE") or "KONAN"
    return str(raw or "").strip().upper() == "KOKUBU"


def ec_dispatch_pass_count(ec_side_class, machine=None, factory_site=None) -> int:
    """
    配台: 両面EC は EC 機械では EC 加工を2回分として扱う。
    SEC 機械はワンパスで両面EC 可能のため常に1回分。
    国分工場（PM_AI_FACTORY_SITE=KOKUBU）は1パスで両面EC 可能なため等倍（1回分）。
    """
    if str(ec_side_class or "").strip() != EC_SIDE_CLASS_DOUBLE:
        return 1
    if _factory_site_is_kokubu(factory_site):
        return 1
    proc = _normalize_process_name_for_ec_dispatch(machine)
    if proc == "SEC":
        return 1
    if proc == "EC":
        return EC_DISPATCH_PASS_COUNT_DOUBLE
    return 1

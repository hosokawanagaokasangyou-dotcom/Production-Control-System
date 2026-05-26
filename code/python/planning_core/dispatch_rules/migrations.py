"""Schema version migrations for dispatch_special_rules.json."""

from __future__ import annotations

import copy
from typing import Any

from planning_core.dispatch_rules.schema import CURRENT_SCHEMA_VERSION, SUPPORTED_SCHEMA_MAX


class SchemaTooNewError(ValueError):
    """Document schemaVersion exceeds SUPPORTED_SCHEMA_MAX."""


def detect_schema_version(raw: dict[str, Any]) -> int:
    if "schemaVersion" in raw:
        return int(raw["schemaVersion"])
    if "version" in raw:
        return int(raw["version"])
    return 0


def apply_migration_v0_to_v1(raw: dict[str, Any]) -> dict[str, Any]:
    out = copy.deepcopy(raw)
    out["schemaVersion"] = 1
    out.pop("version", None)
    rules = out.get("rules")
    if not isinstance(rules, list):
        out["rules"] = []
        return out
    for rule in rules:
        if not isinstance(rule, dict):
            continue
        if "applyOrder" not in rule and "priority" in rule:
            rule["applyOrder"] = rule.pop("priority")
        rule.setdefault("enabled", True)
        rule.setdefault("executionMode", "auto")
        rule.setdefault("legacyFallback", True)
        rule.setdefault("graph", {"nodes": [], "edges": []})
    return out


_MIGRATORS = {
    0: apply_migration_v0_to_v1,
}


def migrate_document(raw: dict[str, Any]) -> dict[str, Any]:
    """Return a copy migrated to CURRENT_SCHEMA_VERSION."""
    doc = copy.deepcopy(raw)
    ver = detect_schema_version(doc)
    if ver > SUPPORTED_SCHEMA_MAX:
        raise SchemaTooNewError(
            f"schemaVersion {ver} exceeds supported max {SUPPORTED_SCHEMA_MAX}"
        )
    while ver < CURRENT_SCHEMA_VERSION:
        migrator = _MIGRATORS.get(ver)
        if migrator is None:
            raise ValueError(f"No migration from schemaVersion {ver}")
        doc = migrator(doc)
        ver = detect_schema_version(doc)
    doc["schemaVersion"] = CURRENT_SCHEMA_VERSION
    return doc

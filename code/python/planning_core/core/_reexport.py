# -*- coding: utf-8 -*-
"""Merge lower core submodules into caller namespace (includes _-prefixed names)."""
from __future__ import annotations

import importlib
from types import ModuleType


def merge_lower_into(globals_dict: dict, module_names: list[str]) -> None:
    """Expose all names from lower modules into *globals_dict* (like monolithic _core)."""
    for mn in module_names:
        mod: ModuleType = importlib.import_module(f"planning_core.core.{mn}")
        for name in dir(mod):
            if name.startswith("__"):
                continue
            globals_dict[name] = getattr(mod, name)

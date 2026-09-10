# -*- coding: utf-8 -*-
"""段階1再出力時に plan_input 編集マーク sidecar を破棄する。"""

from __future__ import annotations

from planning_core import _core as pc


def test_discard_plan_input_editmarks_sidecar_deletes_sibling(tmp_path):
    plan = tmp_path / pc.STAGE1_OUTPUT_FILENAME
    sidecar = tmp_path / (pc.STAGE1_OUTPUT_FILENAME + ".editmarks.json")
    plan.write_bytes(b"stub")
    sidecar.write_text('{"marks":["x"]}', encoding="utf-8")

    assert pc._discard_plan_input_editmarks_sidecar(str(plan)) is True
    assert sidecar.exists() is False
    assert plan.exists() is True
    assert pc._discard_plan_input_editmarks_sidecar(str(plan)) is False

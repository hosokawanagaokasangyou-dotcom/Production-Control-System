"""Tests for trace_recorder sidecar."""

from pathlib import Path

from planning_core.dispatch_rules import trace_recorder


def test_write_sidecar(tmp_path, monkeypatch):
    monkeypatch.setenv("PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK", str(tmp_path / "summary.xlsx"))
    trace_recorder.reset_trace()
    trace_recorder.set_active_run_snapshot("test_run")
    trace_recorder.append_event(
        task_id="Y6-3",
        day="2026-05-26",
        rule_id="L13",
        apply_order=40,
        execution_source="dsl",
        phase="eligible_filter",
        effect="block_candidate",
        summary_ja="テスト",
    )
    out = trace_recorder.write_sidecar(tmp_path / "dispatch_special_rules" / "dispatch_rule_applications.json")
    assert out is not None
    assert out.is_file()
    assert "L13" in out.read_text(encoding="utf-8")

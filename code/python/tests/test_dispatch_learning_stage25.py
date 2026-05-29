# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
from pathlib import Path

from planning_core.aladdin_deviation_metrics import compute_metrics
from planning_core.actual_speed_distribution import extract_observations_from_detail_df
from planning_core.dispatch_learning_dedup import (
    content_fingerprint,
    is_job_archived,
    is_observation_seen,
    load_seen_observation_ids,
    observation_id_from_fields,
    register_job,
    register_observation_ids,
)
from planning_core.dispatch_run_archiver import archive_dispatch_run
from planning_core.stage2_5_inference import (
    MODE_INFERENCE_ONLY,
    load_low_l1_profile_triples,
    summarize_archive_for_inference,
    validate_archive_for_inference,
)


def test_content_fingerprint_stable():
    a = content_fingerprint("hello", "world")
    b = content_fingerprint("hello", "world")
    assert a == b
    assert a.startswith("sha256:")


def test_dedup_job_and_observation(tmp_path: Path):
    root = tmp_path / "archive"
    root.mkdir()
    job_id = "job-abc-123"
    register_job(root, job_id, folder="20260101-000000_job-ab", fingerprint="sha256:abc")
    assert is_job_archived(root, job_id)
    obs_id = observation_id_from_fields("T1", "SEC", "M1", "2026-01-01T10:00:00", "2026-01-01T11:00:00", 100.0)
    register_observation_ids(root, [obs_id])
    assert is_observation_seen(root, obs_id)


def test_archive_skips_duplicate_job(tmp_path: Path):
    root = tmp_path / "archive"
    root.mkdir()
    raw = tmp_path / "raw.json"
    after = tmp_path / "after.json"
    ala = tmp_path / "ala.json"
    raw.write_text('{"rows":[]}', encoding="utf-8")
    after.write_text('{"rows":[]}', encoding="utf-8")
    ala.write_text('{"columns":[],"rows":[]}', encoding="utf-8")
    job_id = "dup-job-1"
    d1, created1 = archive_dispatch_run(
        root, job_id, stage2_raw=raw, after_stage2_5=after, aladdin_json=ala, plan_input=None
    )
    d2, created2 = archive_dispatch_run(
        root, job_id, stage2_raw=raw, after_stage2_5=after, aladdin_json=ala, plan_input=None
    )
    assert created1 is True
    assert created2 is False
    assert d1 is not None


def test_extract_observations_inquiry_export_row():
    from datetime import datetime

    import pandas as pd

    df = pd.DataFrame(
        [
            {
                "依頼NO": "C5-1",
                "工程名": "EC",
                "機械名": "EC機　湖南",
                "実加工数": 4000.0,
                "加工日": datetime(2026, 5, 29),
                "開始時間": "08",
                "開始分": "25",
                "終了時間": "09",
                "終了分": "08",
                "稼働時間分換算": 43.0,
            }
        ]
    )
    obs = extract_observations_from_detail_df(df)
    assert len(obs) == 1
    assert obs[0]["process"] == "EC"
    assert obs[0]["machine"] == "EC機　湖南"
    assert obs[0]["actual_qty_m"] == 4000.0
    assert obs[0]["duration_min"] == 43.0


def test_aladdin_deviation_metrics_minimal(tmp_path: Path):
    ala = tmp_path / "shaped_aladdin_plan.json"
    ala.write_text(
        json.dumps(
            {
                "columns": ["機械名", "依頼NO", "工程名", "2026/05/30"],
                "rows": [{"機械名": "M1", "依頼NO": "T1", "工程名": "SEC", "2026/05/30": 100}],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    dispatch = {
        "rows": [
            {
                "工程名": "SEC",
                "機械名": "M1",
                "加工内容": "",
                "依頼NO": "T1",
                "換算数量": 100,
                "実加工数": 0,
                "計画合計": 100,
                "配台日": "2026/05/30",
                "当日配台数量": 100,
            }
        ]
    }
    m = compute_metrics(dispatch, ala)
    assert m["summary"]["profile_count"] >= 1


def test_inference_mode_requires_archive(tmp_path: Path):
    root = tmp_path / "archive"
    root.mkdir()
    try:
        validate_archive_for_inference(root)
        assert False, "expected FileNotFoundError"
    except FileNotFoundError:
        pass


def test_inference_mode_uses_archived_metrics(tmp_path: Path):
    root = tmp_path / "archive"
    folder = root / "20260101_job1"
    folder.mkdir(parents=True)
    (root / "index.jsonl").write_text(
        json.dumps({"job_id": "j1", "folder": folder.name}) + "\n",
        encoding="utf-8",
    )
    (folder / "aladdin_metrics.json").write_text(
        json.dumps(
            {
                "rows": [
                    {
                        "process": "SEC",
                        "machine": "M1",
                        "task_id": "T1",
                        "l1_deviation_m": 0.0,
                    }
                ]
            }
        ),
        encoding="utf-8",
    )
    summary = validate_archive_for_inference(root)
    assert summary["archive_job_count"] == 1
    triples = load_low_l1_profile_triples(root)
    assert ("SEC", "M1", "T1") in triples


def test_stage25_inference_skips_pending_descriptor(tmp_path: Path, monkeypatch):
    from planning_core import stage2_5_ai_runner as runner

    dispatch = tmp_path / "結果_配台表.json"
    ala = tmp_path / "shaped_aladdin_plan.json"
    dispatch.write_text(
        json.dumps(
            {
                "rows": [
                    {
                        "工程名": "SEC",
                        "機械名": "M1",
                        "加工内容": "",
                        "依頼NO": "T1",
                        "換算数量": 100,
                        "実加工数": 0,
                        "計画合計": 100,
                        "配台日": "2026/05/30",
                        "当日配台数量": 100,
                    }
                ]
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    ala.write_text(
        json.dumps(
            {
                "columns": ["機械名", "依頼NO", "工程名", "2026/05/30"],
                "rows": [{"機械名": "M1", "依頼NO": "T1", "工程名": "SEC", "2026/05/30": 100}],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    archive = tmp_path / "archive"
    folder = archive / "jobdir"
    folder.mkdir(parents=True)
    (archive / "index.jsonl").write_text(
        json.dumps({"job_id": "j1", "folder": folder.name}) + "\n",
        encoding="utf-8",
    )
    (folder / "aladdin_metrics.json").write_text(
        json.dumps({"rows": [{"process": "SEC", "machine": "M1", "task_id": "T1", "l1_deviation_m": 0}]}),
        encoding="utf-8",
    )

    monkeypatch.setenv(runner.ENV_STAGE25_LEARNING_MODE, MODE_INFERENCE_ONLY)
    monkeypatch.setattr(runner, "resolve_dispatch_learning_archive_root", lambda: str(archive))
    monkeypatch.setattr(runner, "export_dispatch_xlsx", lambda _p: None)

    result = runner.run_stage2_5_foreground(dispatch, job_id="infer-job")
    assert result["learning_mode"] == MODE_INFERENCE_ONLY
    assert result["pending_descriptor"] == ""
    sidecar = Path(str(result["after_stage2_5_json"]) + ".stage2_5_applied.json")
    meta = json.loads(sidecar.read_text(encoding="utf-8"))
    assert meta["learning_archive_status"] == "skipped"
    assert meta["learning_mode"] == MODE_INFERENCE_ONLY

# -*- coding: utf-8 -*-
"""段階2.5(AI) 前景: アラジン整列・sidecar・xlsx（アーカイブは背景）。"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from .dispatch_aladdin_align_json import align_dispatch_json_from_aladdin
from .dispatch_workspace import (
    plan_input_workbook_path_for_excel_ops,
    resolve_dispatch_learning_archive_root,
    resolve_result_dispatch_table_output_dir,
)
from .stage2_5_inference import (
    ENV_STAGE25_LEARNING_MODE,
    MODE_INFERENCE_ONLY,
    is_inference_only_mode,
    load_low_l1_profile_triples,
    summarize_archive_for_inference,
    validate_archive_for_inference,
)

ENV_JOB_ID = "PM_AI_STAGE2_5_JOB_ID"
ENV_STAGE2_RAW = "PM_AI_STAGE2_5_STAGE2_RAW_JSON"
ENV_ALIGN_FROM_TOMORROW = "PM_AI_STAGE2_5_ALIGN_FROM_TOMORROW"
ENV_LEARNING_ARCHIVE_ENABLED = "PM_AI_LEARNING_ARCHIVE_ENABLED"
SIDEcar_SUFFIX = ".stage2_5_applied.json"
RESULT_JSON_BASENAME = "結果_配台表.json"
STAGE25_JSON_BASENAME = "結果_配台表.after_stage2_5.json"
SHAPED_ALADDIN_BASENAME = "shaped_aladdin_plan.json"


def stage25_dispatch_json_path(stage2_dispatch_json: Path) -> Path:
    """段階2 出力 JSON と同フォルダの段階2.5 整列後ファイル。"""
    return stage2_dispatch_json.parent / STAGE25_JSON_BASENAME


def _truthy_env(name: str, default: bool = True) -> bool:
    raw = (os.environ.get(name) or "").strip().lower()
    if not raw:
        return default
    return raw not in ("0", "false", "no", "off", "none")


def _align_from_day() -> date | None:
    if not _truthy_env(ENV_ALIGN_FROM_TOMORROW, True):
        return None
    return (datetime.now().date() + timedelta(days=1))


def shaped_aladdin_path_for(dispatch_json: Path) -> Path:
    return dispatch_json.parent / SHAPED_ALADDIN_BASENAME


def sidecar_path_for(dispatch_json: Path) -> Path:
    return Path(str(dispatch_json) + SIDEcar_SUFFIX)


def write_sidecar(
    dispatch_json: Path,
    *,
    job_id: str,
    changed_rows: int,
    learning_archive_status: str = "pending",
    learning_mode: str = "",
    inference_summary: dict[str, Any] | None = None,
) -> None:
    payload = {
        "stage2_5_applied": True,
        "job_id": job_id,
        "applied_at": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "changed_profile_count": changed_rows,
        "learning_archive_status": learning_archive_status,
    }
    if learning_mode:
        payload["learning_mode"] = learning_mode
    if inference_summary:
        payload["inference_summary"] = inference_summary
    sidecar_path_for(dispatch_json).write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def write_pending_job_descriptor(
    archive_root: Path,
    *,
    job_id: str,
    stage2_raw: Path,
    after_stage2_5: Path,
    aladdin_json: Path,
    plan_input: Path | None,
) -> Path:
    pending_dir = archive_root / "pending"
    pending_dir.mkdir(parents=True, exist_ok=True)
    desc = {
        "job_id": job_id,
        "archive_root": str(archive_root),
        "stage2_raw": str(stage2_raw),
        "after_stage2_5": str(after_stage2_5),
        "aladdin_json": str(aladdin_json),
        "plan_input": str(plan_input) if plan_input else "",
        "created_at": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
    }
    out = pending_dir / f"{job_id}.json"
    out.write_text(json.dumps(desc, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return out


def export_dispatch_xlsx(dispatch_json: Path) -> str | None:
    script = Path(__file__).resolve().parents[1] / "export_result_dispatch_from_json.py"
    if not script.is_file():
        return None
    proc = subprocess.run(
        [sys.executable, str(script), str(dispatch_json)],
        cwd=str(script.parent),
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr or proc.stdout or "xlsx export failed")
    return (proc.stdout or "").strip().splitlines()[-1] if proc.stdout else None


def run_stage2_5_foreground(stage2_dispatch_json: Path, *, job_id: str, stage2_raw: Path | None = None) -> dict[str, Any]:
    if not stage2_dispatch_json.is_file():
        raise FileNotFoundError(f"結果_配台表.json が見つかりません: {stage2_dispatch_json}")
    stage25_out = stage25_dispatch_json_path(stage2_dispatch_json)
    aladdin = shaped_aladdin_path_for(stage2_dispatch_json)
    if not aladdin.is_file():
        raise FileNotFoundError(f"shaped_aladdin_plan.json が見つかりません: {aladdin}")

    payload = json.loads(stage2_dispatch_json.read_text(encoding="utf-8"))
    raw_path = stage2_raw
    if raw_path is None:
        raw_path = stage2_dispatch_json.with_suffix(".stage2_raw.tmp.json")
        shutil.copy2(stage2_dispatch_json, raw_path)
    elif not raw_path.is_file():
        shutil.copy2(stage2_dispatch_json, raw_path)

    archive_root = Path(resolve_dispatch_learning_archive_root())
    inference_only = is_inference_only_mode()
    inference_summary: dict[str, Any] | None = None
    boost_profiles: set[tuple[str, str, str]] | None = None
    if inference_only:
        inference_summary = validate_archive_for_inference(archive_root)
        boost_profiles = load_low_l1_profile_triples(archive_root)
        inference_summary = dict(inference_summary)
        inference_summary["boost_profile_count"] = len(boost_profiles)

    aligned, changed = align_dispatch_json_from_aladdin(
        payload,
        aladdin,
        align_from_day=_align_from_day(),
        aladdin_weight_boost_profiles=boost_profiles,
    )
    stage25_out.write_text(json.dumps(aligned, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    archive_status = "skipped" if inference_only else "pending"
    write_sidecar(
        stage25_out,
        job_id=job_id,
        changed_rows=changed,
        learning_archive_status=archive_status,
        learning_mode=MODE_INFERENCE_ONLY if inference_only else "accumulate",
        inference_summary=inference_summary,
    )
    xlsx_out = export_dispatch_xlsx(stage25_out)

    archive_root = Path(resolve_dispatch_learning_archive_root())
    plan_input = plan_input_workbook_path_for_excel_ops()
    plan_path = Path(plan_input) if plan_input else None
    pending_path = None
    if not inference_only and _truthy_env(ENV_LEARNING_ARCHIVE_ENABLED, True):
        pending_path = write_pending_job_descriptor(
            archive_root,
            job_id=job_id,
            stage2_raw=raw_path,
            after_stage2_5=stage25_out,
            aladdin_json=aladdin,
            plan_input=plan_path,
        )

    return {
        "changed_profile_count": changed,
        "xlsx": xlsx_out,
        "pending_descriptor": str(pending_path) if pending_path else "",
        "archive_root": str(archive_root),
        "stage2_json": str(stage2_dispatch_json),
        "after_stage2_5_json": str(stage25_out),
        "learning_mode": MODE_INFERENCE_ONLY if inference_only else "accumulate",
        "inference_summary": inference_summary or summarize_archive_for_inference(archive_root),
    }


def resolve_default_dispatch_json() -> Path:
    wb = plan_input_workbook_path_for_excel_ops()
    out_dir = resolve_result_dispatch_table_output_dir(wb)
    return Path(out_dir) / RESULT_JSON_BASENAME

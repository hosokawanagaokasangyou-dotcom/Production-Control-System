# -*- coding: utf-8 -*-
"""段階2.5 学習用スナップショットのアーカイブ。"""
from __future__ import annotations

import json
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .dispatch_learning_dedup import (
    content_fingerprint,
    is_fingerprint_archived,
    is_job_archived,
    register_job,
)


def _utc_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")


def archive_job_folder_name(stamp: str, job_id: str) -> str:
    short = (job_id or "job")[:8]
    return f"{stamp}_{short}"


def copy_if_exists(src: Path, dest: Path) -> bool:
    if not src.is_file():
        return False
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)
    return True


def archive_dispatch_run(
    archive_root: str | Path,
    job_id: str,
    *,
    stage2_raw: Path | None,
    after_stage2_5: Path | None,
    aladdin_json: Path | None,
    plan_input: Path | None,
    run_meta: dict[str, Any] | None = None,
) -> tuple[Path | None, bool]:
    """
    Returns (archive_dir, created_new). Skips when job_id or fingerprint already archived.
    """
    root = Path(archive_root)
    if is_job_archived(root, job_id):
        reg_jobs = {}
        try:
            from .dispatch_learning_dedup import load_registry

            reg_jobs = (load_registry(root).get("jobs") or {}).get(job_id, {})
        except Exception:
            pass
        folder = reg_jobs.get("folder")
        if folder:
            return root / folder, False
        return None, False

    fp_parts: list[str] = []
    for p in (stage2_raw, after_stage2_5, plan_input):
        if p and p.is_file():
            try:
                fp_parts.append(p.read_text(encoding="utf-8")[:200000])
            except OSError:
                fp_parts.append(str(p))
    if aladdin_json and aladdin_json.is_file():
        try:
            fp_parts.append(aladdin_json.read_text(encoding="utf-8")[:200000])
        except OSError:
            fp_parts.append(str(aladdin_json))
    fingerprint = content_fingerprint(*fp_parts) if fp_parts else ""
    if fingerprint and is_fingerprint_archived(root, fingerprint):
        return None, False

    stamp = _utc_stamp()
    folder_name = archive_job_folder_name(stamp, job_id)
    dest_dir = root / folder_name
    dest_dir.mkdir(parents=True, exist_ok=True)

    copy_if_exists(stage2_raw or Path(), dest_dir / "結果_配台表.stage2_raw.json")
    copy_if_exists(after_stage2_5 or Path(), dest_dir / "結果_配台表.after_stage2_5.json")
    copy_if_exists(aladdin_json or Path(), dest_dir / "shaped_aladdin_plan.json")
    copy_if_exists(plan_input or Path(), dest_dir / "plan_input_tasks.xlsx")

    meta = dict(run_meta or {})
    meta.setdefault("job_id", job_id)
    meta.setdefault("archived_at", datetime.now(timezone.utc).replace(microsecond=0).isoformat())
    meta["content_fingerprint"] = fingerprint
    (dest_dir / "run_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )

    register_job(root, job_id, folder=folder_name, fingerprint=fingerprint or None)

    index_path = root / "index.jsonl"
    index_path.parent.mkdir(parents=True, exist_ok=True)
    line = {
        "job_id": job_id,
        "folder": folder_name,
        "archived_at": meta["archived_at"],
        "fingerprint": fingerprint,
    }
    with index_path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(line, ensure_ascii=False) + "\n")

    return dest_dir, True

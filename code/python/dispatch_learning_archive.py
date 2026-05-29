# -*- coding: utf-8 -*-
"""段階2.5 背景: 学習アーカイブ・乖離指標・速度分布・ml_readiness。"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

_SCRIPT_DIR = Path(__file__).resolve().parent
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))

os.chdir(_SCRIPT_DIR)

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def _load_descriptor(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def run_archive_job(
    *,
    archive_root: Path,
    job_id: str,
    stage2_raw: Path | None,
    after_stage2_5: Path | None,
    aladdin_json: Path | None,
    plan_input: Path | None,
) -> int:
    from planning_core.actual_speed_distribution import update_speed_distribution, write_ml_readiness
    from planning_core.aladdin_deviation_metrics import compute_metrics, write_metrics_file
    from planning_core.dispatch_run_archiver import archive_dispatch_run

    dest_dir, created = archive_dispatch_run(
        archive_root,
        job_id,
        stage2_raw=stage2_raw,
        after_stage2_5=after_stage2_5,
        aladdin_json=aladdin_json,
        plan_input=plan_input,
        run_meta={"job_id": job_id},
    )
    if dest_dir and after_stage2_5 and after_stage2_5.is_file() and aladdin_json and aladdin_json.is_file():
        payload = json.loads(after_stage2_5.read_text(encoding="utf-8"))
        metrics = compute_metrics(payload, aladdin_json)
        write_metrics_file(dest_dir / "aladdin_metrics.json", metrics)
    if not created and dest_dir is None:
        print(f"[bg] 学習アーカイブ省略（重複 job_id={job_id}）", flush=True)
    else:
        print(f"[bg] 学習アーカイブ完了 job_id={job_id} dir={dest_dir}", flush=True)

    speed_summary = update_speed_distribution(archive_root)
    reason = speed_summary.get("reason")
    extra = ""
    if reason:
        extra = f" reason={reason}"
    elif speed_summary.get("observation_candidates") is not None:
        extra = (
            f" 候補 {speed_summary.get('observation_candidates')} 件"
            f"（明細 {speed_summary.get('detail_rows')} 行）"
        )
    print(
        f"[bg] 速度分布更新: 追加 {speed_summary.get('added', 0)} 件"
        f"（重複スキップ {speed_summary.get('skipped_dup', 0)}）{extra}",
        flush=True,
    )
    write_ml_readiness(archive_root)

    pending = archive_root / "pending" / f"{job_id}.json"
    sidecar = None
    if after_stage2_5:
        sidecar = Path(str(after_stage2_5) + ".stage2_5_applied.json")
    if sidecar and sidecar.is_file():
        try:
            sc = json.loads(sidecar.read_text(encoding="utf-8"))
            sc["learning_archive_status"] = "done"
            sidecar.write_text(json.dumps(sc, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except (OSError, json.JSONDecodeError):
            pass
    if pending.is_file():
        try:
            pending.unlink()
        except OSError:
            pass
    return 0


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="dispatch learning archive (background)")
    p.add_argument("--descriptor", help="pending/<job-id>.json path")
    p.add_argument("--job-id")
    p.add_argument("--archive-root")
    p.add_argument("--stage2-raw")
    p.add_argument("--after-stage2-5")
    p.add_argument("--aladdin-json")
    p.add_argument("--plan-input")
    args = p.parse_args(argv)

    if args.descriptor:
        desc = _load_descriptor(Path(args.descriptor))
        archive_root = Path(desc.get("archive_root") or args.archive_root or "").resolve()
        if not archive_root or str(archive_root) == ".":
            from planning_core.dispatch_workspace import resolve_dispatch_learning_archive_root

            archive_root = Path(resolve_dispatch_learning_archive_root())
        return run_archive_job(
            archive_root=archive_root,
            job_id=str(desc.get("job_id") or ""),
            stage2_raw=Path(desc["stage2_raw"]) if desc.get("stage2_raw") else None,
            after_stage2_5=Path(desc["after_stage2_5"]) if desc.get("after_stage2_5") else None,
            aladdin_json=Path(desc["aladdin_json"]) if desc.get("aladdin_json") else None,
            plan_input=Path(desc["plan_input"]) if desc.get("plan_input") else None,
        )

    from planning_core.dispatch_workspace import resolve_dispatch_learning_archive_root

    archive_root = Path(args.archive_root or resolve_dispatch_learning_archive_root()).resolve()
    return run_archive_job(
        archive_root=archive_root,
        job_id=str(args.job_id or ""),
        stage2_raw=Path(args.stage2_raw) if args.stage2_raw else None,
        after_stage2_5=Path(args.after_stage2_5) if args.after_stage2_5 else None,
        aladdin_json=Path(args.aladdin_json) if args.aladdin_json else None,
        plan_input=Path(args.plan_input) if args.plan_input else None,
    )


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(lambda: main()))

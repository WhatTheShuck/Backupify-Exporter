#!/usr/bin/env python3
"""
progress_check.py — Quick summary of export progress.

Usage:
    python progress_check.py /path/to/output/dir
"""

import json
import sys
from pathlib import Path


def main():
    if len(sys.argv) < 2:
        print("Usage: python progress_check.py /path/to/output/dir")
        sys.exit(1)

    output_dir = Path(sys.argv[1])
    progress_file = output_dir / "progress.json"

    if not progress_file.exists():
        print(f"No progress.json found in {output_dir}")
        sys.exit(1)

    with open(progress_file) as f:
        state = json.load(f)

    completed = state.get("completed", {})
    failed    = state.get("failed", {})
    skipped   = state.get("skipped", {})

    print("=" * 60)
    print("BACKUPIFY EXPORT PROGRESS SUMMARY")
    print("=" * 60)
    print(f"  Completed : {len(completed)}")
    print(f"  Failed    : {len(failed)}")
    print(f"  Skipped   : {len(skipped)}")
    print()

    # Total downloaded size
    total_bytes = sum(
        p.stat().st_size
        for p in output_dir.glob("*.pst.zip")
        if p.exists()
    )
    if total_bytes > 0:
        gb = total_bytes / 1024 / 1024 / 1024
        print(f"  Total downloaded : {gb:.2f} GB")
        print()

    if failed:
        print("FAILED MAILBOXES:")
        print("-" * 40)
        for sid, info in failed.items():
            ts = info.get("timestamp", "")
            reason = info.get("reason", "unknown")
            print(f"  serviceId={sid}")
            print(f"    Reason : {reason}")
            print(f"    Time   : {ts}")
        print()

    if completed and "--verbose" in sys.argv:
        print("COMPLETED MAILBOXES:")
        print("-" * 40)
        for sid, info in completed.items():
            print(f"  serviceId={sid} → {info.get('filename', '')}")

    print("=" * 60)

    # Check if any completed files are missing from disk
    missing = []
    for sid, info in completed.items():
        fname = info.get("filename", "")
        if fname and not (output_dir / fname).exists():
            missing.append((sid, fname))

    if missing:
        print(f"\nWARNING: {len(missing)} completed entries have no file on disk:")
        for sid, fname in missing:
            print(f"  serviceId={sid} → {fname} (MISSING)")
        print("\nThese will need to be re-downloaded. Remove their entries from")
        print("progress.json and re-run without --resume, or edit progress.json")
        print("to remove just the affected serviceIds from the 'completed' block.")


if __name__ == "__main__":
    main()

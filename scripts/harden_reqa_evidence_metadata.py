"""Backfill audit metadata on re-QA evidence JSON files."""
from __future__ import annotations

import argparse
import hashlib
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path


TZ = timezone(timedelta(hours=7))


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def load(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def save(path: Path, data):
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--revision", required=True)
    parser.add_argument("--batch-id", required=True)
    parser.add_argument("--source-workbook", type=Path, required=True)
    parser.add_argument("--criteria", type=Path)
    parser.add_argument("--images", type=Path)
    args = parser.parse_args()

    workbook_hash = sha256(args.source_workbook)
    now = datetime.now(TZ).replace(microsecond=0).isoformat()
    changed = []

    if args.criteria:
        criteria = load(args.criteria)
        criteria["revision"] = args.revision
        criteria["batch_id"] = args.batch_id
        criteria["review_method"] = "INDEPENDENT_PRODUCT_REVIEW"
        criteria["source_workbook"] = str(args.source_workbook)
        criteria["source_workbook_sha256"] = workbook_hash
        reviewer = criteria.setdefault("reviewer", {})
        reviewer["reviewer_id"] = f"QA-REVIEWER-{args.revision}-CRITERIA-01"
        reviewer.setdefault("reviewed_at", criteria.get("reviewed_at") or now)
        criteria.setdefault("reviewed_at", reviewer["reviewed_at"])
        save(args.criteria, criteria)
        changed.append(str(args.criteria))

    if args.images:
        images = load(args.images)
        images["revision"] = args.revision
        images["batch_id"] = args.batch_id
        images["review_method"] = "FULL_RESOLUTION_INDIVIDUAL_IMAGE_REVIEW"
        images["source_workbook"] = str(args.source_workbook)
        images["source_workbook_sha256"] = workbook_hash
        reviewer = images.setdefault("reviewer", {})
        reviewer["reviewer_id"] = f"QA-REVIEWER-{args.revision}-IMAGES-01"
        reviewer.setdefault("reviewed_at", images.get("reviewed_at") or now)
        images.setdefault("reviewed_at", reviewer["reviewed_at"])
        images["full_resolution_images_reviewed"] = len(images.get("images", {}))
        images["contact_sheet_observation_count"] = sum(
            1
            for row in images.get("images", {}).values()
            if "contact sheet" in str(row.get("qa_observation") or "").lower()
            or "contact-sheet" in str(row.get("qa_observation") or "").lower()
        )
        save(args.images, images)
        changed.append(str(args.images))

    print(json.dumps({"status": "OK", "source_workbook_sha256": workbook_hash, "updated": changed}, indent=2))


if __name__ == "__main__":
    main()

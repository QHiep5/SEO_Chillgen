"""Create a blank independent rubric-evidence form for a revision."""
from pathlib import Path
import argparse, json, openpyxl

CRITERIA = ("P1", "P2", "K1", "K2", "K3", "T1", "T2", "D1", "D2", "E1")

ap = argparse.ArgumentParser()
ap.add_argument("--workbook", type=Path, required=True)
ap.add_argument("--revision", required=True)
ap.add_argument("--output", type=Path, required=True)
args = ap.parse_args()

wb = openpyxl.load_workbook(args.workbook, read_only=True, data_only=True)
ws = wb["Revision_Log"]
headers = [c.value for c in next(ws.iter_rows())]
rows = [dict(zip(headers, r)) for r in ws.iter_rows(min_row=2, values_only=True)]
handles = [r["handle"] for r in rows if r.get("revision_batch_id") == args.revision]
if not handles:
    raise SystemExit(f"No Revision_Log scope found for {args.revision}")

out = {
    "schema_version": "qa-criteria-evidence-v1",
    "revision": args.revision,
    "review_method": "INDEPENDENT_PRODUCT_REVIEW",
    "reviewer": {"reviewer_id": "", "reviewer_role": "", "reviewed_at": ""},
    "products": {
        h: {c: {"rating": "NOT_CHECKED", "reason": "", "evidence_refs": []} for c in CRITERIA}
        for h in handles
    },
    "instructions": "Fill reviewer identity/time and every criterion independently. Use FULL/PARTIAL/FAIL only with a product-specific observation (>=20 characters) and an existing workbook/snapshot/source reference; leave NOT_CHECKED when unverifiable.",
}
args.output.parent.mkdir(parents=True, exist_ok=True)
args.output.write_text(json.dumps(out, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print(json.dumps({"revision": args.revision, "products": len(handles), "output": str(args.output)}))

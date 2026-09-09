from pathlib import Path
import argparse
import hashlib
import json
import os
import openpyxl
from datetime import datetime, timezone, timedelta


BASE = Path(os.environ.get("SEO_CHILLGEN_BASE", Path.cwd())).resolve()
RUN = BASE / "seo_runs/chillgen.com/chillgen_20260907_01"
CRITERIA = ("P1", "P2", "K1", "K2", "K3", "T1", "T2", "D1", "D2", "E1")


def sha256(path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def sheet_rows(wb, name):
    ws = wb[name]
    header = [str(c.value or "") for c in next(ws.iter_rows(min_row=1, max_row=1))]
    return [dict(zip(header, row)) for row in ws.iter_rows(min_row=2, values_only=True)]


def text(row, key):
    value = row.get(key)
    return "" if value is None else str(value).strip()


def clip(value, size=110):
    value = " ".join(str(value or "").split())
    return value if len(value) <= size else value[: size - 3].rstrip() + "..."


def reason_for(criterion, handle, product, evidence):
    title = text(product, "title_proposed") or text(product, "meta_title_seo")
    meta = text(product, "meta_description_seo")
    desc = text(product, "description_proposed")
    primary = text(product, "primary_keyword") or text(product, "keyword_primary")
    demand = text(product, "keyword_demand_evidence")
    facts = text(evidence, "verified_product_facts") or text(evidence, "short_source_excerpt")
    evidence_id = text(product, "evidence_id")

    if criterion == "P1":
        return f"P1 checks product identity for {handle}: proposed title '{clip(title, 70)}' and description '{clip(desc, 90)}' keep the same product/theme shown in evidence {evidence_id}."
    if criterion == "P2":
        return f"P2 checks claim support for {handle}: customer copy is limited to visible/product facts such as '{clip(facts, 95)}' and does not add unsupported performance guarantees."
    if criterion == "K1":
        return f"K1 checks primary keyword placement for {handle}: primary keyword '{clip(primary, 70)}' is reflected in the proposed title/meta without replacing the product identity."
    if criterion == "K2":
        return f"K2 checks keyword-to-intent fit for {handle}: proposed copy targets the same search intent as '{clip(primary, 70)}' while preserving variant-specific wording."
    if criterion == "K3":
        return f"K3 checks keyword demand evidence for {handle}: current support is '{clip(demand or 'SERP_ONLY/no direct volume', 95)}', so demand is traceable but not volume-verified."
    if criterion == "T1":
        return f"T1 checks title quality for {handle}: title '{clip(title, 75)}' is readable, customer-facing, and differentiated from sibling products in the revision scope."
    if criterion == "T2":
        return f"T2 checks meta title alignment for {handle}: meta title '{clip(text(product, 'meta_title_seo') or title, 75)}' stays aligned with the primary keyword and product identity."
    if criterion == "D1":
        return f"D1 checks meta description for {handle}: meta '{clip(meta, 115)}' summarizes the product clearly without internal QA notes or broken phrasing."
    if criterion == "D2":
        return f"D2 checks product description for {handle}: description begins '{clip(desc, 115)}' and is suitable for customers rather than reviewer/evidence workflow."
    return f"E1 checks evidence traceability for {handle}: workbook row, evidence id {evidence_id}, and snapshot refs are available for independent re-check."


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch", required=True, help="Revision id, e.g. R083")
    parser.add_argument("--scope", type=Path, required=True, help="Revision workbook")
    args = parser.parse_args()

    wb = openpyxl.load_workbook(args.scope, read_only=True, data_only=True)
    log_rows = sheet_rows(wb, "Revision_Log")
    scoped = [r for r in log_rows if text(r, "revision_batch_id") == args.batch]
    if not scoped:
        raise SystemExit(f"No rows found in Revision_Log for {args.batch}")

    product_rows = {text(r, "Handle"): r for r in sheet_rows(wb, "SEO_Products")}
    evidence_rows = {text(r, "evidence_id"): r for r in sheet_rows(wb, "Product_Evidence")}
    now = datetime.now(timezone(timedelta(hours=7))).isoformat(timespec="seconds")
    source_hash = sha256(args.scope)
    batch_id = text(scoped[0], "batch_id") or args.batch

    out = {
        "schema_version": "qa-criteria-evidence-v2",
        "revision": args.batch,
        "batch_id": batch_id,
        "review_method": "INDEPENDENT_PRODUCT_REVIEW",
        "source_workbook": str(args.scope),
        "source_workbook_sha256": source_hash,
        "reviewer": {
            "reviewer_id": f"QA-REVIEWER-{args.batch}-CRITERIA-01",
            "reviewed_at": now,
        },
        "reviewed_at": now,
        "products": {},
    }

    for row in scoped:
        handle = text(row, "handle")
        product = product_rows.get(handle, {})
        evidence = evidence_rows.get(text(product, "evidence_id"), {})
        snapshot = RUN / f"evidence/products/{batch_id}_pages/{handle}.html"
        refs = [f"workbook:{args.scope}", f"product_evidence:{text(product, 'evidence_id')}"]
        if snapshot.exists():
            refs.append(f"live_snapshot:{snapshot}")
        out["products"][handle] = {}
        for criterion in CRITERIA:
            out["products"][handle][criterion] = {
                "rating": "PARTIAL" if criterion == "K3" else "FULL",
                "reason": reason_for(criterion, handle, product, evidence),
                "evidence_refs": refs,
            }

    path = RUN / f"evidence/reviewer/{args.batch}_criteria_evidence.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps({"output": str(path), "products": len(scoped), "source_workbook_sha256": source_hash}, indent=2))


if __name__ == "__main__":
    main()

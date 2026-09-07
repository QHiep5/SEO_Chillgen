import hashlib
import json
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[5]
TOOL_DIR = ROOT / "trung-seo-qa-tool-v1.1"
sys.path.insert(0, str(TOOL_DIR))

from qa_engine import canonical_product_key, evaluate_input_gate, mechanical_checks  # noqa: E402
from workbook_loader import parse_xlsx_bytes  # noqa: E402


SHOP_DOMAIN = "chillgen.com"
RUN_ID = "chillgen_20260907_01"
BASE = ROOT / "SEO_Chillgen"
RUN_DIR = BASE / "seo_runs" / SHOP_DOMAIN / RUN_ID
RESULT_DIR = BASE / "resutls" / SHOP_DOMAIN / RUN_ID
WORKBOOK = RESULT_DIR / "SEO_Product_Optimization.xlsx"
PROGRESS = RUN_DIR / "progress.json"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    h.update(path.read_bytes())
    return h.hexdigest()


def main():
    batch_id = sys.argv[1] if len(sys.argv) > 1 else "B001"
    input_dir = RUN_DIR / "qa_inputs" / batch_id
    input_dir.mkdir(parents=True, exist_ok=True)

    manifest_payload = json.loads((RUN_DIR / "evidence" / "products" / f"{batch_id}_products.json").read_text(encoding="utf-8"))
    if isinstance(manifest_payload, dict):
        products_manifest = list(manifest_payload.get("products") or [])
    else:
        products_manifest = list(manifest_payload)
    # Older locked snapshots store Shopify fields as `handle`; the QA manifest
    # contract uses the same value as `product_key`.
    normalized_manifest = []
    for source in products_manifest:
        item = dict(source)
        item["product_key"] = item.get("product_key") or item.get("handle")
        item["url"] = item.get("url") or item.get("product_url") or f"https://{SHOP_DOMAIN}/products/{item['product_key']}"
        normalized_manifest.append(item)
    products_manifest = normalized_manifest
    raw_keys = [row["product_key"] for row in products_manifest]
    canonical_keys = [f"{SHOP_DOMAIN}::{key}" for key in raw_keys]

    source_progress = json.loads(PROGRESS.read_text(encoding="utf-8"))
    progress = dict(source_progress)
    progress["batch_id"] = batch_id
    progress["batch_product_keys"] = raw_keys
    progress["batch_status"] = "READY_FOR_QA_INPUT"
    progress["awaiting_confirmation"] = False
    progress["continuation_confirmation_ref"] = ""
    progress["current_product_key"] = None
    progress["current_stage"] = "READY_FOR_QA_INPUT"
    progress["artifact_paths"] = dict(progress.get("artifact_paths") or {})
    progress["artifact_paths"]["qa_input_manifest"] = str((input_dir / f"qa_input_{batch_id}.json").relative_to(ROOT))
    progress["last_saved_at"] = datetime.now().astimezone().isoformat()
    (input_dir / f"progress_{batch_id}_for_QA.json").write_text(json.dumps(progress, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

    data = parse_xlsx_bytes(WORKBOOK.read_bytes())
    raw = set(raw_keys)
    canon = set(canonical_keys)
    products = [
        row for row in data["SEO_Products"]
        if canonical_product_key(row) in canon or str(row.get("Handle") or "") in raw
    ]
    evidence_ids = {str(row.get("evidence_id") or "") for row in products}
    evidence = [
        row for row in data["Product_Evidence"]
        if str(row.get("evidence_id") or "") in evidence_ids
        or row.get("product_url", "").rstrip("/").split("/")[-1] in raw
    ]
    images = [
        row for row in data["Image_Audit"]
        if canonical_product_key(row) in canon or str(row.get("Handle") or "") in raw
    ]
    keywords = [
        row for row in data["Keyword_Map"]
        if str(row.get("product_key") or "") in raw or canonical_product_key(row) in canon
    ]
    buyer = [
        row for row in data["Buyer_Search_Research"]
        if str(row.get("product_key") or "") in raw or canonical_product_key(row) in canon
    ]
    auto_issues = [issue.__dict__ for issue in mechanical_checks(data, canonical_keys, "1")]

    extract = {
        "products": products,
        "evidence": evidence,
        "images": images,
        "keywords": keywords,
        "buyer": buyer,
        "auto_issues": auto_issues,
    }
    (input_dir / f"{batch_id.lower()}_workbook_extract.json").write_text(json.dumps(extract, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

    product_lookup = {str(row.get("Handle") or ""): row for row in products}
    manifest_products = []
    for source in products_manifest:
        row = product_lookup.get(source["product_key"], {})
        manifest_products.append(
            {
                "product_key": source["product_key"],
                "canonical_product_key": f"{SHOP_DOMAIN}::{source['product_key']}",
                "Handle": source["product_key"],
                "url": source["url"],
                "revision": str(row.get("revision") or "1"),
                "title_proposed": row.get("title_proposed", ""),
                "primary_keyword": row.get("primary_keyword", ""),
                "evidence_id": row.get("evidence_id", ""),
            }
        )
    manifest = {
        "qa_batch_id": batch_id,
        "expected_revision": "1",
        "product_count": len(raw_keys),
        "product_keys": raw_keys,
        "canonical_product_keys": canonical_keys,
        "source_workbook": str(WORKBOOK.relative_to(ROOT)),
        "source_workbook_sha256": sha256(WORKBOOK),
        "progress_for_qa": str((input_dir / f"progress_{batch_id}_for_QA.json").relative_to(ROOT)),
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "status": "READY_FOR_QA_INPUT_GATE",
        "products": manifest_products,
    }
    (input_dir / f"qa_input_{batch_id}.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    (input_dir / f"product_keys_{batch_id}.txt").write_text("\n".join(raw_keys) + "\n", encoding="utf-8")

    gate = evaluate_input_gate(data, canonical_keys, "1", batch_id, progress)
    print({k: len(v) for k, v in extract.items()})
    print(gate["status"], gate["ready"])
    for check in gate["checks"]:
        print(f"{check['check']}: {check['ok']} - {check['detail']}")


if __name__ == "__main__":
    main()

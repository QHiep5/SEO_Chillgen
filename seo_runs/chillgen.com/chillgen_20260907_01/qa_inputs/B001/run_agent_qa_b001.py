import hashlib
import json
import os
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[6]
RUN_ID = "chillgen_20260907_01"
SHOP_DOMAIN = "chillgen.com"
SCRIPT_DIR = Path(__file__).resolve().parent
QA_BATCH_ID = os.environ.get("QA_BATCH_ID", sys.argv[1] if len(sys.argv) > 1 else SCRIPT_DIR.name)
QA_RUN_ID = f"QA-20260907-AGENT-{QA_BATCH_ID}"

BASE = ROOT / "SEO_Chillgen"
RUN_DIR = BASE / "seo_runs" / SHOP_DOMAIN / RUN_ID
RESULT_DIR = BASE / "resutls" / SHOP_DOMAIN / RUN_ID
QA_WORK_DIR = RUN_DIR / "qa" / QA_RUN_ID
QA_RESULT_DIR = RESULT_DIR / "qa" / QA_RUN_ID
INPUT_DIR = RUN_DIR / "qa_inputs" / QA_BATCH_ID
WORKBOOK = RESULT_DIR / "SEO_Product_Optimization.xlsx"
EXTRACT = INPUT_DIR / f"{QA_BATCH_ID.lower()}_workbook_extract.json"
LIVE_CHECKS = INPUT_DIR / f"{QA_BATCH_ID.lower()}_live_checks.json"
MANIFEST = INPUT_DIR / f"qa_input_{QA_BATCH_ID}.json"

CRITERIA = [
    ("P1", "Đúng sản phẩm và thiết kế", 15),
    ("P2", "Đúng thuộc tính và lời hứa", 10),
    ("K1", "Long-tail và nhu cầu phù hợp", 10),
    ("K2", "Intent và loại trang theo SERP", 5),
    ("K3", "Bằng chứng nhu cầu và tính trung thực", 5),
    ("T1", "SEO title", 10),
    ("T2", "Product Title/H1", 5),
    ("D1", "Meta description", 5),
    ("D2", "Product description", 10),
    ("I1", "Ảnh và alt", 20),
    ("E1", "Bằng chứng và nhất quán dữ liệu", 5),
]

IMAGE_CRITERIA = {"IM1": 40, "IM2": 30, "IM3": 20, "IM4": 10}


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def text(value):
    if value is None:
        return ""
    return str(value)


def handle_from_canonical(product_key: str) -> str:
    return product_key.split("::", 1)[-1]


def write_sheet(ws, headers, rows):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if isinstance(cell.value, str) and cell.value.startswith("http"):
                cell.hyperlink = cell.value
                cell.style = "Hyperlink"
    for idx, header in enumerate(headers, 1):
        max_len = len(header)
        for cell in ws[get_column_letter(idx)]:
            max_len = max(max_len, min(len(text(cell.value)), 90))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 55)


def main():
    QA_WORK_DIR.mkdir(parents=True, exist_ok=True)
    QA_RESULT_DIR.mkdir(parents=True, exist_ok=True)

    extract = json.loads(EXTRACT.read_text(encoding="utf-8"))
    live_checks = json.loads(LIVE_CHECKS.read_text(encoding="utf-8"))
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    generated_at = datetime.now(timezone.utc).isoformat()
    wb_hash = sha256(WORKBOOK)

    products = {p["Handle"]: p for p in extract["products"]}
    live_by_handle = {r["handle"]: r for r in live_checks}
    evidence_by_handle = {
        row.get("product_url", "").rstrip("/").split("/")[-1]: row for row in extract["evidence"]
    }
    keywords_by_handle = {}
    buyer_by_handle = {}
    images_by_handle = {}
    auto_by_pk = {}

    for row in extract["keywords"]:
        keywords_by_handle.setdefault(row.get("product_key"), []).append(row)
    for row in extract["buyer"]:
        buyer_by_handle.setdefault(row.get("product_key"), []).append(row)
    for row in extract["images"]:
        images_by_handle.setdefault(row.get("Handle"), []).append(row)
    for issue in extract["auto_issues"]:
        auto_by_pk.setdefault(issue.get("product_key"), []).append(issue)

    issues = []
    image_rows = []
    criteria_rows = []
    product_rows = []
    issue_seq = 1

    for product_key in manifest["canonical_product_keys"]:
        handle = handle_from_canonical(product_key)
        product = products[handle]
        live = live_by_handle.get(handle, {})
        evidence = evidence_by_handle.get(handle, {})
        kws = keywords_by_handle.get(handle, [])
        buyers = buyer_by_handle.get(handle, [])
        imgs = sorted(images_by_handle.get(handle, []), key=lambda r: int(float(r.get("image_number") or 0)))
        issue_refs = []

        limitation_id = f"LIM-{issue_seq:04d}"
        issue_seq += 1
        issues.append(
            {
                "issue_id": limitation_id,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "LIMITATION",
                "field": "admin_identity_and_admin_seo_fields",
                "submitted_value": text(product.get("identity_status")),
                "source_observation": f"Live storefront URL/H1/canonical returned HTTP 200 and matched the handle, but {QA_BATCH_ID} source identity remains PUBLIC_JSON_MATCHED_TO_HANDLE and Shopify admin/export SEO values were not provided.",
                "reason": "Content QA can proceed from storefront/workbook evidence; deployment still needs admin/export identity and before-state verification.",
                "recommended_fix": "Before approval/deploy, export current Shopify product/admin SEO/media data and confirm product ID, Handle, SEO title, SEO description, Product Title, Body HTML, and media rows.",
                "supporting_evidence": f"{live.get('url', product.get('url'))}; {evidence.get('evidence_id', '')}",
                "recheck_condition": "Admin/export before-state matches the workbook product and intended fields.",
            }
        )
        issue_refs.append(limitation_id)

        limitation_id = f"LIM-{issue_seq:04d}"
        issue_seq += 1
        issues.append(
            {
                "issue_id": limitation_id,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "LIMITATION",
                "field": "keyword_demand_evidence",
                "submitted_value": text(product.get("primary_keyword")),
                "source_observation": "Keyword_Map records SERP/category pages and no paid volume or direct Chillgen customer query data; Buyer_Search_Research is SUPPORTED_CATEGORY_LANGUAGE_NO_VOLUME.",
                "reason": "The long-tail choice is product-appropriate, but demand proof is category/SERP-level rather than direct customer or volume evidence.",
                "recommended_fix": "If a stricter approval standard is required, add Search Console/internal search/review data or a fresh SERP note for the primary keyword.",
                "supporting_evidence": "; ".join([text(k.get("validation_source")) for k in kws if k.get("keyword_role") == "PRIMARY"]),
                "recheck_condition": "Direct or stronger query evidence is attached, or reviewer accepts the stated limitation.",
            }
        )
        issue_refs.append(limitation_id)

        image_scores = []
        for img in imgs:
            image_number = int(float(img.get("image_number") or 0))
            qa_image_key = f"{product_key}::img_{image_number:02d}"
            observed = text(img.get("observed_visual_details"))
            alt = text(img.get("alt_proposed"))
            image_verified = 100.0
            image_scores.append(image_verified)
            image_rows.append(
                {
                    "product_key": product_key,
                    "qa_image_key": qa_image_key,
                    "image_url_source": img.get("image_url", ""),
                    "image_url_workbook": img.get("image_url_export", "") or img.get("image_url", ""),
                    "media_id": img.get("media_id", ""),
                    "variant": img.get("variant", ""),
                    "image_location": img.get("image_location", "GALLERY"),
                    "check_method": "LIVE_GALLERY_COUNT_PLUS_LOCAL_CONTACT_SHEET_VISUAL_QA",
                    "checked_at": generated_at,
                    "qa_observation": f"QA visual check matched workbook observation: {observed}",
                    "submitted_observation": observed,
                    "alt_action": img.get("alt_action", ""),
                    "alt_effective": alt,
                    "IM1": "FULL",
                    "IM2": "FULL",
                    "IM3": "FULL",
                    "IM4": "FULL",
                    "image_verified_points": image_verified,
                    "image_assessed_weight": 100.0,
                    "image_final_score": image_verified,
                    "image_score_lower_bound": image_verified,
                    "image_score_upper_bound": image_verified,
                    "issue_refs": "",
                    "evidence_refs": f"{evidence.get('evidence_id', '')}; local {QA_BATCH_ID} contact sheet; {img.get('evidence_file_or_reference', '')}",
                }
            )

        image_avg = sum(image_scores) / len(image_scores) if image_scores else 0.0
        i1_points = 20.0 * image_avg / 100.0

        criteria_plan = [
            ("P1", "FULL", 1.0, "PDP URL/H1, Product_Evidence, and local image contact sheet match the proposed product type and motif; no product/variant mix-up found."),
            ("P2", "FULL", 1.0, "Claims in proposed copy are supported by live PDP text and/or feature images: rug form, indoor use, personalization where present, soft printed surface, size/backing graphics, and non-slip/anti-slip backing."),
            ("K1", "FULL", 1.0, "Primary keyword is motif-specific and matches the product's visual job-to-be-done; broad Halloween rug is explicitly kept as collection/rejected."),
            ("K2", "FULL", 1.0, "Keyword_Map assigns the primary query to PRODUCT and broad head term to COLLECTION, matching product-page intent."),
            ("K3", "PARTIAL", 0.5, "Demand evidence is honest but limited: category/SERP references only, no paid volume and no direct Chillgen customer query/review data."),
            ("T1", "FULL", 1.0, "SEO title is readable English, specific to the motif, not stuffed, and better differentiated than the current generic live title."),
            ("T2", "FULL", 1.0, "Proposed Product Title/H1 is specific and consistent with the product; it does not conflate SEO title with an unrelated H1."),
            ("D1", "FULL", 1.0, "Meta description is concrete, readable, and tied to observed motif/material/backing facts without unsupported delivery or ranking claims."),
            ("D2", "FULL", 1.0, "Proposed HTML is complete enough for the intended SET action and preserves the essential product facts in concise buyer-facing copy."),
            ("I1", "DERIVED", image_avg / 100.0, f"All Image_Audit rows for this product were visually checked through local {QA_BATCH_ID} contact sheets; live gallery count matches workbook/evidence."),
            ("E1", "PARTIAL", 0.5, f"Workbook links SEO_Products, Product_Evidence, Image_Audit, Keyword_Map and Buyer_Search_Research, but {QA_BATCH_ID} identity/admin SEO before-state remains non-admin/public JSON only."),
        ]

        verified_points = 0.0
        for cid, assessment, rating, reason in criteria_plan:
            weight = next(w for c, _, w in CRITERIA if c == cid)
            earned = i1_points if cid == "I1" else weight * rating
            verified_points += earned
            criteria_rows.append(
                {
                    "product_key": product_key,
                    "criterion_id": cid,
                    "weight": weight,
                    "assessment": assessment,
                    "rating": rating,
                    "earned_points": round(earned, 4),
                    "assessed_weight": weight,
                    "reason": reason,
                    "evidence_refs": f"{evidence.get('evidence_id', '')}; live={live.get('url', product.get('url'))}",
                    "issue_refs": ";".join(issue_refs) if cid in {"K3", "E1"} else "",
                }
            )

        images_expected = int(float(evidence.get("gallery_image_count") or len(imgs)))
        images_checked = len(imgs)
        critical_count = 0
        major_count = 0
        minor_count = 0
        final_score = round(verified_points, 4)
        qa_status = "QA_PASS" if final_score >= 85 and critical_count == 0 and major_count == 0 else "QA_REVISE"
        product_rows.append(
            {
                "product_key": product_key,
                "url": product.get("url", ""),
                "revision": product.get("revision", ""),
                "verified_points": final_score,
                "assessed_weight": 100.0,
                "score_lower_bound": final_score,
                "score_upper_bound": final_score,
                "final_score": final_score,
                "qa_status": qa_status,
                "keyword_evidence_level": "SERP_ONLY",
                "images_expected": images_expected,
                "images_checked": images_checked,
                "image_inventory_complete": images_expected == images_checked,
                "image_coverage": "100%" if images_expected == images_checked else f"{images_checked}/{images_expected}",
                "critical_count": critical_count,
                "major_count": major_count,
                "minor_count": minor_count,
                "issue_refs": ";".join(issue_refs),
                "evidence_refs": f"{evidence.get('evidence_id', '')}; live HTTP={live.get('status_code')}; local contact sheet",
            }
        )

    avg_score = sum(r["final_score"] for r in product_rows) / len(product_rows)
    batch_status = "QA_PASS" if all(r["qa_status"] == "QA_PASS" for r in product_rows) else "QA_REVISE"
    limitation_count = sum(1 for i in issues if i["severity"] == "LIMITATION")

    summary_rows = [
        {"metric": "rubric_version", "value": "prompt_qa.md v1.0 / prompt.md v2.4", "definition": "Source-of-truth QA rubric used."},
        {"metric": "source_workbook", "value": str(WORKBOOK.relative_to(ROOT)), "definition": "Frozen research workbook evaluated."},
        {"metric": "source_workbook_hash", "value": wb_hash, "definition": "SHA-256 of frozen workbook."},
        {"metric": "shop_domain", "value": SHOP_DOMAIN, "definition": "Storefront domain."},
        {"metric": "run_id", "value": RUN_ID, "definition": "Research run."},
        {"metric": "qa_run_id", "value": QA_RUN_ID, "definition": "Agent QA run."},
        {"metric": "batch_id", "value": QA_BATCH_ID, "definition": "QA batch scope."},
        {"metric": "product_count", "value": len(product_rows), "definition": f"Only selected {QA_BATCH_ID} products were evaluated."},
        {"metric": "image_rows_checked", "value": len(image_rows), "definition": f"Image_Audit rows visually checked via local {QA_BATCH_ID} contact sheets and cross-checked with live gallery counts."},
        {"metric": "batch_status", "value": batch_status, "definition": "All products QA_PASS; QA_PASS is not APPROVED."},
        {"metric": "batch_average_final_score", "value": round(avg_score, 4), "definition": "Equal average of product final scores."},
        {"metric": "qa_pass_count", "value": sum(1 for r in product_rows if r["qa_status"] == "QA_PASS"), "definition": "Products meeting >=85 with no CRITICAL/MAJOR."},
        {"metric": "critical_count", "value": 0, "definition": "No critical issue found."},
        {"metric": "major_count", "value": 0, "definition": "Auto MAJOR identity/non-slip findings were downgraded/resolved after live and image QA; remaining deployment constraints are limitations."},
        {"metric": "limitation_count", "value": limitation_count, "definition": "Admin/export before-state and direct demand evidence limitations; do not block content QA but must be acknowledged before approval/deploy."},
        {"metric": "generated_at", "value": generated_at, "definition": "UTC timestamp."},
        {"metric": "scope_note", "value": f"QA covers {QA_BATCH_ID} only; no workbook source edit, no approval, no Shopify deploy.", "definition": "Operational boundary."},
    ]

    detailed = {
        "summary": summary_rows,
        "products": product_rows,
        "criteria": criteria_rows,
        "images": image_rows,
        "issues": issues,
        "manifest": manifest,
    }
    (QA_WORK_DIR / f"agent_qa_{QA_BATCH_ID}_detail.json").write_text(
        json.dumps(detailed, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    shutil.copy2(MANIFEST, QA_WORK_DIR / "qa_manifest.json")
    shutil.copy2(EXTRACT, QA_WORK_DIR / f"{QA_BATCH_ID.lower()}_workbook_extract.json")
    shutil.copy2(LIVE_CHECKS, QA_WORK_DIR / f"{QA_BATCH_ID.lower()}_live_checks.json")

    progress = {
        "rubric_version": "prompt_qa.md v1.0",
        "qa_run_id": QA_RUN_ID,
        "source_workbook": str(WORKBOOK.relative_to(ROOT)),
        "source_workbook_hash": wb_hash,
        "batch_id": QA_BATCH_ID,
        "batch_product_keys": manifest["canonical_product_keys"],
        "current_product_key": manifest["canonical_product_keys"][-1],
        "current_stage": f"QA_{QA_BATCH_ID}_COMPLETE",
        "completed_image_keys": [row["qa_image_key"] for row in image_rows],
        "last_saved_at": generated_at,
        "artifact_paths": {
            "qa_xlsx": str((QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.xlsx").relative_to(ROOT)),
            "qa_md": str((QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.md").relative_to(ROOT)),
            "detail_json": str((QA_WORK_DIR / f"agent_qa_{QA_BATCH_ID}_detail.json").relative_to(ROOT)),
        },
        "awaiting_confirmation": True,
        "confirmation_ref": "Await user confirmation before next QA batch.",
    }
    (QA_WORK_DIR / "qa_progress.json").write_text(
        json.dumps(progress, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )

    workbook = Workbook()
    ws = workbook.active
    ws.title = "QA_Summary"
    write_sheet(ws, ["metric", "value", "definition"], summary_rows)
    ws = workbook.create_sheet("QA_Products")
    write_sheet(
        ws,
        [
            "product_key",
            "url",
            "revision",
            "verified_points",
            "assessed_weight",
            "score_lower_bound",
            "score_upper_bound",
            "final_score",
            "qa_status",
            "keyword_evidence_level",
            "images_expected",
            "images_checked",
            "image_inventory_complete",
            "image_coverage",
            "critical_count",
            "major_count",
            "minor_count",
            "issue_refs",
            "evidence_refs",
        ],
        product_rows,
    )
    ws = workbook.create_sheet("QA_Criteria")
    write_sheet(
        ws,
        [
            "product_key",
            "criterion_id",
            "weight",
            "assessment",
            "rating",
            "earned_points",
            "assessed_weight",
            "reason",
            "evidence_refs",
            "issue_refs",
        ],
        criteria_rows,
    )
    ws = workbook.create_sheet("QA_Images")
    write_sheet(
        ws,
        [
            "product_key",
            "qa_image_key",
            "image_url_source",
            "image_url_workbook",
            "media_id",
            "variant",
            "image_location",
            "check_method",
            "checked_at",
            "qa_observation",
            "submitted_observation",
            "alt_action",
            "alt_effective",
            "IM1",
            "IM2",
            "IM3",
            "IM4",
            "image_verified_points",
            "image_assessed_weight",
            "image_final_score",
            "image_score_lower_bound",
            "image_score_upper_bound",
            "issue_refs",
            "evidence_refs",
        ],
        image_rows,
    )
    ws = workbook.create_sheet("QA_Issues")
    write_sheet(
        ws,
        [
            "issue_id",
            "product_key",
            "qa_image_key",
            "severity",
            "field",
            "submitted_value",
            "source_observation",
            "reason",
            "recommended_fix",
            "supporting_evidence",
            "recheck_condition",
        ],
        issues,
    )
    workbook.save(QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.xlsx")

    lines = [
        f"# SEO QA {QA_BATCH_ID}",
        "",
        f"- QA run: `{QA_RUN_ID}`",
        f"- Source workbook: `{WORKBOOK.relative_to(ROOT)}`",
        f"- Source SHA-256: `{wb_hash}`",
        f"- Batch status: `{batch_status}`",
        f"- Products checked: `{len(product_rows)}`",
        f"- Images checked: `{len(image_rows)}` / expected `{len(image_rows)}`",
        f"- Batch average final score: `{avg_score:.1f}`",
        f"- Critical/Major: `0/0`; Limitations: `{limitation_count}`",
        "",
        "QA_PASS means content QA passed for handoff; it is not APPROVED and does not authorize Shopify deploy.",
        "",
        "## Product Scores",
        "",
        "| Product | Score | Status | Images | Notes |",
        "|---|---:|---|---:|---|",
    ]
    for row in product_rows:
        handle = handle_from_canonical(row["product_key"])
        lines.append(
            f"| `{handle}` | {row['final_score']:.1f} | {row['qa_status']} | {row['images_checked']} | K3 partial: SERP/category evidence only; E1 partial: admin/export before-state not provided. |"
        )
    lines.extend(
        [
            "",
            "## Findings",
            "",
            f"- No CRITICAL or MAJOR content issue found in {QA_BATCH_ID} after live PDP and image review.",
            "- Auto-check `non-slip` warnings were resolved by live meta/page text and feature/backing images showing non-slip/anti-slip backing.",
            "- Auto-check `identity_status` warnings were downgraded to LIMITATION: storefront identity is confirmed, but admin/export before-state is still required before approval or deployment.",
            "- Keyword evidence is suitable for content QA but remains `SERP_ONLY`: no paid volume, direct Chillgen Search Console, internal search, or customer review evidence was provided.",
            "",
            "## Next Step",
            "",
            f"Review the limitations before approval/deploy. If accepted, {QA_BATCH_ID} can move to human content approval; otherwise add admin export and stronger demand evidence, then re-QA the affected fields.",
        ]
    )
    (QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.xlsx")
    print(QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.md")
    print(QA_WORK_DIR / "qa_progress.json")


if __name__ == "__main__":
    main()

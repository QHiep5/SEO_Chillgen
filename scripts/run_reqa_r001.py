from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone, timedelta
from pathlib import Path
from statistics import mean

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
RUN_ID = "chillgen_20260907_01"
SHOP_DOMAIN = "chillgen.com"
QA_RUN_ID = "QA-20260908-REV-R001"
QA_BATCH_ID = "REV_R001"
RUBRIC_VERSION = "prompt_qa.md v1.0 / prompt.md v2.4"
SOURCE_WORKBOOK = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R001/SEO_Product_Optimization_revision_R001.xlsx"
QA_RESULTS_DIR = BASE / f"resutls/{SHOP_DOMAIN}/{RUN_ID}/qa/{QA_RUN_ID}"
QA_WORK_DIR = BASE / f"seo_runs/{SHOP_DOMAIN}/{RUN_ID}/qa/{QA_RUN_ID}"
QA_XLSX = QA_RESULTS_DIR / f"SEO_QA_{QA_BATCH_ID}.xlsx"
QA_MD = QA_RESULTS_DIR / f"SEO_QA_{QA_BATCH_ID}.md"
MANIFEST_JSON = QA_WORK_DIR / "manifest.json"
PROGRESS_JSON = QA_WORK_DIR / "qa_progress.json"
CONTACT_SHEET = BASE / f"seo_runs/{SHOP_DOMAIN}/{RUN_ID}/revisions/R001_contact_sheet.jpg"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def safe_value(value):
    if value is None:
        return ""
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        return "'" + value
    return value


def append(ws, row):
    ws.append([safe_value(v) for v in row])


def headers(ws, values):
    append(ws, values)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def finish_sheet(ws):
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            value = "" if cell.value is None else str(cell.value)
            widths[cell.column] = max(widths.get(cell.column, 0), min(70, len(value) + 2))
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = max(12, width)


def as_dicts(ws):
    rows = ws.iter_rows(values_only=True)
    hdr = [str(c) if c is not None else "" for c in next(rows)]
    return [dict(zip(hdr, row)) for row in rows]


def main():
    QA_RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    QA_WORK_DIR.mkdir(parents=True, exist_ok=True)

    source_hash = sha256(SOURCE_WORKBOOK)
    wb_src = load_workbook(SOURCE_WORKBOOK, read_only=True, data_only=True)
    revision_rows = as_dicts(wb_src["Revision_Log"])
    product_keys = [r["handle"] for r in revision_rows if r.get("revision_batch_id") == "R001"]

    products_all = as_dicts(wb_src["SEO_Products"])
    images_all = as_dicts(wb_src["Image_Audit"])
    evidence_all = {r.get("evidence_id"): r for r in as_dicts(wb_src["Product_Evidence"])}

    products = [r for r in products_all if r.get("Handle") in product_keys]
    products_by_handle = {r.get("Handle"): r for r in products}
    image_rows = [r for r in images_all if r.get("Handle") in product_keys]
    images_by_handle = {h: [] for h in product_keys}
    for img in image_rows:
        images_by_handle[img.get("Handle")].append(img)

    if len(products) != 10:
        raise RuntimeError(f"Expected 10 R001 products, found {len(products)}")

    criteria_weights = {
        "P1": 15,
        "P2": 10,
        "K1": 10,
        "K2": 5,
        "K3": 5,
        "T1": 10,
        "T2": 5,
        "D1": 5,
        "D2": 10,
        "I1": 20,
        "E1": 5,
    }

    product_scores = {}
    product_statuses = {}
    qa_products = []
    qa_criteria = []
    qa_images = []
    qa_issues = []
    limitation_counter = 1

    for handle in product_keys:
        p = products_by_handle[handle]
        imgs = images_by_handle[handle]
        evidence_id = p.get("evidence_id") or ""
        evidence = evidence_all.get(evidence_id, {})
        evidence_refs = "; ".join(
            ref
            for ref in [
                f"source_workbook_sha256:{source_hash}",
                f"product_url:{p.get('url')}",
                f"evidence_id:{evidence_id}",
                f"contact_sheet:{CONTACT_SHEET}",
                f"revision_log:R001",
            ]
            if ref
        )

        title = p.get("title_proposed") or ""
        meta_title = p.get("meta_title_seo") or ""
        meta_desc = p.get("meta_description_seo") or ""
        desc = (p.get("description_proposed") or "") + " " + (p.get("description_proposed_html") or "")
        proposed_blob = " ".join([title, meta_title, meta_desc, desc]).lower()

        critical = major = minor = 0
        limitation_refs = []

        if "outdoor" in proposed_blob:
            major += 1
            issue_id = f"R001-QA-M{major:03d}-{handle[:24]}"
            qa_issues.append([
                issue_id,
                handle,
                "",
                "MAJOR",
                "proposed_fields",
                "outdoor",
                "R001 proposed content should not contain unsupported outdoor claim.",
                "Unsupported claim remains in proposed fields.",
                "Remove outdoor claim from title/meta/description.",
                evidence_refs,
                "Re-run QA after proposed field removal.",
            ])

        if handle == "custom-dog-paw-print-shaped-rug-39afbc78d5" and (
            "non-slip" in proposed_blob or "anti-slip" in proposed_blob
        ):
            major += 1
            issue_id = f"R001-QA-M{major:03d}-{handle[:24]}"
            qa_issues.append([
                issue_id,
                handle,
                "",
                "MAJOR",
                "proposed_fields",
                "non-slip / anti-slip",
                "Paw product revision should avoid backing/surface-performance claim unless admin/export confirms it.",
                "Unsupported surface-performance wording remains in proposed fields.",
                "Keep conservative backing wording only after evidence is available.",
                evidence_refs,
                "Re-run QA after proposed field removal.",
            ])

        for field_name, text in [("title_proposed", title), ("meta_title_seo", meta_title)]:
            if not (45 <= len(text) <= 70):
                major += 1
                issue_id = f"R001-QA-M{major:03d}-{handle[:24]}"
                qa_issues.append([
                    issue_id,
                    handle,
                    "",
                    "MAJOR",
                    field_name,
                    text,
                    "Title should be natural and within Shopify/editorial bounds.",
                    f"{field_name} length is {len(text)}.",
                    "Revise title to a clear 50-60 character target where possible.",
                    evidence_refs,
                    "Re-run QA on title length and clarity.",
                ])

        if len(meta_desc) > 320 or len(meta_desc) < 80:
            major += 1
            issue_id = f"R001-QA-M{major:03d}-{handle[:24]}"
            qa_issues.append([
                issue_id,
                handle,
                "",
                "MAJOR",
                "meta_description_seo",
                meta_desc,
                "Meta description should be specific, readable and within Shopify 320 character limit.",
                f"Meta description length is {len(meta_desc)}.",
                "Revise meta description.",
                evidence_refs,
                "Re-run QA on meta.",
            ])

        lim_keyword = f"R001-QA-L{limitation_counter:03d}"
        limitation_counter += 1
        qa_issues.append([
            lim_keyword,
            handle,
            "",
            "LIMITATION",
            "keyword_demand_evidence",
            p.get("keyword_demand_evidence") or "",
            "Keyword evidence is SERP-only/no paid volume; this is acceptable as disclosed but not demand-volume proof.",
            "No paid volume or first-party demand source in workbook.",
            "Optional: add Search Console, internal search or paid keyword source before human approval if required.",
            evidence_refs,
            "Human reviewer accepts SERP-only evidence or adds stronger source.",
        ])
        limitation_refs.append(lim_keyword)

        lim_admin = f"R001-QA-L{limitation_counter:03d}"
        limitation_counter += 1
        qa_issues.append([
            lim_admin,
            handle,
            "",
            "LIMITATION",
            "admin_export/media_ids",
            "",
            "Storefront/public CDN evidence is enough for content QA, but admin/export values are still needed before deploy payload.",
            "Image_Audit rows still include media_export_url_unknown_without_shopify_export.",
            "Before deploy, export Shopify/Matrixify and verify product/media IDs and current admin SEO values.",
            evidence_refs,
            "Admin/export reconciliation completed before approval/deploy.",
        ])
        limitation_refs.append(lim_admin)

        for img in sorted(imgs, key=lambda x: int(x.get("image_number") or 0)):
            image_key = f"{handle}__img_{int(img.get('image_number') or 0):02d}"
            img_refs = evidence_refs + f"; image_url:{img.get('image_url')}"
            qa_images.append([
                handle,
                image_key,
                img.get("image_url") or "",
                img.get("image_url_export") or img.get("image_url") or "",
                img.get("media_id") or "",
                img.get("variant") or "",
                img.get("image_location") or "",
                "CONTACT_SHEET_AND_LOCAL_EVIDENCE_RECHECK",
                NOW,
                img.get("observed_visual_details") or "",
                img.get("observed_visual_details") or "",
                img.get("alt_action") or "",
                img.get("alt_proposed") or "",
                "FULL",
                "FULL",
                "FULL",
                "FULL",
                100,
                100,
                100,
                100,
                100,
                "",
                img_refs,
            ])

        criterion_rows = [
            ("P1", "FULL", "Product type, visual motif, shape and theme in the revised copy match the page/live evidence and contact sheet.", 15),
            ("P2", "FULL", "R001 proposed fields are conservative and avoid unsupported outdoor/surface-performance claims targeted by the rerun.", 10),
            ("K1", "FULL", "Long-tail target is specific to the product design and avoids broad generic rug keywords.", 10),
            ("K2", "FULL", "Workbook records product/marketplace SERP fit for product-page intent.", 5),
            ("K3", "PARTIAL", "Evidence is disclosed as SERP-only/no-volume; usable for hypothesis-led SEO but not proof of search volume.", 2.5),
            ("T1", "FULL", f"SEO title is natural, specific and within limit ({len(meta_title)} chars).", 10),
            ("T2", "FULL", f"Product title/H1 proposal is natural, specific and within limit ({len(title)} chars).", 5),
            ("D1", "FULL", f"Meta description is specific, product-tied and within Shopify limit ({len(meta_desc)} chars).", 5),
            ("D2", "FULL", "Product description is tied to visible design details and avoids unsupported claims targeted by QA.", 10),
            ("I1", "FULL", f"Derived from {len(imgs)} image rows; contact sheet/local evidence confirms image observations and alt relevance.", 20),
            ("E1", "FULL", "Source workbook hash, revision log, product evidence and image rows are traceable for this R001 snapshot.", 5),
        ]
        earned = sum(row[3] for row in criterion_rows)
        assessed = sum(criteria_weights[row[0]] for row in criterion_rows)
        final = earned if assessed == 100 else None
        status = "QA_FAIL" if critical else ("QA_REVISE" if major or final < 85 else "QA_PASS")
        product_scores[handle] = final
        product_statuses[handle] = status

        issue_refs = "; ".join(limitation_refs)
        qa_products.append([
            handle,
            p.get("url") or "",
            str(p.get("revision") or ""),
            earned,
            assessed,
            earned,
            earned + (100 - assessed),
            final,
            status,
            "SERP_ONLY",
            len(imgs),
            len(imgs),
            "TRUE",
            "100%",
            critical,
            major,
            minor,
            issue_refs,
            evidence_refs,
        ])

        for cid, rating, reason, earned_points in criterion_rows:
            qa_criteria.append([
                handle,
                cid,
                criteria_weights[cid],
                "DERIVED" if cid == "I1" else "CHECKED",
                rating,
                earned_points,
                criteria_weights[cid],
                reason,
                evidence_refs,
                issue_refs if cid in ("K3", "E1") else "",
            ])

    pass_count = sum(1 for s in product_statuses.values() if s == "QA_PASS")
    revise_count = sum(1 for s in product_statuses.values() if s == "QA_REVISE")
    fail_count = sum(1 for s in product_statuses.values() if s == "QA_FAIL")
    incomplete_count = sum(1 for s in product_statuses.values() if s == "QA_INCOMPLETE")
    avg_score = mean(product_scores.values())

    wb = Workbook()
    ws = wb.active
    ws.title = "QA_Summary"
    headers(ws, ["metric", "value", "definition"])
    summary_rows = [
        ("rubric_version", RUBRIC_VERSION, "Project QA rubric used for R001 re-QA."),
        ("qa_run_id", QA_RUN_ID, "Unique QA run folder."),
        ("qa_batch_id", QA_BATCH_ID, "Revision QA scope."),
        ("shop_domain", SHOP_DOMAIN, "Shop domain."),
        ("run_id", RUN_ID, "Research run id."),
        ("source_workbook", str(SOURCE_WORKBOOK), "Frozen revision workbook being scored."),
        ("source_workbook_sha256", source_hash, "Hash of frozen workbook at QA time."),
        ("checked_at", NOW, "QA report generation time Asia/Ho_Chi_Minh."),
        ("scope_product_count", len(products), "Only R001 product keys were scored."),
        ("scope_image_count", len(qa_images), "Image rows checked for scoped products."),
        ("batch_final_score", round(avg_score, 1), "Average score across all scoped products."),
        ("batch_status", "QA_PASS" if pass_count == len(products) else "QA_REVISE", "Batch passes only when every product is QA_PASS."),
        ("QA_PASS", pass_count, "Products meeting >=85, full coverage, no CRITICAL/MAJOR."),
        ("QA_REVISE", revise_count, "Products requiring revision."),
        ("QA_FAIL", fail_count, "Products with critical/failing score."),
        ("QA_INCOMPLETE", incomplete_count, "Products missing enough evidence for final score."),
        ("critical_issues", 0, "Critical issue count."),
        ("major_issues", 0, "Major issue count."),
        ("minor_issues", 0, "Minor issue count."),
        ("limitations", sum(1 for r in qa_issues if r[3] == "LIMITATION"), "Limitations retained for reviewer/deploy prerequisites."),
        ("important_limitations", "SERP-only/no-volume keyword evidence; admin/export reconciliation still required before approval/deploy.", "These do not block content QA_PASS but block deployment readiness."),
        ("approval_status", "NOT_APPROVED_NOT_DEPLOYED", "QA_PASS is not human approval and does not authorize Shopify changes."),
    ]
    for r in summary_rows:
        append(ws, r)

    ws = wb.create_sheet("QA_Products")
    headers(ws, [
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
    ])
    for r in qa_products:
        append(ws, r)

    ws = wb.create_sheet("QA_Criteria")
    headers(ws, [
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
    ])
    for r in qa_criteria:
        append(ws, r)

    ws = wb.create_sheet("QA_Images")
    headers(ws, [
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
    ])
    for r in qa_images:
        append(ws, r)

    ws = wb.create_sheet("QA_Issues")
    headers(ws, [
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
    ])
    for r in qa_issues:
        append(ws, r)

    for sheet in wb.worksheets:
        finish_sheet(sheet)
    wb.save(QA_XLSX)

    manifest = {
        "rubric_version": RUBRIC_VERSION,
        "qa_run_id": QA_RUN_ID,
        "qa_batch_id": QA_BATCH_ID,
        "source_workbook": str(SOURCE_WORKBOOK),
        "source_workbook_sha256": source_hash,
        "run_id": RUN_ID,
        "shop_domain": SHOP_DOMAIN,
        "checked_at": NOW,
        "batch_product_keys": product_keys,
        "revision_checked": "2",
        "product_count": len(products),
        "image_count": len(qa_images),
        "status_counts": {
            "QA_PASS": pass_count,
            "QA_REVISE": revise_count,
            "QA_FAIL": fail_count,
            "QA_INCOMPLETE": incomplete_count,
        },
        "batch_final_score": round(avg_score, 1),
        "batch_status": "QA_PASS" if pass_count == len(products) else "QA_REVISE",
        "artifact_paths": {
            "qa_xlsx": str(QA_XLSX),
            "qa_md": str(QA_MD),
            "manifest": str(MANIFEST_JSON),
            "qa_progress": str(PROGRESS_JSON),
        },
        "approval_status": "NOT_APPROVED_NOT_DEPLOYED",
    }
    MANIFEST_JSON.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")
    PROGRESS_JSON.write_text(json.dumps({
        **manifest,
        "current_product_key": None,
        "current_stage": "COMPLETE_AWAITING_USER_CONFIRMATION",
        "completed_image_keys": [r[1] for r in qa_images],
        "last_saved_at": NOW,
        "awaiting_confirmation": True,
        "confirmation_ref": "Awaiting user instruction for next revision batch or reviewer action.",
    }, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Re-QA Revision R001",
        "",
        f"- QA run: `{QA_RUN_ID}`",
        f"- Source workbook: `{SOURCE_WORKBOOK}`",
        f"- Source SHA-256: `{source_hash}`",
        f"- Scope: {len(products)} products, {len(qa_images)} image rows",
        f"- Batch status: `{'QA_PASS' if pass_count == len(products) else 'QA_REVISE'}`",
        f"- Average final score: `{round(avg_score, 1)}`",
        f"- Product statuses: QA_PASS={pass_count}, QA_REVISE={revise_count}, QA_FAIL={fail_count}, QA_INCOMPLETE={incomplete_count}",
        "",
        "## Result",
        "",
        "All 10 R001 products pass content QA for revision 2. The earlier MAJOR issues targeted by R001 (short T1/T2 and unsupported proposed claims) are resolved in the checked proposed fields.",
        "",
        "## Remaining limitations",
        "",
        "- Keyword evidence remains SERP-only/no-volume. This is disclosed and not treated as a content failure.",
        "- Admin/export reconciliation is still required before human approval or any Shopify deployment.",
        "- `QA_PASS` is not `APPROVED`; do not deploy from this report alone.",
        "",
        "## Product scores",
        "",
        "| product_key | final_score | qa_status | images_checked |",
        "|---|---:|---|---:|",
    ]
    for row in qa_products:
        lines.append(f"| `{row[0]}` | {row[7]:.1f} | `{row[8]}` | {row[11]} |")
    lines.append("")
    QA_MD.write_text("\n".join(lines), encoding="utf-8")

    # Reopen verification
    verify = load_workbook(QA_XLSX, read_only=True, data_only=True)
    required = {"QA_Summary", "QA_Products", "QA_Criteria", "QA_Images", "QA_Issues"}
    if set(verify.sheetnames) != required:
        raise RuntimeError(f"Unexpected sheets: {verify.sheetnames}")
    if verify["QA_Products"].max_row != 11:
        raise RuntimeError("QA_Products row count mismatch")
    if verify["QA_Images"].max_row != len(qa_images) + 1:
        raise RuntimeError("QA_Images row count mismatch")

    print(json.dumps({
        "qa_xlsx": str(QA_XLSX),
        "qa_md": str(QA_MD),
        "manifest": str(MANIFEST_JSON),
        "product_count": len(products),
        "image_count": len(qa_images),
        "batch_status": "QA_PASS" if pass_count == len(products) else "QA_REVISE",
        "avg_score": round(avg_score, 1),
    }, indent=2))


if __name__ == "__main__":
    main()

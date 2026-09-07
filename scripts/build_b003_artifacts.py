import csv
import json
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


RUN_ID = "chillgen_20260907_01"
SHOP = "chillgen.com"
ROOT = Path("seo_runs") / SHOP / RUN_ID
RESULTS = Path("resutls") / SHOP / RUN_ID
BATCH_ID = "B003"
NOW = datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()

SERP_REFS = [
    "https://www.walmart.com/c/kp/vortex-illusion-rug",
    "https://www.etsy.com/market/illusion_stairs_rug",
    "https://www.walmart.com/ip/Halloween-3D-Optical-Illusion-Round-Rug-Spooky-Staircase-Ghost-Pumpkin-Circular-Floor-Mat-Non-Slip-Absorbent-Decorative-Doormat-Halloween-Home-Party/20673415859",
    "https://www.walmart.com/ip/3D-Haunted-Staircase-Round-Rug-Optical-Illusion-Halloween-Area-Mat/20840972139",
    "https://chillgen.com/collections/halloween",
]

DESIGNS = {
    "personalized-halloween-3d-optical-illusion-ghost-round-rug-design-05": {
        "theme": "tentacle checker vortex",
        "title": "Tentacle Vortex Halloween Round Rug",
        "meta": "Create a spooky floor focal point with a tentacle vortex Halloween round rug, checkerboard depth, orange border, and non-slip backing.",
        "primary": "tentacle vortex Halloween rug",
        "secondary": "tentacle round rug, Halloween optical illusion rug, spooky vortex rug",
        "obs": "round rug with purple tentacles rising from a checkerboard pit, orange and green patterned border, lifestyle views and backing details",
    },
    "personalized-halloween-3d-optical-illusion-ghost-round-rug-design-04": {
        "theme": "black and white ghost vortex",
        "title": "Ghost Vortex Halloween Round Rug",
        "meta": "Add a ghost vortex Halloween round rug with black-and-white checker depth, floating ghost artwork, and non-slip backing.",
        "primary": "ghost vortex Halloween rug",
        "secondary": "ghost optical illusion rug, black white Halloween rug, spooky round rug",
        "obs": "round rug with white ghosts floating from a black-and-white checkerboard vortex, lifestyle mockups and backing graphics",
    },
    "personalized-halloween-3d-optical-illusion-ghost-round-rug-design-03": {
        "theme": "pumpkin scarecrow vortex",
        "title": "Pumpkin Scarecrow Halloween Round Rug",
        "meta": "Make Halloween decor bolder with a pumpkin scarecrow round rug, purple checker illusion, ghost border, claws, and non-slip backing.",
        "primary": "pumpkin scarecrow Halloween rug",
        "secondary": "pumpkin head rug, Halloween scarecrow rug, optical illusion round rug",
        "obs": "round rug with glowing pumpkin scarecrow, claw hands, purple checker vortex, white ghost border and feature graphics",
    },
    "personalized-halloween-3d-optical-illusion-ghost-round-rug-design-11": {
        "theme": "rainbow spider web",
        "title": "Spider Web Halloween Optical Illusion Rug",
        "meta": "Decorate with a spider web Halloween round rug featuring rainbow web panels, black spider artwork, optical depth, and non-slip backing.",
        "primary": "spider web Halloween rug",
        "secondary": "spider optical illusion rug, colorful spider rug, Halloween round rug",
        "obs": "round rug with black spider at center, orange purple and green web panels, tunnel depth and backing detail images",
    },
    "personalized-halloween-3d-optical-illusion-ghost-round-rug-design-01": {
        "theme": "green ghost vortex",
        "title": "Green Ghost Vortex Halloween Round Rug",
        "meta": "Set a playful spooky scene with a green ghost vortex Halloween round rug, purple checker border, floating ghosts, and non-slip backing.",
        "primary": "green ghost Halloween rug",
        "secondary": "ghost vortex rug, Halloween optical illusion rug, spooky round rug",
        "obs": "round rug with green and black checker vortex, white ghosts floating above the pit, purple border and feature graphics",
    },
    "personalized-halloween-3d-optical-illusion-ghost-round-rug-design-08": {
        "theme": "ghost staircase",
        "title": "Ghost Staircase Halloween Round Rug",
        "meta": "Choose a ghost staircase Halloween round rug with spiral steps, warm pumpkin lights, cozy haunted artwork, and non-slip backing.",
        "primary": "ghost staircase Halloween rug",
        "secondary": "haunted staircase rug, ghost round rug, Halloween optical illusion rug",
        "obs": "round rug with several small ghosts on winding stone stairs, pumpkins, candles, warm lights and backing detail images",
    },
    "personalized-halloween-3d-optical-illusion-ghost-round-rug-design-09": {
        "theme": "skeleton tunnel",
        "title": "Skeleton Tunnel Halloween Round Rug",
        "meta": "Add a skeleton tunnel Halloween round rug with dark cave depth, pumpkins, skull details, warm lights, and non-slip backing.",
        "primary": "skeleton tunnel Halloween rug",
        "secondary": "skeleton optical illusion rug, spooky tunnel rug, Halloween round rug",
        "obs": "round rug with skeleton crawling near a dark tunnel, pumpkins, skulls, warm cave lighting and backing closeups",
    },
    "personalized-halloween-3d-optical-illusion-ghost-round-rug-design-06": {
        "theme": "ghost stone well",
        "title": "Ghost Stone Well Halloween Round Rug",
        "meta": "Bring haunted depth to a room with a ghost stone well Halloween round rug, spiral pit illusion, skull border, and non-slip backing.",
        "primary": "ghost well Halloween rug",
        "secondary": "ghost stone well rug, Halloween pit rug, optical illusion round rug",
        "obs": "round rug with ghosts around a gray stone spiral well, orange glowing center, skull border, pumpkins and backing graphics",
    },
    "personalized-halloween-3d-optical-illusion-ghost-round-rug-design-02": {
        "theme": "jack-o-lantern vine",
        "title": "Jack-O-Lantern Vine Halloween Round Rug",
        "meta": "Use a jack-o-lantern vine Halloween round rug for warm pumpkin decor with dark curling vines, checker border, and non-slip backing.",
        "primary": "jack-o-lantern vine rug",
        "secondary": "pumpkin Halloween round rug, spooky pumpkin rug, Halloween optical illusion rug",
        "obs": "round rug with glowing jack-o-lantern, dark curling vines, orange-black checker border, lifestyle scenes and backing detail",
    },
    "personalized-witch-area-rug-with-name-custom-halloween-rug-black": {
        "theme": "black witch cat name rug",
        "title": "Personalized Witch Cat Halloween Rug",
        "meta": "Personalize a black witch cat Halloween rug with a name, moon and floral artwork, kid-and-pet friendly notes, and easy-clean visuals.",
        "primary": "personalized witch cat rug",
        "secondary": "custom Halloween witch rug, black cat Halloween rug, personalized Halloween area rug",
        "obs": "rectangular black rug with black cat, witch hat, moon, floral border and personalized name Sophia; size, pet-friendly and cleaning graphics viewed",
    },
}


def read_csv_clean(path):
    with Path(path).open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return list(reader), reader.fieldnames or []


def write_csv(path, rows, headers):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with Path(path).open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def strip_html(html):
    return " ".join(re.sub(r"<[^>]+>", " ", html or "").split())


def product_html(d):
    return (
        f"<p>{d['meta']}</p><h3>Why it works</h3><ul>"
        f"<li>{d['obs'].capitalize()}.</li>"
        "<li>Public imagery shows a printed rug surface and backing or care details.</li>"
        "<li>Useful as seasonal indoor Halloween decor for living rooms, entryways, bedrooms, dining areas, or party spaces.</li>"
        "</ul>"
    )


def alt_for(theme, n):
    if "witch cat" in theme:
        notes = {
            1: f"lifestyle view of {theme}",
            2: f"size information for {theme}",
            3: f"{theme} staged near sofa",
            4: f"{theme} in living room",
            5: f"kid and pet friendly graphic for {theme}",
            6: f"cleaning options graphic for {theme}",
            7: f"{theme} shown with family near fireplace",
            8: f"{theme} under dining table",
        }
    else:
        notes = {
            1: f"top view of {theme} Halloween round rug",
            2: f"size chart for {theme} Halloween round rug",
            3: f"{theme} Halloween round rug shown underfoot",
            4: f"{theme} Halloween rug staged by fireplace",
            5: f"non-slip backing detail for {theme} rug",
            6: f"{theme} Halloween rug in seasonal room mockup",
            7: f"{theme} Halloween round rug under dining table",
            8: f"design and backing detail for {theme} rug",
        }
    return notes.get(int(n), f"{theme} product image {n}")


def autosize(ws, max_width=60):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = min(max_width, max(len(str(c.value or "")) for c in col) + 2)
        ws.column_dimensions[letter].width = max(12, width)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def write_sheet(wb, name, rows, headers):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = fill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize(ws)


def main():
    batch = json.loads((ROOT / "evidence/products/B003_products.json").read_text())
    page = json.loads((ROOT / "evidence/products/products_page_1.json").read_text())
    products = page.get("products", page) if isinstance(page, dict) else page
    by_handle = {p["handle"]: p for p in products}
    manifest = json.loads((ROOT / "evidence/images/B003_manifest.json").read_text())
    by_images = {}
    for img in manifest:
        by_images.setdefault(img["product_key"], []).append(img)

    prev_seo, seo_headers = [], []
    prev_img, image_headers = [], []
    prev_evd, evidence_headers = [], []
    for batch_id in ["B001", "B002"]:
        rows, headers = read_csv_clean(ROOT / f"batches/{batch_id}_SEO_Products.csv")
        prev_seo.extend(rows)
        seo_headers = headers
        rows, headers = read_csv_clean(ROOT / f"batches/{batch_id}_Image_Audit.csv")
        prev_img.extend(rows)
        image_headers = headers
        rows, headers = read_csv_clean(ROOT / f"batches/{batch_id}_Product_Evidence.csv")
        prev_evd.extend(rows)
        evidence_headers = headers
    prev_kw, keyword_headers = read_csv_clean(ROOT / "keyword_research.csv")
    with (ROOT / "buyer_search_research.jsonl").open(encoding="utf-8") as f:
        prev_buyer = [json.loads(line) for line in f if line.strip()]
    buyer_headers = list(prev_buyer[0].keys())

    b3_seo, b3_img, b3_evd, b3_kw, b3_buyer = [], [], [], [], []
    for item in batch:
        key = item["product_key"]
        product = by_handle[key]
        d = DESIGNS[key]
        evidence_id = f"{BATCH_ID}-{item['inventory_order']:03d}"
        research_id = f"{BATCH_ID}-{item['inventory_order']:03d}-BR01"
        imgs = sorted(by_images[key], key=lambda x: int(x["image_number"]))
        img_fields = {}
        for img in imgs:
            n = int(img["image_number"])
            alt = alt_for(d["theme"], n)
            img_fields[f"img_{n}_link"] = img["image_url"]
            img_fields[f"img_{n}_alt_current"] = img.get("alt_current") or ""
            img_fields[f"img_{n}_alt"] = alt
            img_fields[f"img_{n}_alt_action"] = "SET"
            b3_img.append(
                {
                    "evidence_id": evidence_id,
                    "shop_domain": SHOP,
                    "Handle": key,
                    "product_id": str(product.get("id", "")),
                    "media_id": img["image_url"].split("/")[-1].split("?")[0],
                    "image_location": "public Shopify CDN",
                    "image_number": n,
                    "variant": "",
                    "image_url": img["image_url"],
                    "image_url_export": "UNKNOWN_WITHOUT_SHOPIFY_EXPORT",
                    "identity_status": "PUBLIC_JSON_MATCHED_TO_HANDLE",
                    "shared_media_references": "",
                    "viewed_status": "VIEWED_CONTACT_SHEET_AND_LOCAL_FILE",
                    "viewed_at": NOW,
                    "observed_visual_details": alt,
                    "alt_current": img.get("alt_current") or "",
                    "alt_proposed": alt,
                    "alt_action": "SET",
                    "review_status": "NEEDS_REVIEW",
                    "revision": "1",
                    "approved_by": "",
                    "approved_at": "",
                    "approved_fields": "",
                    "approved_revision": "",
                    "evidence_file_or_reference": img["file"],
                    "issues": "alt_current_empty_in_public_json",
                }
            )

        seo = {h: "" for h in seo_headers}
        seo.update(
            {
                "schema_version": "2.4",
                "prompt_version": "2.4",
                "run_id": RUN_ID,
                "batch_id": BATCH_ID,
                "shop_domain": SHOP,
                "product_key": key,
                "Handle": key,
                "product_id": str(product.get("id", "")),
                "product_gid": f"gid://shopify/Product/{product.get('id', '')}",
                "identity_status": "PUBLIC_JSON_MATCHED_TO_HANDLE",
                "source_export_ref": item.get("source_json", ""),
                "source_exported_at": "UNKNOWN_PUBLIC_JSON_CAPTURE",
                "locale": "en-US",
                "market": "US",
                "url": item["url"],
                "canonical_url": item["url"],
                "product_type": product.get("product_type", ""),
                "title_current": product.get("title", ""),
                "h1_current": product.get("title", ""),
                "title_proposed": d["title"],
                "title_action": "SET",
                "h1_mapping_status": "NEEDS_THEME_CONFIRMATION",
                "meta_title_current": "",
                "meta_title_source_state": "UNKNOWN_WITHOUT_SHOPIFY_EXPORT",
                "rendered_title_current": product.get("title", ""),
                "theme_title_suffix": "UNKNOWN",
                "meta_title_seo": d["title"],
                "meta_title_action": "SET",
                "meta_title_chars": len(d["title"]),
                "meta_description_current": "",
                "meta_description_source_state": "UNKNOWN_WITHOUT_SHOPIFY_EXPORT",
                "rendered_meta_description_current": "UNKNOWN_WITHOUT_RENDERED_PAGE_META_CAPTURE",
                "meta_description_seo": d["meta"],
                "meta_description_action": "SET",
                "meta_description_chars": len(d["meta"]),
                "primary_keyword": d["primary"],
                "secondary_keywords": d["secondary"],
                "long_tail_candidates": f"{d['primary']} for entryway; {d['theme']} rug; Halloween floor mat",
                "keyword_strategy": "Use motif-specific product terms; reserve broad Halloween round rug or Halloween area rug for collections.",
                "season": "HALLOWEEN",
                "keyword_demand_evidence": "SERP observed for vortex, haunted staircase, ghost, Halloween optical illusion, and personalized Halloween rug terms; no paid volume used.",
                "keyword_serp_fit": "SUPPORTED_BY_PRODUCT_AND_CATEGORY_SERP_OBSERVATION",
                "meta_keyword": d["primary"],
                "search_intent": "Commercial investigation / product purchase",
                "keyword_validation_status": "SERP_OBSERVED_NO_VOLUME",
                "keyword_selection_reason": d["obs"],
                "buyer_research_refs": research_id,
                "buyer_search_summary": "Halloween shoppers can use this motif-specific rug as a spooky seasonal floor centerpiece for rooms, entryways, parties, or personal decor.",
                "baseline_status": "PUBLIC_DATA_ONLY",
                "change_scope": "SEO title; meta description; product title/H1 proposal; description proposal; image alt proposal",
                "mapping_status": "NEEDS_REVIEW",
                "mapping_version": "2.4",
                "description_current_html": product.get("body_html", ""),
                "description_changes_needed": "Replace generic duplicate wording with motif-specific, evidence-backed copy.",
                "description_proposed": strip_html(product_html(d)),
                "description_proposed_html": product_html(d),
                "description_change_mode": "REWRITE_FROM_PUBLIC_FACTS",
                "description_target_section": "Full product body",
                "description_action": "SET",
                "revision": "1",
                "review_status": "NEEDS_REVIEW",
                "review_reason": "Research proposal only; Shopify admin/export values are unknown and require review before approval.",
                "approved_by": "",
                "approved_at": "",
                "approved_fields": "",
                "approved_revision": "",
                "processing_status": "DRAFTED",
                "evidence_status": "COMPLETE_PUBLIC_ONLY",
                "content_qa_status": "NOT_RUN",
                "field_evidence_map": "title/meta/description/alt derived from public JSON, viewed images, and SERP observation.",
                "evidence_id": evidence_id,
                "issues": "Admin SEO fields unknown without Shopify export; QA not run; not approved for deployment.",
            }
        )
        seo.update(img_fields)
        b3_seo.append(seo)

        b3_evd.append(
            {
                "evidence_id": evidence_id,
                "product_url": item["url"],
                "reviewed_at": NOW,
                "sources_accessed": "; ".join([item["url"], item.get("source_json", ""), *SERP_REFS]),
                "current_H1": product.get("title", ""),
                "current_meta_title": "UNKNOWN_WITHOUT_RENDERED_PAGE_META_CAPTURE",
                "short_source_excerpt": strip_html(product.get("body_html", ""))[:600],
                "verified_product_facts": f"{d['obs']}; product type={product.get('product_type')}; variants={len(product.get('variants', []))}; images={len(imgs)}.",
                "gallery_image_count": len(imgs),
                "images_viewed_count": len(imgs),
                "image_audit_references": "; ".join([f"{evidence_id}-IMG{int(i['image_number']):02d}" for i in imgs]),
                "SERP_evidence_references": "; ".join(SERP_REFS),
                "buyer_research_references": research_id,
                "factual_conflicts": "None observed in public data; admin values unknown.",
                "processing_status": "DRAFTED",
                "confidence_and_reason": "MEDIUM: product facts and images verified from public data; keyword demand has SERP observation but no volume data.",
                "fact_to_source_map": "product facts=products_page_1.json; visuals=B003 local images/contact sheets; keyword context=SERP references.",
                "proposed_field_to_fact_map": "motif to title/meta/alt; backing/care claims only used when visible in product images.",
            }
        )

        b3_buyer.append(
            {
                "research_id": research_id,
                "product_key": key,
                "supporting_fact_ids": evidence_id,
                "purchase_context": "Seasonal Halloween decorating, novelty rug shopping, or personalized Halloween home decor.",
                "jtbd_statement": f"When setting up Halloween decor, the buyer wants a {d['theme']} rug that makes the room feel themed and visually surprising without installing permanent decor.",
                "functional_motivation": "Add a printed rug with a clear Halloween motif and non-slip backing or care features shown in product imagery.",
                "emotional_social_motivation": "Make the room feel spooky, fun, personalized, and more memorable for guests or family.",
                "purchase_concerns": "Design clarity, size fit, whether the illusion reads well, slipping, cleaning, and personalization accuracy where relevant.",
                "customer_language": "3D Halloween rug; optical illusion rug; ghost rug; spooky round rug; personalized Halloween rug",
                "language_origin": "SERP_AND_MARKETPLACE_PRODUCT_LANGUAGE",
                "source_refs": "; ".join(SERP_REFS),
                "source_scope": "CATEGORY",
                "evidence_excerpt": "Paraphrased from public SERP/category results; no direct customer-review quote used.",
                "observed_at": NOW,
                "market": "US",
                "source_language": "en",
                "research_status": "SUPPORTED_CATEGORY_LANGUAGE_NO_VOLUME",
                "limitations": "No paid search volume or direct Chillgen customer review source checked.",
                "seo_application": "Prefer motif-specific product keywords and avoid assigning broad seasonal head terms to each product.",
            }
        )

        for kw, role, target, reason in [
            (d["primary"], "PRIMARY", item["url"], "Best fit for observed motif and product-level commercial intent."),
            (d["secondary"].split(", ")[0], "SECONDARY", item["url"], "Relevant motif synonym supporting the page without replacing primary."),
            ("Halloween round rug" if "witch cat" not in d["theme"] else "personalized Halloween rug", "REJECTED", "https://chillgen.com/collections/halloween", "Broad term overlaps many products; better treated as collection-level target."),
        ]:
            b3_kw.append(
                {
                    "keyword": kw,
                    "product_key": key,
                    "buyer_research_refs": research_id if role != "REJECTED" else "",
                    "query_origin": "SERP_OBSERVED_AND_AGENT_EXPANSION",
                    "semantic_cluster": d["primary"],
                    "intent": "Commercial",
                    "target_page_type": "PRODUCT" if role != "REJECTED" else "COLLECTION",
                    "target_url": target,
                    "keyword_role": role,
                    "decision_reason": reason,
                    "supporting_fact_ids": evidence_id,
                    "demand_evidence": "SERP pages observed; no search volume source used.",
                    "season": "HALLOWEEN",
                    "research_period": "2026-09-07",
                    "validation_source": "; ".join(SERP_REFS),
                    "checked_at": NOW,
                    "representative_SERP_URLs": "; ".join(SERP_REFS[:3]),
                    "possible_overlap_with_other_products": "YES" if role == "REJECTED" else "MEDIUM_WITH_SIMILAR_HALLOWEEN_PRODUCTS",
                    "mapping_reason": "Product-specific motif retained for product; broad seasonal term left for collection.",
                    "mapping_status": "NEEDS_REVIEW",
                    "mapping_version": "2.4",
                }
            )

    write_csv(ROOT / "batches/B003_SEO_Products.csv", b3_seo, seo_headers)
    write_csv(ROOT / "batches/B003_Image_Audit.csv", b3_img, image_headers)
    write_csv(ROOT / "batches/B003_Product_Evidence.csv", b3_evd, evidence_headers)
    write_csv(ROOT / "keyword_research.csv", prev_kw + b3_kw, keyword_headers)
    with (ROOT / "buyer_search_research.jsonl").open("w", encoding="utf-8") as f:
        for row in prev_buyer + b3_buyer:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")

    inv, inv_headers = read_csv_clean(ROOT / "inventory.csv")
    keys = {row["product_key"] for row in b3_seo}
    for row in inv:
        if row["product_key"] in keys:
            row["assigned_batch"] = BATCH_ID
            row["batch_status"] = "COMPLETED"
            row["research_status"] = "DRAFTED"
            row["review_status"] = "NEEDS_REVIEW"
    write_csv(ROOT / "inventory.csv", inv, inv_headers)

    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "SEO_Products", prev_seo + b3_seo, seo_headers)
    write_sheet(wb, "Image_Audit", prev_img + b3_img, image_headers)
    write_sheet(wb, "Product_Evidence", prev_evd + b3_evd, evidence_headers)
    write_sheet(wb, "Keyword_Map", prev_kw + b3_kw, keyword_headers)
    write_sheet(wb, "Buyer_Search_Research", prev_buyer + b3_buyer, buyer_headers)
    readme = [
        {"metric": "scope", "value": "B001 through B003: 30 public products for chillgen.com, US/en-US", "definition": "Research proposal workbook only; all product and image proposals are NEEDS_REVIEW."},
        {"metric": "sources", "value": "Chillgen public products.json, product URLs, downloaded CDN images, contact sheets, SERP snippets and marketplace/category pages.", "definition": "Admin SEO fields, current stored meta fields, and Shopify export media IDs remain unknown without Shopify CSV/export."},
        {"metric": "quality_status", "value": "DRAFTED_PUBLIC_ONLY", "definition": "239/239 images viewed via contact sheets across B001-B003; QA not run; no Shopify payload created."},
        {"metric": "deployment_status", "value": "NOT_APPROVED_NOT_DEPLOYABLE", "definition": "Do not import this workbook. Run QA and obtain explicit approval before any Shopify/Matrixify payload."},
        {"metric": "next_step", "value": "Stop after B003 and wait for user confirmation before B004.", "definition": "Required by project instruction."},
    ]
    write_sheet(wb, "README_QA", readme, ["metric", "value", "definition"])
    out = RESULTS / "batches" / f"SEO_Product_Optimization_through_{BATCH_ID}.xlsx"
    wb.save(out)
    reopened = load_workbook(out, read_only=True, data_only=False)
    counts = {name: reopened[name].max_row - 1 for name in reopened.sheetnames}

    summary = ROOT / "batches/B003_summary.md"
    summary.write_text(
        "\n".join(
            [
                "# B003 Summary",
                "",
                f"- Run: {RUN_ID}",
                f"- Shop: {SHOP}",
                "- Scope: 10 products, 80 images",
                "- Cumulative workbook scope: B001-B003, 30 products, 239 images",
                "- Status: NEEDS_REVIEW; QA_NOT_RUN; NOT_APPROVED; NO_DEPLOYMENT_PAYLOAD",
                f"- Workbook: {out}",
                f"- Generated at: {NOW}",
                "",
                "## Product Keys",
                *[f"- {row['product_key']}" for row in b3_seo],
            ]
        ),
        encoding="utf-8",
    )

    progress = json.loads((ROOT / "progress.json").read_text())
    progress.update(
        {
            "batch_id": BATCH_ID,
            "batch_product_keys": [row["product_key"] for row in b3_seo],
            "batch_status": "COMPLETED_NEEDS_REVIEW",
            "awaiting_confirmation": True,
            "continuation_confirmation_ref": "USER_CONFIRMED_CONTINUE_B003",
            "current_product_key": None,
            "current_stage": "B003_COMPLETE_AWAITING_USER_CONFIRMATION",
            "images_completed": progress.get("images_completed", []) + [f"{m['product_key']}#img_{int(m['image_number']):02d}" for m in manifest],
            "last_saved_at": NOW,
        }
    )
    progress["artifact_paths"].update(
        {
            "latest_batch_workbook": str(out),
            "b003_seo_products_csv": str(ROOT / "batches/B003_SEO_Products.csv"),
            "b003_image_audit_csv": str(ROOT / "batches/B003_Image_Audit.csv"),
            "b003_product_evidence_csv": str(ROOT / "batches/B003_Product_Evidence.csv"),
            "b003_summary": str(summary),
            "b003_image_manifest": str(ROOT / "evidence/images/B003_manifest.json"),
            "b003_contact_sheets": str(ROOT / "evidence/images/B003_contact_sheets"),
        }
    )
    tmp = ROOT / "progress.json.tmp"
    tmp.write_text(json.dumps(progress, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(ROOT / "progress.json")

    print(json.dumps({"workbook": str(out), "sheet_counts": counts, "summary": str(summary)}, indent=2))


if __name__ == "__main__":
    main()

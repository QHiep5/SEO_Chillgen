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
BATCH_ID = "B004"
NOW = datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()

SERP_REFS = [
    "https://www.etsy.com/market/cat_rug_halloween",
    "https://www.etsy.com/listing/1271501672/personalized-rug-with-name-custom",
    "https://www.etsy.com/listing/4555542549/personalized-witch-halloween-rug-custom",
    "https://www.ebay.com/b/Gothic-Area-Rug-Area-Rugs/262983/bn_95136494",
    "https://ruggable.com/collections/halloween",
    "https://chillgen.com/collections/halloween",
]

DESIGNS = {
    "personalized-witch-area-rug-with-name-custom-halloween-r-design-05": {
        "theme": "skull and red roses",
        "title": "Personalized Skull Rose Halloween Rug",
        "meta": "Personalize a skull rose Halloween rug with a name, gothic black-and-red artwork, easy-clean visuals, and indoor area rug styling.",
        "primary": "personalized skull Halloween rug",
        "secondary": "skull rose rug, gothic Halloween area rug, custom name Halloween rug",
        "obs": "rectangular black rug with large skull, red roses, ornate scroll pattern and personalized name Sophia; size, pet-friendly and cleaning graphics viewed",
    },
    "personalized-witch-area-rug-with-name-custom-halloween-r-design-09": {
        "theme": "skeleton grave",
        "title": "Personalized Skeleton Grave Halloween Rug",
        "meta": "Create a personalized skeleton grave Halloween rug with cracked earth artwork, custom name, gothic decor style, and easy-clean visuals.",
        "primary": "personalized skeleton Halloween rug",
        "secondary": "skeleton grave rug, gothic Halloween rug, custom name rug",
        "obs": "rectangular cracked-earth rug with skeleton skull and hand emerging from a dark grave, autumn leaves and personalized name Emily",
    },
    "personalized-witch-area-rug-with-name-custom-halloween-r-design-04": {
        "theme": "floral sugar skull",
        "title": "Personalized Floral Skull Halloween Rug",
        "meta": "Add a personalized floral skull Halloween rug with gothic roses, lilies, custom name artwork, and indoor easy-clean decor images.",
        "primary": "personalized floral skull rug",
        "secondary": "sugar skull Halloween rug, gothic floral rug, custom Halloween area rug",
        "obs": "rectangular black rug with cream skull face, burgundy roses, pale lilies, custom name Sophia and care/lifestyle graphics",
    },
    "personalized-witch-area-rug-with-name-custom-halloween-r-design-02": {
        "theme": "stained glass witch silhouette",
        "title": "Personalized Stained Glass Witch Rug",
        "meta": "Choose a personalized stained glass witch rug with moonlit silhouette art, bright purple and yellow panels, and custom name detail.",
        "primary": "personalized stained glass witch rug",
        "secondary": "witch silhouette rug, colorful witch Halloween rug, custom witch area rug",
        "obs": "rectangular rug with witch silhouette on a broom against moonlit stained-glass style panels and personalized name Sophia",
    },
    "personalized-witch-area-rug-with-name-custom-halloween-r-design-06": {
        "theme": "moonlit witch path",
        "title": "Personalized Moonlit Witch Halloween Rug",
        "meta": "Personalize a moonlit witch Halloween rug with a name, purple night path, pumpkins, lanterns, and indoor seasonal styling.",
        "primary": "personalized witch Halloween rug",
        "secondary": "moonlit witch rug, gothic witch area rug, custom Halloween rug",
        "obs": "rectangular purple rug with moonlit witch silhouette, spooky path, pumpkins, lanterns, spell book, potion bottle and custom name Sophia",
    },
    "personalized-witch-area-rug-with-name-custom-halloween-r-design-08": {
        "theme": "graveyard skeleton",
        "title": "Personalized Graveyard Skeleton Rug",
        "meta": "Make a spooky room statement with a personalized graveyard skeleton rug, cracked stone artwork, pumpkins, lanterns, and custom name.",
        "primary": "personalized graveyard skeleton rug",
        "secondary": "Halloween skeleton area rug, gothic graveyard rug, custom name Halloween rug",
        "obs": "rectangular rug with skeleton crawling from cracked graveyard ground, purple cloak, pumpkins, lanterns and personalized name Benjamin",
    },
    "personalized-witch-area-rug-with-name-custom-halloween-r-design-10": {
        "theme": "dark witch broom",
        "title": "Personalized Dark Witch Broom Rug",
        "meta": "Decorate with a personalized dark witch broom rug featuring gothic leaves, crescent moon, custom name, and easy-clean indoor visuals.",
        "primary": "personalized witch broom rug",
        "secondary": "witchy area rug, gothic witch rug, custom Halloween rug",
        "obs": "rectangular dark brown and black rug with witch broom, crescent moon, autumn leaves, ornate border and personalized name Olivia",
    },
    "personalized-witch-area-rug-with-name-custom-halloween-r-design-01": {
        "theme": "raven gothic floral",
        "title": "Personalized Raven Gothic Halloween Rug",
        "meta": "Personalize a raven gothic Halloween rug with moody floral artwork, custom name detail, and indoor area rug lifestyle images.",
        "primary": "personalized raven Halloween rug",
        "secondary": "gothic raven rug, dark floral Halloween rug, custom name area rug",
        "obs": "rectangular dark rug with large black raven, burgundy botanical leaves, gray background and personalized name Sophia",
    },
    "personalized-witch-area-rug-with-name-custom-halloween-r-design-12": {
        "theme": "ghost reading room",
        "title": "Personalized Ghost Reading Halloween Rug",
        "meta": "Choose a personalized ghost reading Halloween rug with cozy library artwork, pumpkin accents, custom name, and easy-clean visuals.",
        "primary": "personalized ghost Halloween rug",
        "secondary": "ghost reading rug, cozy Halloween area rug, custom ghost rug",
        "obs": "rectangular rug with white ghost reading in a dark room, books, candles, pumpkin, framed border and personalized name Tiffany",
    },
    "personalized-witch-area-rug-with-name-custom-halloween-r-design-11": {
        "theme": "autumn ghost reading",
        "title": "Personalized Autumn Ghost Halloween Rug",
        "meta": "Add a personalized autumn ghost Halloween rug with warm pumpkin artwork, reading ghost scene, custom name, and indoor decor mockups.",
        "primary": "personalized autumn ghost rug",
        "secondary": "ghost Halloween area rug, pumpkin ghost rug, custom name Halloween rug",
        "obs": "rectangular warm brown rug with ghost reading under a tree, jack-o-lantern moon, pumpkins, books and personalized name Tiffany",
    },
}


def read_csv(path):
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


def description_html(design):
    return (
        f"<p>{design['meta']}</p><h3>Why it works</h3><ul>"
        f"<li>{design['obs'].capitalize()}.</li>"
        "<li>Public page metadata and product images support the Halloween, personalized-name, and indoor area rug positioning.</li>"
        "<li>Use the visible care and backing claims only where the product page or infographic supports them.</li>"
        "</ul>"
    )


def alt_for(theme, n):
    notes = {
        1: f"lifestyle view of personalized {theme} Halloween rug",
        2: f"alternate room view of personalized {theme} Halloween rug",
        3: f"personalized {theme} Halloween rug staged by fireplace",
        4: f"size information for personalized {theme} rug",
        5: f"kid and pet friendly graphic for personalized {theme} rug",
        6: f"easy-clean care graphic for personalized {theme} rug",
        7: f"lighting color note for personalized {theme} rug",
        8: f"personalized {theme} rug under dining table",
    }
    if "skeleton grave" in theme:
        notes[2] = f"personalized {theme} rug staged near sofa"
        notes[3] = f"personalized {theme} rug in bedroom"
        notes[4] = f"personalized {theme} rug staged by fireplace"
        notes[5] = f"personalized {theme} rug under dining table"
        notes[6] = f"kid and pet friendly graphic for personalized {theme} rug"
        notes[7] = f"easy-clean care graphic for personalized {theme} rug"
        notes[8] = f"size information for personalized {theme} rug"
    return notes.get(int(n), f"personalized {theme} Halloween rug image {n}")


def autosize(ws, max_width=60):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = max(12, min(max_width, max(len(str(c.value or "")) for c in col) + 2))
        ws.column_dimensions[letter].width = width
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
    batch = json.loads((ROOT / "evidence/products/B004_products.json").read_text())
    page_meta = {r["product_key"]: r for r in json.loads((ROOT / "evidence/products/B004_page_meta.json").read_text())}
    page = json.loads((ROOT / "evidence/products/products_page_1.json").read_text())
    products = page.get("products", page) if isinstance(page, dict) else page
    by_handle = {p["handle"]: p for p in products}
    manifest = json.loads((ROOT / "evidence/images/B004_manifest.json").read_text())
    by_images = {}
    for img in manifest:
        by_images.setdefault(img["product_key"], []).append(img)

    prev_seo, prev_img, prev_evd = [], [], []
    seo_headers = image_headers = evidence_headers = []
    for batch_id in ["B001", "B002", "B003"]:
        rows, headers = read_csv(ROOT / f"batches/{batch_id}_SEO_Products.csv")
        prev_seo.extend(rows)
        seo_headers = headers
        rows, headers = read_csv(ROOT / f"batches/{batch_id}_Image_Audit.csv")
        prev_img.extend(rows)
        image_headers = headers
        rows, headers = read_csv(ROOT / f"batches/{batch_id}_Product_Evidence.csv")
        prev_evd.extend(rows)
        evidence_headers = headers
    prev_kw, keyword_headers = read_csv(ROOT / "keyword_research.csv")
    with (ROOT / "buyer_search_research.jsonl").open(encoding="utf-8") as f:
        prev_buyer = [json.loads(line) for line in f if line.strip()]
    buyer_headers = list(prev_buyer[0].keys())

    b4_seo, b4_img, b4_evd, b4_kw, b4_buyer = [], [], [], [], []
    for item in batch:
        key = item["product_key"]
        product = by_handle[key]
        design = DESIGNS[key]
        meta = page_meta[key]
        evidence_id = f"{BATCH_ID}-{item['inventory_order']:03d}"
        research_id = f"{BATCH_ID}-{item['inventory_order']:03d}-BR01"
        imgs = sorted(by_images[key], key=lambda x: int(x["image_number"]))
        img_fields = {}

        for img in imgs:
            n = int(img["image_number"])
            alt = alt_for(design["theme"], n)
            img_fields[f"img_{n}_link"] = img["image_url"]
            img_fields[f"img_{n}_alt_current"] = img.get("alt_current") or ""
            img_fields[f"img_{n}_alt"] = alt
            img_fields[f"img_{n}_alt_action"] = "SET"
            b4_img.append(
                {
                    "evidence_id": evidence_id,
                    "shop_domain": SHOP,
                    "Handle": key,
                    "product_id": str(product.get("id", "")),
                    "media_id": img["image_url"].split("/")[-1].split("?")[0],
                    "image_location": "GALLERY_PUBLIC_CDN",
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
                "source_export_ref": f"{item.get('source_json', '')}; {meta['html_file']}",
                "source_exported_at": meta["captured_at"],
                "locale": "en-US",
                "market": "US",
                "url": item["url"],
                "canonical_url": meta.get("canonical") or item["url"],
                "product_type": product.get("product_type", ""),
                "title_current": product.get("title", ""),
                "h1_current": meta.get("h1", ""),
                "title_proposed": design["title"],
                "title_action": "SET",
                "h1_mapping_status": "RENDERED_H1_CAPTURED_THEME_CONFIRMATION_STILL_NEEDED_FOR_DEPLOY",
                "meta_title_current": "",
                "meta_title_source_state": "UNKNOWN_WITHOUT_SHOPIFY_EXPORT",
                "rendered_title_current": meta.get("rendered_title", ""),
                "theme_title_suffix": "NONE_OR_NOT_OBSERVED",
                "meta_title_seo": design["title"],
                "meta_title_action": "SET",
                "meta_title_chars": len(design["title"]),
                "meta_description_current": "",
                "meta_description_source_state": "UNKNOWN_WITHOUT_SHOPIFY_EXPORT",
                "rendered_meta_description_current": meta.get("rendered_meta_description", ""),
                "meta_description_seo": design["meta"],
                "meta_description_action": "SET",
                "meta_description_chars": len(design["meta"]),
                "primary_keyword": design["primary"],
                "secondary_keywords": design["secondary"],
                "long_tail_candidates": f"{design['primary']}; {design['secondary']}; personalized Halloween rug with name",
                "keyword_strategy": "LONG_TAIL_CANDIDATE",
                "season": "HALLOWEEN",
                "keyword_demand_evidence": "SERP_ONLY: public search results show personalized witch/cat/Halloween rug pages; no paid search volume or Search Console source used.",
                "keyword_serp_fit": "SERP_ONLY_PRODUCT_AND_MARKETPLACE_RESULTS",
                "meta_keyword": design["primary"],
                "search_intent": "COMMERCIAL_INVESTIGATION",
                "keyword_validation_status": "SERP_ONLY_NO_VOLUME",
                "keyword_selection_reason": f"Selected because product imagery and rendered page facts support {design['theme']} and personalized Halloween area rug intent.",
                "buyer_research_refs": research_id,
                "buyer_search_summary": "Buyer likely wants a personalized Halloween area rug with a name and gothic/witch motif for seasonal indoor decorating or gifting; treated as hypothesis supported by category SERP language, not direct Chillgen customer data.",
                "baseline_status": "PUBLIC_RENDERED_PAGE_AND_JSON_CAPTURED",
                "change_scope": "CLARIFICATION",
                "mapping_status": "NEEDS_REVIEW",
                "mapping_version": "2.4",
                "description_current_html": product.get("body_html", ""),
                "description_changes_needed": "Make current generic witch-rug copy more motif-specific while retaining only visible/public claims.",
                "description_proposed": strip_html(description_html(design)),
                "description_proposed_html": description_html(design),
                "description_change_mode": "REPLACE_ALL",
                "description_target_section": "Full product body",
                "description_action": "SET",
                "revision": "1",
                "review_status": "NEEDS_REVIEW",
                "review_reason": "Research proposal only; rendered page captured but admin/export SEO values and media export URLs remain unknown.",
                "approved_by": "",
                "approved_at": "",
                "approved_fields": "",
                "approved_revision": "",
                "processing_status": "DRAFTED",
                "evidence_status": "COMPLETE_PUBLIC_RENDERED_PAGE_AND_IMAGES",
                "content_qa_status": "NOT_RUN",
                "field_evidence_map": "title/meta/description/alt derived from rendered page meta, public JSON, viewed images, and SERP-only validation.",
                "evidence_id": evidence_id,
                "issues": "Admin SEO fields unknown without Shopify export; QA not run; not approved for deployment.",
            }
        )
        seo.update(img_fields)
        b4_seo.append(seo)

        b4_evd.append(
            {
                "evidence_id": evidence_id,
                "product_url": item["url"],
                "reviewed_at": NOW,
                "sources_accessed": "; ".join([item["url"], item.get("source_json", ""), meta["html_file"], *SERP_REFS]),
                "current_H1": meta.get("h1", ""),
                "current_meta_title": meta.get("rendered_title", ""),
                "short_source_excerpt": strip_html(product.get("body_html", ""))[:600],
                "verified_product_facts": f"{design['obs']}; HTTP {meta['http_status']}; canonical={meta.get('canonical')}; product type={product.get('product_type')}; variants={len(product.get('variants', []))}; images={len(imgs)}.",
                "gallery_image_count": len(imgs),
                "images_viewed_count": len(imgs),
                "image_audit_references": "; ".join([f"{evidence_id}-IMG{int(i['image_number']):02d}" for i in imgs]),
                "SERP_evidence_references": "; ".join(SERP_REFS),
                "buyer_research_references": research_id,
                "factual_conflicts": "None observed in public data; admin values unknown.",
                "processing_status": "DRAFTED",
                "confidence_and_reason": "MEDIUM: rendered page, product JSON, and images verified; keyword demand has SERP-only evidence and no volume data.",
                "fact_to_source_map": "product facts=products_page_1.json and B004_pages HTML; visuals=B004 local images/contact sheets; keyword context=SERP references.",
                "proposed_field_to_fact_map": "motif and personalization to title/meta/alt; care/backing only where page/image evidence supports.",
            }
        )

        b4_buyer.append(
            {
                "research_id": research_id,
                "product_key": key,
                "supporting_fact_ids": evidence_id,
                "purchase_context": "Seasonal Halloween decorating, gothic room styling, or personalized-name rug gifting.",
                "jtbd_statement": f"When decorating for Halloween or a gothic-themed room, the buyer wants a personalized {design['theme']} rug to make the space feel seasonal and personal.",
                "functional_motivation": "Add an indoor area rug with a custom name and a visible Halloween/gothic motif.",
                "emotional_social_motivation": "Make a room feel distinctive, spooky, giftable, and more personal through the name detail.",
                "purchase_concerns": "Design clarity, name accuracy, size fit, cleaning, backing grip, and whether the gothic style fits the room.",
                "customer_language": "personalized Halloween rug; custom name rug; witch rug; gothic area rug; black cat Halloween rug",
                "language_origin": "AGENT_HYPOTHESIS_WITH_CATEGORY_SERP_PARAPHRASE",
                "source_refs": "; ".join(SERP_REFS),
                "source_scope": "CATEGORY",
                "evidence_excerpt": "Paraphrased from public search/category results; no direct Chillgen customer quote or volume source used.",
                "observed_at": NOW,
                "market": "US",
                "source_language": "en",
                "research_status": "HYPOTHESIS",
                "limitations": "No Search Console, internal search, direct review corpus, or paid volume tool used for this product.",
                "seo_application": "Use motif-specific personalized Halloween rug language and keep broad head terms for collection-level decisions.",
            }
        )

        for kw, role, target, reason in [
            (design["primary"], "PRIMARY", item["url"], "Best product-level match for observed motif and personalized rug intent."),
            (design["secondary"].split(", ")[0], "SECONDARY", item["url"], "Useful motif synonym supported by the page and image evidence."),
            ("personalized Halloween rug", "REJECTED", "https://chillgen.com/collections/halloween", "Too broad for every product in this group; better considered at collection/category level."),
        ]:
            b4_kw.append(
                {
                    "keyword": kw,
                    "product_key": key,
                    "buyer_research_refs": research_id if role != "REJECTED" else "",
                    "query_origin": "AGENT_PROPOSED",
                    "semantic_cluster": design["primary"],
                    "intent": "COMMERCIAL_INVESTIGATION",
                    "target_page_type": "PRODUCT" if role != "REJECTED" else "COLLECTION",
                    "target_url": target,
                    "keyword_role": role,
                    "decision_reason": reason,
                    "supporting_fact_ids": evidence_id,
                    "demand_evidence": "SERP_ONLY, no volume verified.",
                    "season": "HALLOWEEN",
                    "research_period": "2026-09-07",
                    "validation_source": "; ".join(SERP_REFS),
                    "checked_at": NOW,
                    "representative_SERP_URLs": "; ".join(SERP_REFS[:4]),
                    "possible_overlap_with_other_products": "YES" if role == "REJECTED" else "MEDIUM_WITH_SIMILAR_PERSONALIZED_WITCH_RUG_PRODUCTS",
                    "mapping_reason": "Motif-specific keyword retained for product; broad personalized Halloween term reserved for collection review.",
                    "mapping_status": "NEEDS_REVIEW",
                    "mapping_version": "2.4",
                }
            )

    write_csv(ROOT / "batches/B004_SEO_Products.csv", b4_seo, seo_headers)
    write_csv(ROOT / "batches/B004_Image_Audit.csv", b4_img, image_headers)
    write_csv(ROOT / "batches/B004_Product_Evidence.csv", b4_evd, evidence_headers)
    write_csv(ROOT / "keyword_research.csv", prev_kw + b4_kw, keyword_headers)
    with (ROOT / "buyer_search_research.jsonl").open("w", encoding="utf-8") as f:
        for row in prev_buyer + b4_buyer:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")

    inv, inv_headers = read_csv(ROOT / "inventory.csv")
    keys = {row["product_key"] for row in b4_seo}
    for row in inv:
        if row["product_key"] in keys:
            row["assigned_batch"] = BATCH_ID
            row["batch_status"] = "COMPLETED"
            row["research_status"] = "DRAFTED"
            row["review_status"] = "NEEDS_REVIEW"
    write_csv(ROOT / "inventory.csv", inv, inv_headers)

    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "SEO_Products", prev_seo + b4_seo, seo_headers)
    write_sheet(wb, "Image_Audit", prev_img + b4_img, image_headers)
    write_sheet(wb, "Product_Evidence", prev_evd + b4_evd, evidence_headers)
    write_sheet(wb, "Keyword_Map", prev_kw + b4_kw, keyword_headers)
    write_sheet(wb, "Buyer_Search_Research", prev_buyer + b4_buyer, buyer_headers)
    readme = [
        {"metric": "scope", "value": "B001 through B004: 40 public products for chillgen.com, US/en-US", "definition": "Research proposal workbook only; all product and image proposals are NEEDS_REVIEW."},
        {"metric": "B004_method_upgrade", "value": "B004 captured product page HTML/meta before SEO drafting.", "definition": "B001-B003 remain public JSON/image evidence without full rendered page/meta backfill."},
        {"metric": "sources", "value": "Chillgen public products.json, product URLs, rendered HTML captures, downloaded CDN images, contact sheets, SERP snippets and marketplace/category pages.", "definition": "Admin SEO fields and Shopify export media URLs remain unknown without Shopify CSV/export."},
        {"metric": "quality_status", "value": "DRAFTED_PUBLIC_RENDERED_PAGE_FOR_B004", "definition": "319/319 images viewed via contact sheets across B001-B004; QA not run; no Shopify payload created."},
        {"metric": "deployment_status", "value": "NOT_APPROVED_NOT_DEPLOYABLE", "definition": "Do not import this workbook. Run QA and obtain explicit approval before any Shopify/Matrixify payload."},
        {"metric": "next_step", "value": "Stop after B004 and wait for user confirmation before B005.", "definition": "Required by project instruction."},
    ]
    write_sheet(wb, "README_QA", readme, ["metric", "value", "definition"])
    out = RESULTS / "batches" / f"SEO_Product_Optimization_through_{BATCH_ID}.xlsx"
    wb.save(out)
    reopened = load_workbook(out, read_only=True, data_only=False)
    counts = {name: reopened[name].max_row - 1 for name in reopened.sheetnames}

    summary = ROOT / "batches/B004_summary.md"
    summary.write_text(
        "\n".join(
            [
                "# B004 Summary",
                "",
                f"- Run: {RUN_ID}",
                f"- Shop: {SHOP}",
                "- Scope: 10 products, 80 images",
                "- Cumulative workbook scope: B001-B004, 40 products, 319 images",
                "- Method upgrade: B004 has pre-batch checkpoint plus rendered page/meta HTML capture.",
                "- Status: NEEDS_REVIEW; QA_NOT_RUN; NOT_APPROVED; NO_DEPLOYMENT_PAYLOAD",
                f"- Workbook: {out}",
                f"- Generated at: {NOW}",
                "",
                "## Product Keys",
                *[f"- {row['product_key']}" for row in b4_seo],
            ]
        ),
        encoding="utf-8",
    )

    progress = json.loads((ROOT / "progress.json").read_text())
    progress.update(
        {
            "batch_id": BATCH_ID,
            "batch_product_keys": [row["product_key"] for row in b4_seo],
            "batch_status": "COMPLETED_NEEDS_REVIEW",
            "awaiting_confirmation": True,
            "continuation_confirmation_ref": "USER_CONFIRMED_START_B004",
            "current_product_key": None,
            "current_stage": "B004_COMPLETE_AWAITING_USER_CONFIRMATION",
            "images_completed": progress.get("images_completed", []) + [f"{m['product_key']}#img_{int(m['image_number']):02d}" for m in manifest],
            "last_saved_at": NOW,
        }
    )
    progress["artifact_paths"].update(
        {
            "latest_batch_workbook": str(out),
            "b004_seo_products_csv": str(ROOT / "batches/B004_SEO_Products.csv"),
            "b004_image_audit_csv": str(ROOT / "batches/B004_Image_Audit.csv"),
            "b004_product_evidence_csv": str(ROOT / "batches/B004_Product_Evidence.csv"),
            "b004_summary": str(summary),
            "b004_page_meta": str(ROOT / "evidence/products/B004_page_meta.json"),
            "b004_image_manifest": str(ROOT / "evidence/images/B004_manifest.json"),
            "b004_contact_sheets": str(ROOT / "evidence/images/B004_contact_sheets"),
        }
    )
    tmp = ROOT / "progress.json.tmp"
    tmp.write_text(json.dumps(progress, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(ROOT / "progress.json")

    print(json.dumps({"workbook": str(out), "sheet_counts": counts, "summary": str(summary)}, indent=2))


if __name__ == "__main__":
    main()

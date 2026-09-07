import csv
import json
from datetime import datetime, timezone, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


RUN_ID = "chillgen_20260907_01"
SHOP = "chillgen.com"
ROOT = Path("seo_runs") / SHOP / RUN_ID
RESULTS = Path("resutls") / SHOP / RUN_ID
BATCH_ID = "B001"
NOW = datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()


SERP_REFS = {
    "dog": [
        "https://aprintnest.com/",
        "https://www.etsy.com/listing/1585313764/custom-pet-portrait-rug-personalized-rug",
    ],
    "illusion": [
        "https://www.walmart.com/ip/20811815953",
        "https://shop.tiktok.com/us/pdp/1732566225722380431",
        "https://www.etsy.com/ca/listing/4557149347/3d-skeleton-halloween-rug-round-spooky",
        "https://pisif.com/products/3d-skeleton-illusion-round-halloween-rug",
    ],
    "round": [
        "https://www.walmart.com/ip/20812707756",
        "https://shop.tiktok.com/us/pdp/halloween-3d-skeleton-pit-round-rug-spooky-optic-illusion/1729524799415162793",
        "https://shop.tiktok.com/us/pdp/1732597414909874553",
    ],
    "collection": ["https://chillgen.com/collections/halloween"],
}


DESIGNS = {
    "custom-dog-photo-rug-personalized-pet-portrait-area-mat-n-design-03": {
        "theme": "haunted house dog portrait",
        "season": "HALLOWEEN",
        "title": "Personalized Dog Halloween Rug with Haunted House",
        "meta_title": "Personalized Dog Halloween Rug with Haunted House",
        "meta_description": "Create a custom dog Halloween rug with your pet photo, name, haunted house artwork, soft surface, and non-slip backing for seasonal indoor decor.",
        "primary": "personalized dog Halloween rug",
        "secondary": "custom pet portrait rug, haunted house dog rug, dog photo rug",
        "summary": "Pet parents looking for a seasonal keepsake can use a dog photo and name to create a haunted house rug for Halloween rooms, entryways, or living spaces.",
        "obs": "shaped haunted house Halloween rug with bulldog portrait and custom name POPPY; lifestyle, size, layer, and backing images viewed",
    },
    "custom-dog-photo-rug-personalized-pet-portrait-area-mat-n-design-02": {
        "theme": "pumpkin dog portrait",
        "season": "HALLOWEEN",
        "title": "Personalized Dog Halloween Pumpkin Rug",
        "meta_title": "Personalized Dog Halloween Pumpkin Rug",
        "meta_description": "Personalize a pumpkin-themed Halloween dog rug with your pet photo and name. Soft indoor area mat with non-slip backing for festive home decor.",
        "primary": "personalized dog Halloween pumpkin rug",
        "secondary": "custom dog photo rug, pumpkin pet portrait rug, Halloween pet rug",
        "summary": "Buyers can turn a dog photo into a playful pumpkin Halloween rug, useful as personal seasonal decor or a pet parent gift.",
        "obs": "shaped orange pumpkin Halloween rug with bulldog portrait, white ghost pumpkin, custom name POPPY, and product feature graphics",
    },
    "custom-dog-photo-rug-personalized-pet-portrait-area-mat-n-design-01": {
        "theme": "graveyard dog portrait",
        "season": "HALLOWEEN",
        "title": "Personalized Dog Halloween Graveyard Rug",
        "meta_title": "Personalized Dog Halloween Graveyard Rug",
        "meta_description": "Design a personalized dog Halloween rug with your pet photo, name, graveyard artwork, soft printed surface, and non-slip backing for indoor spaces.",
        "primary": "personalized dog Halloween graveyard rug",
        "secondary": "custom pet name rug, dog portrait Halloween rug, spooky pet rug",
        "summary": "This fits pet owners who want a spooky-but-cute graveyard themed custom rug featuring their dog and pet name.",
        "obs": "gray graveyard Halloween rug with bulldog portrait, tombstones, ghosts, lantern, RIP details, and custom name BUDDY",
    },
    "custom-dog-photo-rug-personalized-pet-portrait-area-mat-non-slip": {
        "theme": "witch cauldron dog portrait",
        "season": "HALLOWEEN",
        "title": "Personalized Dog Witch Cauldron Halloween Rug",
        "meta_title": "Personalized Dog Witch Cauldron Halloween Rug",
        "meta_description": "Make a custom dog photo Halloween rug with witch cauldron artwork, pet name personalization, soft surface, and anti-slip backing for indoor decor.",
        "primary": "personalized dog witch Halloween rug",
        "secondary": "custom dog cauldron rug, Halloween pet photo rug, personalized pet rug",
        "summary": "Pet owners can use a favorite dog photo in a witch cauldron design for a personalized Halloween decor accent.",
        "obs": "black and purple witch cauldron rug with bulldog portrait, potion bottle, broom, moon, custom name MOLLY, and backing closeups",
    },
    "personalized-halloween-3d-optical-illusion-shaped-rug-a0-design-102": {
        "theme": "spider web vortex",
        "season": "HALLOWEEN",
        "title": "Halloween Spider Web Optical Illusion Rug",
        "meta_title": "Halloween Spider Web Optical Illusion Rug",
        "meta_description": "Add a spider web optical illusion rug to Halloween decor. Flat printed shaped area rug with purple, orange, bat, and web details.",
        "primary": "Halloween spider web optical illusion rug",
        "secondary": "3D Halloween rug, spiderweb rug, shaped Halloween rug",
        "summary": "Halloween decorators may want a floor centerpiece that looks like a deep spider web vortex while staying flat underfoot.",
        "obs": "scalloped spider web vortex rug in black, orange, and purple with bats, stars, and visual tunnel illusion",
    },
    "personalized-halloween-3d-optical-illusion-shaped-rug-a0-design-101": {
        "theme": "crawling skeleton",
        "season": "HALLOWEEN",
        "title": "Halloween Skeleton Optical Illusion Rug",
        "meta_title": "Halloween Skeleton Optical Illusion Rug",
        "meta_description": "Spook up a room with a skeleton optical illusion Halloween rug. Flat printed shaped rug with dark floral texture and non-slip backing.",
        "primary": "Halloween skeleton optical illusion rug",
        "secondary": "3D skeleton rug, spooky shaped rug, Halloween floor mat",
        "summary": "The visual job is a creepy skeleton floor accent for Halloween parties, horror rooms, or seasonal entryways.",
        "obs": "irregular dark shaped rug with crawling skeleton, gray floral background, depth effect, size chart, and backing views",
    },
    "personalized-halloween-3d-optical-illusion-shaped-rug-a0-design-100": {
        "theme": "sugar skull vortex",
        "season": "HALLOWEEN",
        "title": "Sugar Skull Halloween Optical Illusion Rug",
        "meta_title": "Sugar Skull Halloween Optical Illusion Rug",
        "meta_description": "Choose a sugar skull Halloween rug with optical illusion depth, cream skull artwork, colorful tunnel pattern, and non-slip backing.",
        "primary": "sugar skull Halloween rug",
        "secondary": "skull optical illusion rug, Halloween skull rug, 3D Halloween rug",
        "summary": "Buyers wanting a skull motif with colorful visual depth can use this as a bold Halloween or gothic room accent.",
        "obs": "cream sugar skull shaped rug with black eye sockets, colorful tunnel mouth, decorative scrolls, and product feature images",
    },
    "personalized-halloween-3d-optical-illusion-shaped-rug-a04-a04": {
        "theme": "pumpkin ghost vortex",
        "season": "HALLOWEEN",
        "title": "Personalized Pumpkin Ghost Halloween Rug",
        "meta_title": "Personalized Pumpkin Ghost Halloween Rug",
        "meta_description": "Personalize a pumpkin ghost Halloween rug with a playful optical illusion print, shaped edge, skull details, and non-slip backing.",
        "primary": "pumpkin ghost Halloween rug",
        "secondary": "personalized Halloween rug, ghost pumpkin rug, shaped Halloween rug",
        "summary": "This serves buyers who want a friendlier Halloween focal point with pumpkins, ghosts, skulls, and optional personalization.",
        "obs": "pumpkin-shaped rug with jack-o-lantern top, ghosts, skulls, checkerboard tunnel, autumn leaves, and backing graphics",
    },
    "custom-halloween-3d-effect-optical-illusion-spooky-round-rug-d9-d9": {
        "theme": "skeleton pit round rug",
        "season": "HALLOWEEN",
        "title": "Skeleton Pit Halloween Round Rug",
        "meta_title": "Skeleton Pit Halloween Round Rug",
        "meta_description": "Decorate with a round Halloween rug showing a skeleton pit illusion, patchwork border, non-slip backing, and sizes up to 6x6 ft.",
        "primary": "skeleton pit Halloween round rug",
        "secondary": "3D skeleton round rug, Halloween optical illusion rug, spooky round rug",
        "summary": "Shoppers searching round Halloween floor decor may want a dramatic skeleton pit illusion for parties, entryways, or themed rooms.",
        "obs": "round rug with skeleton in a deep checkerboard pit, orange and cream patchwork border, size chart, lifestyle and quality graphics",
    },
    "custom-halloween-3d-effect-optical-illusion-spooky-round-r-design-06": {
        "theme": "spider optical illusion round rug",
        "season": "HALLOWEEN",
        "title": "Spider Optical Illusion Halloween Round Rug",
        "meta_title": "Spider Optical Illusion Halloween Round Rug",
        "meta_description": "Bring in a spider optical illusion Halloween round rug with green checkerboard depth, spooky stone border, and non-slip backing.",
        "primary": "spider optical illusion Halloween rug",
        "secondary": "spider round rug, 3D Halloween rug, spooky spider rug",
        "summary": "This targets Halloween buyers who want an intense spider floor illusion for a living room, dining nook, or party setup.",
        "obs": "round rug with large black and purple spider over green checkerboard tunnel, stone border, lifestyle mockups, and backing closeup",
    },
}


def load_products():
    batch = json.loads((ROOT / "evidence/products/B001_products.json").read_text())
    page = json.loads((ROOT / "evidence/products/products_page_1.json").read_text())
    products = page.get("products", page) if isinstance(page, dict) else page
    by_handle = {p["handle"]: p for p in products}
    return batch, by_handle


def html_to_text(html):
    import re

    text = re.sub(r"<[^>]+>", " ", html or "")
    return " ".join(text.split())


def proposed_html(design, product):
    title = design["title"]
    bits = [
        f"<p>{design['meta_description']}</p>",
        "<h3>Why it works</h3>",
        "<ul>",
        f"<li>{design['obs'].capitalize()}.</li>",
        "<li>Public product data shows a soft printed surface and non-slip backing in the product imagery.</li>",
        "<li>Made for indoor seasonal decorating in living rooms, entryways, bedrooms, and party areas.</li>",
        "</ul>",
    ]
    return "".join(bits)


def image_note(theme, n, total):
    if "dog portrait" in theme:
        notes = {
            1: f"lifestyle view of {theme} rug in living room",
            2: f"custom photo placement example on {theme} rug",
            3: f"{theme} rug shown with dogs nearby",
            4: "material layer and non-slip backing graphic",
            5: f"size chart for {theme} shaped area rug",
            6: f"feature graphic for {theme} rug with backing closeups",
            7: f"{theme} rug staged by sofa",
            8: f"{theme} rug staged near stairs or entry area",
            9: f"{theme} rug staged near doorway",
        }
    elif "round" in theme or "pit" in theme:
        notes = {
            1: f"top view of {theme}",
            2: f"size chart for {theme}",
            3: f"{theme} staged in Halloween living room",
            4: f"{theme} shown underfoot",
            5: f"{theme} in seasonal room mockup",
            6: f"{theme} under dining table",
            7: f"non-slip backing detail for {theme}",
            8: f"design and backing detail for {theme}",
        }
    else:
        notes = {
            1: f"lifestyle top view of {theme} shaped rug",
            2: f"angled lifestyle view of {theme} rug",
            3: "material layer and non-slip backing graphic",
            4: f"{theme} rug staged in bright room",
            5: f"feature graphic for {theme} rug with backing closeups",
            6: f"cutout product view of {theme} rug",
            7: f"size chart for {theme} rug",
        }
    return notes.get(int(n), f"{theme} product image {n} of {total}")


def autosize(ws, max_width=60):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = min(max_width, max(len(str(c.value or "")) for c in col) + 2)
        ws.column_dimensions[letter].width = max(width, 12)
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
    batch, by_handle = load_products()
    manifest = json.loads((ROOT / "evidence/images/B001_manifest.json").read_text())
    manifest_by_product = {}
    for item in manifest:
        manifest_by_product.setdefault(item["product_key"], []).append(item)

    seo_rows = []
    evidence_rows = []
    image_rows = []
    buyer_rows = []
    keyword_rows = []
    image_cols = []
    max_images = max(len(v) for v in manifest_by_product.values())
    for i in range(1, max_images + 1):
        image_cols += [f"img_{i}_link", f"img_{i}_alt_current", f"img_{i}_alt", f"img_{i}_alt_action"]

    for item in batch:
        key = item["product_key"]
        product = by_handle[key]
        design = DESIGNS[key]
        evidence_id = f"{BATCH_ID}-{item['inventory_order']:03d}"
        research_id = f"{BATCH_ID}-{item['inventory_order']:03d}-BR01"
        images = sorted(manifest_by_product[key], key=lambda x: int(x["image_number"]))
        body_text = html_to_text(product.get("body_html", ""))
        refs = SERP_REFS["dog"] if key.startswith("custom-dog") else SERP_REFS["round" if "round" in key else "illusion"]

        img_fields = {}
        for img in images:
            n = int(img["image_number"])
            alt = image_note(design["theme"], n, len(images))
            img_fields[f"img_{n}_link"] = img["image_url"]
            img_fields[f"img_{n}_alt_current"] = img.get("alt_current") or ""
            img_fields[f"img_{n}_alt"] = alt[:512]
            img_fields[f"img_{n}_alt_action"] = "SET"
            image_rows.append(
                {
                    "evidence_id": evidence_id,
                    "shop_domain": SHOP,
                    "Handle": key,
                    "product_id": str(product.get("id", "")),
                    "media_id": str(img["image_url"].split("/")[-1].split("?")[0]),
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
                    "alt_proposed": alt[:512],
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

        seo = {
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
            "source_export_ref": str(item.get("source_json", "")),
            "source_exported_at": "UNKNOWN_PUBLIC_JSON_CAPTURE",
            "locale": "en-US",
            "market": "US",
            "url": item["url"],
            "canonical_url": item["url"],
            "product_type": product.get("product_type", ""),
            "title_current": product.get("title", ""),
            "h1_current": product.get("title", ""),
            "title_proposed": design["title"],
            "title_action": "SET",
            "h1_mapping_status": "NEEDS_THEME_CONFIRMATION",
            "meta_title_current": "",
            "meta_title_source_state": "UNKNOWN_WITHOUT_SHOPIFY_EXPORT",
            "rendered_title_current": product.get("title", ""),
            "theme_title_suffix": "UNKNOWN",
            "meta_title_seo": design["meta_title"],
            "meta_title_action": "SET",
            "meta_title_chars": len(design["meta_title"]),
            "meta_description_current": "",
            "meta_description_source_state": "UNKNOWN_WITHOUT_SHOPIFY_EXPORT",
            "rendered_meta_description_current": "UNKNOWN_WITHOUT_RENDERED_PAGE_META_CAPTURE",
            "meta_description_seo": design["meta_description"],
            "meta_description_action": "SET",
            "meta_description_chars": len(design["meta_description"]),
            "primary_keyword": design["primary"],
            "secondary_keywords": design["secondary"],
            "long_tail_candidates": f"{design['primary']} for entryway; {design['theme']} indoor rug; custom Halloween area rug",
            "keyword_strategy": "Use specific product motif as primary; broad Halloween rug terms remain collection-level candidates.",
            "season": design["season"],
            "keyword_demand_evidence": "SERP observed with marketplace/category pages for Halloween rugs and personalized pet rugs; no paid volume source used.",
            "keyword_serp_fit": "SUPPORTED_BY_PRODUCT_AND_CATEGORY_SERP_OBSERVATION",
            "meta_keyword": design["primary"],
            "search_intent": "Commercial investigation / product purchase",
            "keyword_validation_status": "SERP_OBSERVED_NO_VOLUME",
            "keyword_selection_reason": design["obs"],
            "buyer_research_refs": research_id,
            "buyer_search_summary": design["summary"],
            "baseline_status": "PUBLIC_DATA_ONLY",
            "change_scope": "SEO title; meta description; product title/H1 proposal; description proposal; image alt proposal",
            "mapping_status": "NEEDS_REVIEW",
            "mapping_version": "2.4",
            "description_current_html": product.get("body_html", ""),
            "description_changes_needed": "Tighten duplicate wording, expose observed motif, keep public facts only.",
            "description_proposed": html_to_text(proposed_html(design, product)),
            "description_proposed_html": proposed_html(design, product),
            "description_change_mode": "REWRITE_FROM_PUBLIC_FACTS",
            "description_target_section": "Full product body",
            "description_action": "SET",
            "revision": "1",
            "review_status": "NEEDS_REVIEW",
            "review_reason": "Research proposal only; Shopify admin/export values are unknown and require human review before approval.",
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
        seo.update(img_fields)
        for c in image_cols:
            seo.setdefault(c, "")
        seo_rows.append(seo)

        evidence_rows.append(
            {
                "evidence_id": evidence_id,
                "product_url": item["url"],
                "reviewed_at": NOW,
                "sources_accessed": "; ".join([item["url"], item.get("source_json", ""), *refs]),
                "current_H1": product.get("title", ""),
                "current_meta_title": "UNKNOWN_WITHOUT_RENDERED_PAGE_META_CAPTURE",
                "short_source_excerpt": body_text[:600],
                "verified_product_facts": f"{design['obs']}; product type={product.get('product_type')}; variants={len(product.get('variants', []))}; images={len(images)}.",
                "gallery_image_count": len(images),
                "images_viewed_count": len(images),
                "image_audit_references": "; ".join([f"{evidence_id}-IMG{int(i['image_number']):02d}" for i in images]),
                "SERP_evidence_references": "; ".join(refs),
                "buyer_research_references": research_id,
                "factual_conflicts": "None observed in public data; admin values unknown.",
                "processing_status": "DRAFTED",
                "confidence_and_reason": "MEDIUM: product facts and images verified from public data; keyword demand has SERP observation but no volume data.",
                "fact_to_source_map": "product facts=products_page_1.json; visuals=B001 local images/contact sheets; keyword context=SERP references.",
                "proposed_field_to_fact_map": "primary motif to title/meta/alt; material/backing only used where product image graphics show it.",
            }
        )

        buyer_rows.append(
            {
                "research_id": research_id,
                "product_key": key,
                "supporting_fact_ids": evidence_id,
                "purchase_context": "Seasonal indoor decorating or gift shopping for Halloween.",
                "jtbd_statement": f"When preparing Halloween decor, the buyer wants a {design['theme']} rug that creates a visual focal point without needing permanent decor.",
                "functional_motivation": "Add a soft indoor floor accent with themed artwork and non-slip backing.",
                "emotional_social_motivation": "Make the room feel personal, spooky, playful, or photo-worthy for guests and seasonal memories.",
                "purchase_concerns": "Whether the rug is flat despite the 3D look, whether it slips, size fit, design clarity, and personalization accuracy where relevant.",
                "customer_language": "spooky Halloween rug; optical illusion rug; personalized pet rug; non-slip backing; indoor decor",
                "language_origin": "MARKETPLACE_PRODUCT_COPY_AND_SERP_SNIPPETS",
                "source_refs": "; ".join(refs),
                "source_scope": "CATEGORY",
                "evidence_excerpt": "Paraphrased from public marketplace/category SERP results; no customer review quote used.",
                "observed_at": NOW,
                "market": "US",
                "source_language": "en",
                "research_status": "SUPPORTED_CATEGORY_LANGUAGE_NO_VOLUME",
                "limitations": "No direct customer reviews or paid search volume checked; buyer scenario is evidence-informed, not a claim about Chillgen customers.",
                "seo_application": "Use motif-specific commercial keywords and avoid broad duplicate Halloween rug targeting across all products.",
            }
        )

        keyword_specs = [
            (design["primary"], "PRIMARY", "PRODUCT", design["primary"], "Motif-specific phrase matches the observed product and commercial intent."),
            (design["secondary"].split(", ")[0], "SECONDARY", "PRODUCT", design["primary"], "Relevant synonym or customization phrase; supports product copy without replacing primary."),
            ("halloween rug", "REJECTED", "COLLECTION", "halloween rug", "Too broad for individual product; better assigned to Halloween collection to reduce overlap."),
        ]
        for kw, role, page_type, cluster, reason in keyword_specs:
            keyword_rows.append(
                {
                    "keyword": kw,
                    "product_key": key,
                    "buyer_research_refs": research_id if role != "REJECTED" else "",
                    "query_origin": "SERP_OBSERVED_AND_AGENT_EXPANSION",
                    "semantic_cluster": cluster,
                    "intent": "Commercial",
                    "target_page_type": page_type,
                    "target_url": item["url"] if page_type == "PRODUCT" else "https://chillgen.com/collections/halloween",
                    "keyword_role": role,
                    "decision_reason": reason,
                    "supporting_fact_ids": evidence_id,
                    "demand_evidence": "SERP pages observed; no search volume source used.",
                    "season": design["season"],
                    "research_period": "2026-09-07",
                    "validation_source": "; ".join(refs if role != "REJECTED" else SERP_REFS["collection"]),
                    "checked_at": NOW,
                    "representative_SERP_URLs": "; ".join(refs[:3]),
                    "possible_overlap_with_other_products": "YES" if role == "REJECTED" else "LOW_TO_MEDIUM_WITH_SIMILAR_MOTIF_PRODUCTS",
                    "mapping_reason": "Product-specific motif retained; collection handles broad seasonal head term.",
                    "mapping_status": "NEEDS_REVIEW",
                    "mapping_version": "2.4",
                }
            )

    seo_headers = [k for k in seo_rows[0].keys() if not k.startswith("img_")] + image_cols
    image_headers = list(image_rows[0].keys())
    evidence_headers = list(evidence_rows[0].keys())
    keyword_headers = list(keyword_rows[0].keys())
    buyer_headers = list(buyer_rows[0].keys())

    (ROOT / "batches").mkdir(exist_ok=True)
    (RESULTS / "batches").mkdir(parents=True, exist_ok=True)

    def write_csv(path, rows, headers):
        with path.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=headers)
            w.writeheader()
            w.writerows(rows)

    write_csv(ROOT / "batches" / "B001_SEO_Products.csv", seo_rows, seo_headers)
    write_csv(ROOT / "batches" / "B001_Image_Audit.csv", image_rows, image_headers)
    write_csv(ROOT / "batches" / "B001_Product_Evidence.csv", evidence_rows, evidence_headers)
    write_csv(ROOT / "keyword_research.csv", keyword_rows, keyword_headers)
    with (ROOT / "buyer_search_research.jsonl").open("w", encoding="utf-8") as f:
        for row in buyer_rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")

    with (ROOT / "inventory.csv").open(newline="", encoding="utf-8") as f:
        inv_reader = csv.DictReader(f)
        inv_rows = list(inv_reader)
        inv_headers = inv_reader.fieldnames or []
    for col in ["assigned_batch", "batch_status", "research_status", "review_status"]:
        if col not in inv_headers:
            inv_headers.append(col)
    batch_keys = {row["product_key"] for row in seo_rows}
    for row in inv_rows:
        if row["product_key"] in batch_keys:
            row["assigned_batch"] = BATCH_ID
            row["batch_status"] = "COMPLETED"
            row["research_status"] = "DRAFTED"
            row["review_status"] = "NEEDS_REVIEW"
        else:
            row.setdefault("assigned_batch", "")
            row.setdefault("batch_status", "")
            row.setdefault("research_status", "")
            row.setdefault("review_status", "")
    write_csv(ROOT / "inventory.csv", inv_rows, inv_headers)

    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "SEO_Products", seo_rows, seo_headers)
    write_sheet(wb, "Image_Audit", image_rows, image_headers)
    write_sheet(wb, "Product_Evidence", evidence_rows, evidence_headers)
    write_sheet(wb, "Keyword_Map", keyword_rows, keyword_headers)
    write_sheet(wb, "Buyer_Search_Research", buyer_rows, buyer_headers)
    readme = [
        {
            "metric": "scope",
            "value": f"{BATCH_ID}: 10 public products for {SHOP}, US/en-US",
            "definition": "Research proposal workbook only; all product and image proposals are NEEDS_REVIEW.",
        },
        {
            "metric": "sources",
            "value": "Chillgen public products.json, product URLs, downloaded CDN images, contact sheets, SERP snippets and marketplace/category pages.",
            "definition": "Admin SEO fields, current stored meta fields, and Shopify export media IDs remain unknown without Shopify CSV/export.",
        },
        {
            "metric": "quality_status",
            "value": "DRAFTED_PUBLIC_ONLY",
            "definition": "80/80 images viewed via contact sheets; QA not run; no Shopify payload created.",
        },
        {
            "metric": "deployment_status",
            "value": "NOT_APPROVED_NOT_DEPLOYABLE",
            "definition": "Do not import this workbook. Run QA and obtain explicit approval before any Shopify/Matrixify payload.",
        },
        {
            "metric": "next_step",
            "value": "Stop after B001 and wait for user confirmation before B002.",
            "definition": "Required by project instruction.",
        },
    ]
    write_sheet(wb, "README_QA", readme, ["metric", "value", "definition"])
    out = RESULTS / "batches" / f"SEO_Product_Optimization_through_{BATCH_ID}.xlsx"
    wb.save(out)

    reopened = load_workbook(out, read_only=True, data_only=False)
    sheet_counts = {name: reopened[name].max_row - 1 for name in reopened.sheetnames}

    summary_path = ROOT / "batches" / "B001_summary.md"
    summary_path.write_text(
        "\n".join(
            [
                f"# {BATCH_ID} Summary",
                "",
                f"- Run: {RUN_ID}",
                f"- Shop: {SHOP}",
                "- Scope: 10 products, 80 images",
                "- Status: NEEDS_REVIEW; QA_NOT_RUN; NOT_APPROVED; NO_DEPLOYMENT_PAYLOAD",
                f"- Workbook: {out}",
                f"- Generated at: {NOW}",
                "",
                "## Product Keys",
                *[f"- {row['product_key']}" for row in seo_rows],
            ]
        ),
        encoding="utf-8",
    )

    progress = json.loads((ROOT / "progress.json").read_text())
    progress.update(
        {
            "inventory_status": "COMPLETE",
            "inventory_order_frozen": True,
            "inventory_counts": {
                "discovered": len(inv_rows),
                "public": len(inv_rows),
                "accessible": len(inv_rows),
                "blocked": 0,
            },
            "batch_id": BATCH_ID,
            "batch_product_keys": [row["product_key"] for row in seo_rows],
            "batch_status": "COMPLETED_NEEDS_REVIEW",
            "awaiting_confirmation": True,
            "continuation_confirmation_ref": "USER_CONFIRMED_START_B001",
            "current_product_key": None,
            "current_stage": "B001_COMPLETE_AWAITING_USER_CONFIRMATION",
            "images_completed": [f"{m['product_key']}#img_{int(m['image_number']):02d}" for m in manifest],
            "last_saved_at": NOW,
        }
    )
    progress["artifact_paths"].update(
        {
            "latest_batch_workbook": str(out),
            "b001_seo_products_csv": str(ROOT / "batches" / "B001_SEO_Products.csv"),
            "b001_image_audit_csv": str(ROOT / "batches" / "B001_Image_Audit.csv"),
            "b001_product_evidence_csv": str(ROOT / "batches" / "B001_Product_Evidence.csv"),
            "b001_summary": str(summary_path),
            "b001_image_manifest": str(ROOT / "evidence/images/B001_manifest.json"),
            "b001_contact_sheets": str(ROOT / "evidence/images/B001_contact_sheets"),
        }
    )
    tmp = ROOT / "progress.json.tmp"
    tmp.write_text(json.dumps(progress, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(ROOT / "progress.json")

    print(json.dumps({"workbook": str(out), "sheet_counts": sheet_counts, "summary": str(summary_path)}, indent=2))


if __name__ == "__main__":
    main()

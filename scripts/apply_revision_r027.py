from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R026/SEO_Product_Optimization_revision_R026.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R027"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R027.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R027"
MANIFEST = RUN_DIR / "revision_R027_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R027_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-album-cover-rug-abb1eeca32": {
        "title": "Pink Album Cover Music Rug with Player Controls",
        "meta": "Personalize a pink album cover music rug with custom image, song title, artist text, progress bar and player controls.",
        "description": "Personalize a pink album cover music rug with custom image, song title, artist text, progress bar and player controls.\n\nDesign details\n- Pink rectangular album-player rug artwork shows a custom image placeholder, song and artist text, playback buttons and progress bar.\n- Product images include bedroom and living room mockups, close-up design panels and size reference graphics.\n- SEO copy stays tied to visible album-player artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "custom-road-map-welcome-mat-708eca28bb": {
        "title": "Daniel City Street Road Map Rug with Car Artwork",
        "meta": "Personalize a Daniel city street road map rug with roads, roundabout, hospital, gas station, cars and traffic lights.",
        "description": "Personalize a Daniel city street road map rug with roads, roundabout, hospital, gas station, cars and traffic lights.\n\nDesign details\n- Green rectangular city street rug artwork shows Daniel's City, roadways, roundabout, gas station, hospital, houses and vehicles.\n- Visible details include bridge, traffic lights, room mockups, toy car scenes and size reference graphics.\n- SEO copy focuses on visible city road-map artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "custom-road-map-area-rug-play-mat-9aed6a6ee9-9aed6a6ee9": {
        "title": "Adventure Town Road Map Rug with Pirate Pond Art",
        "meta": "Personalize an adventure town road map rug with custom town name, pirate ship pond, winding roads, homes and campsite.",
        "description": "Personalize an adventure town road map rug with custom town name, pirate ship pond, winding roads, homes and campsite.\n\nDesign details\n- Green road-map rug artwork shows Liam's Town, pirate ship pond, winding roads, houses, hospital, campsite and forest details.\n- Product images include classroom and playroom mockups, toy car scenes, close-up panels and size reference graphics.\n- SEO copy stays tied to visible adventure town artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "custom-road-map-area-rug-personalized-city-street-car-rug-6313a0fee7": {
        "title": "Robert Construction City Road Map Rug with Trucks",
        "meta": "Personalize a Robert construction city road map rug with trucks, excavators, cranes, road signs and roundabout.",
        "description": "Personalize a Robert construction city road map rug with trucks, excavators, cranes, road signs and roundabout.\n\nDesign details\n- Gray and cream road-map rug artwork shows Robert name, construction zone text, dump trucks, excavators, crane and cones.\n- Visible details include road signs, roundabout, houses, classroom and playroom mockups plus size reference graphics.\n- SEO copy focuses on visible construction city artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-christmas-doormat-custom-family-name-7358e56962-7358e56962": {
        "title": "Santa Stop Here Family Doormat with Boot Artwork",
        "meta": "Personalize a Santa Stop Here family doormat with custom family name, Santa boots, green polka dots and red banner.",
        "description": "Personalize a Santa Stop Here family doormat with custom family name, Santa boots, green polka dots and red banner.\n\nDesign details\n- Green Christmas doormat artwork shows white polka dots, Santa boots, red pants, Santa Stop Here text and family name banner.\n- Product images include entryway mockups, person holding mat, close-up design views and size reference graphics.\n- SEO copy stays tied to visible Santa boot artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-christmas-doormat-80a01c19f6-80a01c19f6": {
        "title": "Pet Stocking Christmas Doormat with Dog Portraits",
        "meta": "Personalize a pet stocking Christmas doormat with dog portraits, pet names, string lights, snowflakes and family name.",
        "description": "Personalize a pet stocking Christmas doormat with dog portraits, pet names, string lights, snowflakes and family name.\n\nDesign details\n- Black Christmas doormat artwork shows four dogs sitting in red stockings with pet names, Merry Christmas text and lights.\n- Visible details include custom family name, snowflakes, entryway mockups, person holding mat and close-up design views.\n- SEO copy focuses on visible pet stocking artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-christmas-doormat-family-name-5fb80e7372-5fb80e7372": {
        "title": "Red Monogram Christmas Doormat with Family Name",
        "meta": "Personalize a red monogram Christmas doormat with custom family name, Happy Holidays script, ribbon stripe and snowflakes.",
        "description": "Personalize a red monogram Christmas doormat with custom family name, Happy Holidays script, ribbon stripe and snowflakes.\n\nDesign details\n- Bright red Christmas doormat artwork shows white ribbon stripe, circular monogram initial, Happy Holidays script and snowflakes.\n- Product images include entryway mockups, person holding mat, close-up design views and size reference graphics.\n- SEO copy stays tied to visible monogram holiday artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-christmas-doormat-with-custom-family-name-83c3ea451f": {
        "title": "Elf Stocking Family Christmas Doormat with Pet Options",
        "meta": "Personalize an elf stocking family Christmas doormat with custom family name, individual names, cat and dog options.",
        "description": "Personalize an elf stocking family Christmas doormat with custom family name, individual names, cat and dog options.\n\nDesign details\n- Tan and green Christmas doormat artwork shows elf stocking characters hanging from a mantle with names and family banner.\n- Visible details include string lights, Choose Cat and Choose Dog option charts, entryway mockups and door placement views.\n- SEO copy focuses on visible elf stocking artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-christmas-doormat-with-custom-family-name-d66e920e14": {
        "title": "3D Snowman Christmas Doormat with Quilted Look",
        "meta": "Personalize a 3D snowman Christmas doormat with custom text area, red green quilted look and snowflake accents.",
        "description": "Personalize a 3D snowman Christmas doormat with custom text area, red green quilted look and snowflake accents.\n\nDesign details\n- Red, green and white Christmas doormat artwork shows a smiling snowman in Santa hat and scarf with snowflake squares.\n- Product images include color variants, doorway mockups, person holding mat, close-up design views and size reference graphics.\n- SEO copy stays tied to visible snowman artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-christmas-welcome-doormat-28640346cc": {
        "title": "Plaid Pet Family Christmas Welcome Mat with Names",
        "meta": "Personalize a plaid pet family Christmas welcome mat with dog and cat options, family name, pet names and Santa hats.",
        "description": "Personalize a plaid pet family Christmas welcome mat with dog and cat options, family name, pet names and Santa hats.\n\nDesign details\n- Red plaid Christmas welcome mat artwork shows dogs wearing Santa hats, custom family name, individual pet names and white name strip.\n- Visible details include Choose Cat and Choose Dog option charts, snowy red background, entryway and door mockups.\n- SEO copy focuses on visible pet family holiday artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "anti slip", "non-slip", "non slip",
    "non-skid", "non skid", "nonslip", "kid friendly", "pet friendly", "safe durable",
    "machine-washable", "machine washable", "washable", "quick-dry", "quick dry",
    "memory foam", "microfiber", "velvet", "stain", "fade resistant", "easy clean",
    "easy-clean", "waterproof", "absorbent", "absorption", "backing", "rubber",
    "layered construction", "hd printing", "ultra-soft", "soft ", "cushion",
    "support classroom", "support emotional", "teach", "learning", "build a toy",
]


def headers(ws):
    return {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}


def validate_text(handle: str, item: dict[str, str]) -> None:
    title, meta = item["title"], item["meta"]
    blob = " ".join([title, meta, item["description"]]).lower()
    if not 45 <= len(title) <= 70:
        raise RuntimeError(f"{handle} title length outside 45-70: {len(title)}")
    if len(meta) > 320:
        raise RuntimeError(f"{handle} meta too long: {len(meta)}")
    hits = [term for term in FORBIDDEN if term in blob]
    if hits:
        raise RuntimeError(f"{handle} forbidden terms: {hits}")


def main():
    for handle, item in REVISIONS.items():
        validate_text(handle, item)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(SOURCE)
    ws = wb["SEO_Products"]
    cols = headers(ws)
    required = [
        "Handle", "title_proposed", "meta_title_seo", "meta_title_chars",
        "meta_description_seo", "meta_description_chars", "description_proposed",
        "description_proposed_html", "revision", "review_status", "review_reason",
        "processing_status", "content_qa_status", "issues",
    ]
    missing = [c for c in required if c not in cols]
    if missing:
        raise RuntimeError(f"Missing columns: {missing}")

    changed = []
    for row_idx in range(2, ws.max_row + 1):
        handle = ws.cell(row_idx, cols["Handle"]).value
        if handle not in REVISIONS:
            continue
        item = REVISIONS[handle]
        title, meta, desc = item["title"], item["meta"], item["description"]
        old_revision = str(ws.cell(row_idx, cols["revision"]).value or "")
        ws.cell(row_idx, cols["title_proposed"], title)
        ws.cell(row_idx, cols["meta_title_seo"], title)
        ws.cell(row_idx, cols["meta_title_chars"], len(title))
        ws.cell(row_idx, cols["meta_description_seo"], meta)
        ws.cell(row_idx, cols["meta_description_chars"], len(meta))
        ws.cell(row_idx, cols["description_proposed"], desc)
        ws.cell(row_idx, cols["description_proposed_html"], "<p>" + desc.replace("\n\n", "</p><p>").replace("\n", "<br>") + "</p>")
        ws.cell(row_idx, cols["revision"], "2")
        ws.cell(row_idx, cols["review_status"], "NEEDS_REVIEW")
        ws.cell(row_idx, cols["review_reason"], "Revision R027 fixes QA MAJOR issue: unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R027 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        missing = sorted(set(REVISIONS) - {c[0] for c in changed})
        raise RuntimeError(f"Did not update all handles. Missing: {missing}")

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append([
            "revision_batch_id", "generated_at", "handle", "source_row",
            "old_revision", "new_revision", "changed_fields", "qa_issue_fields",
            "recheck_condition",
        ])
    else:
        log = wb["Revision_Log"]
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R027"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, _title, _meta in changed:
        log.append([
            "R027", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R027",
        "generated_at": NOW,
        "source_workbook": str(SOURCE),
        "output_workbook": str(OUTPUT),
        "source_revision_plan": str(PLAN),
        "product_count": len(changed),
        "handles": [c[0] for c in changed],
        "status": "COMPLETE_AWAITING_REQA",
        "review_status": "NEEDS_REVIEW",
        "content_qa_status": "NOT_RUN",
        "not_approved_not_deployed": True,
        "next_step": "Bắt đầu revision R028 or Re-QA revision R027",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R027 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R027.",
        "- Fixes targeted QA issue: unsupported claims.",
        "- Kept `review_status=NEEDS_REVIEW` and `content_qa_status=NOT_RUN`.",
        "- No Shopify deploy, no approval.", "",
        "## Updated products", "",
        "| Handle | New title | Title chars | Meta chars |",
        "|---|---|---:|---:|",
    ]
    for handle, _row_idx, _old, _new, title, meta in changed:
        lines.append(f"| `{handle}` | {title} | {len(title)} | {len(meta)} |")
    SUMMARY.write_text("\n".join(lines) + "\n")

    reopened = load_workbook(OUTPUT, read_only=True, data_only=True)
    out_ws = reopened["SEO_Products"]
    out_cols = headers(out_ws)
    seen = 0
    for row in out_ws.iter_rows(min_row=2, values_only=True):
        handle = row[out_cols["Handle"] - 1]
        if handle not in REVISIONS:
            continue
        seen += 1
        title = row[out_cols["title_proposed"] - 1]
        meta_title = row[out_cols["meta_title_seo"] - 1]
        meta = row[out_cols["meta_description_seo"] - 1]
        desc = row[out_cols["description_proposed"] - 1]
        blob = " ".join([str(title or ""), str(meta or ""), str(desc or "")]).lower()
        if title != meta_title:
            raise RuntimeError(f"{handle} meta title mismatch")
        if not 45 <= len(title) <= 70 or len(meta) > 320:
            raise RuntimeError(f"{handle} title/meta length failed after save")
        if str(row[out_cols["revision"] - 1]) != "2" or row[out_cols["review_status"] - 1] != "NEEDS_REVIEW" or row[out_cols["content_qa_status"] - 1] != "NOT_RUN":
            raise RuntimeError(f"{handle} status failed after save")
        hits = [term for term in FORBIDDEN if term in blob]
        if hits:
            raise RuntimeError(f"{handle} forbidden terms after save: {hits}")

    log = reopened["Revision_Log"]
    counts = {}
    for row in log.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            counts[row[0]] = counts.get(row[0], 0) + 1
    meta_titles, dup_meta_titles = {}, []
    for row in out_ws.iter_rows(min_row=2, values_only=True):
        mt, handle = row[out_cols["meta_title_seo"] - 1], row[out_cols["Handle"] - 1]
        if mt:
            if mt in meta_titles:
                dup_meta_titles.append((mt, meta_titles[mt], handle))
            meta_titles[mt] = handle
    print(json.dumps({
        "output": str(OUTPUT),
        "updated": seen,
        "revision_log_rows": log.max_row - 1,
        "r027_log_rows": counts.get("R027", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

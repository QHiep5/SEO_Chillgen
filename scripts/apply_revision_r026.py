from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R025/SEO_Product_Optimization_revision_R025.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R026"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R026.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R026"
MANIFEST = RUN_DIR / "revision_R026_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R026_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-album-cover-rug-4af618eed5": {
        "title": "Black Album Cover Music Rug with Green Cover Art",
        "meta": "Personalize a black album cover music rug with green cover art area, song title, artist text and player controls.",
        "description": "Personalize a black album cover music rug with green cover art area, song title, artist text and player controls.\n\nDesign details\n- Black rectangular album-player rug artwork shows a green cover image area, song and artist placeholders, playback controls and progress bar.\n- Visible scenes include music studio, living room and fireplace mockups plus close-up panels and size reference graphics.\n- SEO copy stays tied to visible album-player artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-album-cover-welcome-mat-26099b1534": {
        "title": "Teal Album Cover Welcome Mat with Player Controls",
        "meta": "Personalize a teal album cover welcome mat with photo artwork, song title, artist text and player controls.",
        "description": "Personalize a teal album cover welcome mat with photo artwork, song title, artist text and player controls.\n\nDesign details\n- Teal rectangular music-player mat artwork shows a cover photo area, song and artist placeholders, playback buttons and progress bar.\n- Product images include living room and entryway mockups, close-up design panels and size reference graphics.\n- SEO copy focuses on visible custom music artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-album-cover-area-rug-648fd8e2ff": {
        "title": "Orange Album Cover Music Rug with Gray Photo Art",
        "meta": "Personalize an orange album cover music rug with gray photo art area, song title, artist text and player controls.",
        "description": "Personalize an orange album cover music rug with gray photo art area, song title, artist text and player controls.\n\nDesign details\n- Orange rectangular album-cover rug artwork shows a grayscale photo area, song and artist text placeholders and player buttons.\n- Visible scenes include bedroom and living room mockups, close-up panels and size reference graphics.\n- SEO copy stays tied to visible album-cover artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-album-cover-rug-3bfef01c49-3bfef01c49": {
        "title": "Brown Album Cover Music Rug with Custom Song Text",
        "meta": "Personalize a brown album cover music rug with custom song text, artist name, cover image area and player controls.",
        "description": "Personalize a brown album cover music rug with custom song text, artist name, cover image area and player controls.\n\nDesign details\n- Brown rectangular album-player rug artwork shows a cover image area, song and artist placeholders, playback controls and progress bar.\n- Product images include bright living room mockups, fireplace scene, close-up panels and size reference graphics.\n- SEO copy focuses on visible music-player artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-album-cover-area-rug-52d1b2412d-52d1b2412d": {
        "title": "Burgundy Album Cover Music Rug with Player Controls",
        "meta": "Personalize a burgundy album cover music rug with grayscale cover art, song title, artist text and player controls.",
        "description": "Personalize a burgundy album cover music rug with grayscale cover art, song title, artist text and player controls.\n\nDesign details\n- Burgundy rectangular album-cover rug artwork shows a grayscale image area, song and artist placeholders and playback controls.\n- Visible scenes include living room and entryway mockups, close-up design panels and size reference graphics.\n- SEO copy stays tied to visible album artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-album-cover-rug-75f5a15f3e-75f5a15f3e": {
        "title": "Gray Album Cover Music Rug with Bedroom Mockup",
        "meta": "Personalize a gray album cover music rug with cover image, song title, artist text, player controls and bedroom mockup.",
        "description": "Personalize a gray album cover music rug with cover image, song title, artist text, player controls and bedroom mockup.\n\nDesign details\n- Gray rectangular album-player rug artwork shows a cover image area, song and artist placeholders, playback controls and progress bar.\n- Product images include moody bedroom, living room and entryway scenes plus close-up panels and size reference graphics.\n- SEO copy focuses on visible music-player artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-music-album-cover-welcome-mat-6b66b8a24c": {
        "title": "Turntable Album Cover Music Mat with Vinyl Photo",
        "meta": "Personalize a turntable album cover music mat with vinyl photo area, song name, artist text and gradient colors.",
        "description": "Personalize a turntable album cover music mat with vinyl photo area, song name, artist text and gradient colors.\n\nDesign details\n- Rectangular music mat artwork shows a record turntable graphic with custom photo in the vinyl area and player controls below.\n- Visible variants include pink, green, peach and purple colors, room mockups and size reference graphics.\n- SEO copy stays tied to visible turntable artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-music-album-cover-rug-5088343293-5088343293": {
        "title": "Photo Waveform Album Music Rug with Color Variants",
        "meta": "Personalize a photo waveform music rug with custom image, song title, artist name, heart icon and color variants.",
        "description": "Personalize a photo waveform music rug with custom image, song title, artist name, heart icon and color variants.\n\nDesign details\n- Rectangular music rug artwork shows a couple photo area, song and artist placeholders, play button, heart icon and waveform line.\n- Visible variants include wood texture and blue, green and magenta color options plus room mockups and size reference graphics.\n- SEO copy focuses on visible photo-song artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "custom-road-map-area-rug-8262be4ef1-8262be4ef1": {
        "title": "Personalized Alphabet Road Map Rug with Animal Art",
        "meta": "Personalize an alphabet road map rug with custom town name, animal artwork, number blocks, roads and city details.",
        "description": "Personalize an alphabet road map rug with custom town name, animal artwork, number blocks, roads and city details.\n\nDesign details\n- Bright blue rectangular road-map rug artwork shows Liam's Town, alphabet border, numbers, animals, colored shapes and city roads.\n- Product images include room mockups, close-up map panels, toy car scenes and size reference graphics.\n- SEO copy stays tied to visible alphabet road-map artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "custom-road-map-area-rug-town-play-mat-a27d8ff428": {
        "title": "Construction Town Road Map Rug with Truck Artwork",
        "meta": "Personalize a construction town road map rug with custom name, cranes, dump trucks, road signs and autumn trees.",
        "description": "Personalize a construction town road map rug with custom name, cranes, dump trucks, road signs and autumn trees.\n\nDesign details\n- Gray rectangular construction road-map rug artwork shows Lucas name, yellow trucks, excavators, crane, cones and road signs.\n- Visible details include roundabout, houses, autumn trees, room mockups, close-up map panels and size reference graphics.\n- SEO copy focuses on visible construction town artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "anti slip", "non-slip", "non slip",
    "non-skid", "non skid", "nonslip", "kid friendly", "pet friendly", "safe durable",
    "machine-washable", "machine washable", "washable", "quick-dry", "quick dry",
    "memory foam", "microfiber", "velvet", "stain", "fade resistant", "easy clean",
    "easy-clean", "waterproof", "absorbent", "absorption", "backing", "rubber",
    "layered construction", "hd printing", "ultra-soft", "soft ", "cushion",
    "support classroom", "support emotional", "teach", "learning", "build a truck-play zone", "combine abc",
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
        ws.cell(row_idx, cols["review_reason"], "Revision R026 fixes QA MAJOR issue: unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R026 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
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
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R026"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, _title, _meta in changed:
        log.append([
            "R026", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R026",
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
        "next_step": "Bắt đầu revision R027 or Re-QA revision R026",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R026 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R026.",
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
        "r026_log_rows": counts.get("R026", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

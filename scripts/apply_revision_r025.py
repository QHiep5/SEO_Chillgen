from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R024/SEO_Product_Optimization_revision_R024.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R025"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R025.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R025"
MANIFEST = RUN_DIR / "revision_R025_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R025_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-soccer-trophy-world-flags-doormat-4c01fd4ed3-4c01fd4ed3": {
        "title": "Global Hosts Soccer Stadium Rug with Trophy Art",
        "meta": "Decorate with a global hosts soccer stadium rug featuring blue field artwork, trophy center and red green flag rings.",
        "description": "Decorate with a global hosts soccer stadium rug featuring blue field artwork, trophy center and red green flag rings.\n\nDesign details\n- Rectangular blue soccer rug artwork shows an aerial stadium-style field, center trophy crest and circular host color rings.\n- Visible details include Global Hosts 2026 text, room mockups, close-up panels and size reference graphics.\n- SEO copy focuses on visible soccer stadium artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-soccer-tournament-trophy-rug-0bdc87e8ab-0bdc87e8ab": {
        "title": "Statue Trophy Soccer Rug with Flag Stripe Artwork",
        "meta": "Decorate with a statue trophy soccer rug featuring flag stripe border, gold cup, soccer ball and The World Game text.",
        "description": "Decorate with a statue trophy soccer rug featuring flag stripe border, gold cup, soccer ball and The World Game text.\n\nDesign details\n- Rectangular dark soccer rug artwork shows a Statue of Liberty-style figure, gold trophy, soccer ball and bold flag stripe border.\n- Visible details include The World Game text, star accents, room mockups, close-up views and size reference graphics.\n- SEO copy stays tied to visible tournament artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "feelings-wheel-emotions-round-rug-f48bff995a": {
        "title": "Color Zones Feelings Wheel Round Rug for Classroom",
        "meta": "Decorate with a color zones feelings wheel round rug featuring calm, ready, sad, angry, worried and excited labels.",
        "description": "Decorate with a color zones feelings wheel round rug featuring calm, ready, sad, angry, worried and excited labels.\n\nDesign details\n- Round classroom rug artwork shows green, blue, yellow and red zones with labeled feeling words and simple face icons.\n- Product images include classroom scene, child reading mockup, close-up graphics and size reference panel.\n- SEO copy describes visible emotion-wheel artwork without unsupported educational outcome, material, reverse-side or surface-performance claims.",
    },
    "wheel-of-feelings-and-emotions-round-rug-a7bf6ef454": {
        "title": "Emotions Activity Wheel Round Rug with Face Icons",
        "meta": "Decorate with an emotions activity wheel round rug featuring face icons, color segments and classroom activity prompts.",
        "description": "Decorate with an emotions activity wheel round rug featuring face icons, color segments and classroom activity prompts.\n\nDesign details\n- Round classroom rug artwork shows a multi-ring feelings wheel with face icons and activity prompts such as music and reading cues.\n- Visible details include color-coded segments, classroom group mockups, child reading scene and size reference panel.\n- SEO copy focuses on visible activity-wheel artwork without unsupported wellness, material, reverse-side or surface-performance claims.",
    },
    "custom-mathematics-education-rug-classroom-playroom-eee0004975": {
        "title": "Math Dance Move Classroom Rug with Colorful Icons",
        "meta": "Decorate with a Math Dance Move classroom rug featuring rainbow math text, movement icons and math symbol artwork.",
        "description": "Decorate with a Math Dance Move classroom rug featuring rainbow math text, movement icons and math symbol artwork.\n\nDesign details\n- Black rectangular classroom rug artwork shows Math Dance Move text with colorful stick-figure movement icons and math symbols.\n- Product images include classroom and playroom mockups, close-up artwork panels and group learning scenes.\n- SEO copy stays tied to visible math-themed artwork without unsupported learning outcome, material, reverse-side or surface-performance claims.",
    },
    "custom-mathematics-education-rug-3015a3c96c": {
        "title": "Problem Solve Math Classroom Rug with Strategy Words",
        "meta": "Decorate with a problem solve math classroom rug featuring colorful strategy words, arrows, dots and math icons.",
        "description": "Decorate with a problem solve math classroom rug featuring colorful strategy words, arrows, dots and math icons.\n\nDesign details\n- Black rectangular math rug artwork shows colorful strategy words including problem solve, analyze data and find patterns.\n- Visible details include arrow graphics, dot accents, classroom mockups and children shown in group scenes.\n- SEO copy describes visible math strategy artwork without unsupported lesson-result, material, reverse-side or surface-performance claims.",
    },
    "feelings-wheel-emotions-round-rug-fca05b7c0f": {
        "title": "Detailed Feelings Wheel Round Rug with Emotion Labels",
        "meta": "Decorate with a detailed feelings wheel round rug featuring many emotion labels, color sections and classroom scenes.",
        "description": "Decorate with a detailed feelings wheel round rug featuring many emotion labels, color sections and classroom scenes.\n\nDesign details\n- Round feelings wheel rug artwork shows many emotion labels arranged by color categories such as sad, anger, fear and joy.\n- Product images include classroom group mockups, child reading scenes, close-up panels and size reference graphics.\n- SEO copy focuses on visible feelings-wheel artwork without unsupported educational outcome, material, reverse-side or surface-performance claims.",
    },
    "custom-composition-notebook-classroom-kids-rug-3bd742c8b9-3bd742c8b9": {
        "title": "Personalized Composition Notebook Classroom Rug",
        "meta": "Personalize a composition notebook classroom rug with educator name artwork, rainbow accents, pencils and smiley flowers.",
        "description": "Personalize a composition notebook classroom rug with educator name artwork, rainbow accents, pencils and smiley flowers.\n\nDesign details\n- Rectangular composition notebook rug artwork shows a black-and-white notebook pattern with a custom educator name label.\n- Visible details include rainbow, pencil, hearts, smiley flowers, classroom mockups and close-up design panels.\n- SEO copy stays tied to visible personalized notebook artwork without unsupported material, reverse-side, classroom outcome or surface-performance claims.",
    },
    "personalized-album-cover-area-rug-0587b91810-0587b91810": {
        "title": "Beige Album Cover Music Rug with Player Controls",
        "meta": "Personalize a beige album cover music rug with custom cover image, song title, artist text and player controls.",
        "description": "Personalize a beige album cover music rug with custom cover image, song title, artist text and player controls.\n\nDesign details\n- Beige rectangular music player rug artwork shows an album cover image area, song and artist placeholders and white playback controls.\n- Product images include bedroom and living room mockups, close-up panels and size reference graphics.\n- SEO copy focuses on visible custom album artwork without unsupported material, padding, cleaning, reverse-side or surface-performance claims.",
    },
    "custom-album-cover-rug-music-decor-f2f905d19b-f2f905d19b": {
        "title": "Black Album Cover Music Rug with Studio Mockup",
        "meta": "Personalize a black album cover music rug with custom cover art, song title, artist text and player controls.",
        "description": "Personalize a black album cover music rug with custom cover art, song title, artist text and player controls.\n\nDesign details\n- Black rectangular music player rug artwork shows a custom cover image area, song and artist placeholders and white playback controls.\n- Visible scenes include music studio setting with guitars, speakers and records, plus room mockups and size reference graphics.\n- SEO copy stays tied to visible album-player artwork without unsupported material, padding, cleaning, reverse-side or surface-performance claims.",
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "anti slip", "non-slip", "non slip",
    "non-skid", "non skid", "nonslip", "kid friendly", "pet friendly", "safe durable",
    "machine-washable", "machine washable", "washable ", "quick-dry", "quick dry",
    "memory foam", "microfiber", "velvet", "stain", "fade resistant", "easy clean",
    "easy-clean", "waterproof", "absorbent", "absorption", "backing", "rubber",
    "layered construction", "hd printing", "ultra-soft", "soft ", "cushion",
    "superior quality", "hours of fun", "support classroom", "support emotional",
    "create a calm corner", "make math time active", "inspire math practice",
    "learning-space", "learning space", "learning pattern", "teach", "therapeutic",
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
        ws.cell(row_idx, cols["review_reason"], "Revision R025 fixes QA MAJOR issue: unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R025 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
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
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R025"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, _title, _meta in changed:
        log.append([
            "R025", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R025",
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
        "next_step": "Bắt đầu revision R026 or Re-QA revision R025",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R025 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R025.",
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
        "r025_log_rows": counts.get("R025", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R030/SEO_Product_Optimization_revision_R030.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R031"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R031.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R031"
MANIFEST = RUN_DIR / "revision_R031_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R031_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-classroom-doormat-teacher-name-notebook-pattern-42da091fe1-42da091fe1": {
        "title": "Ocean Classroom Welcome Mat with Sea Animal Art",
        "meta": "Personalize an ocean classroom welcome mat with educator name, whale, octopus, jellyfish, coral and wave artwork.",
        "description": "Personalize an ocean classroom welcome mat with educator name, whale, octopus, jellyfish, coral and wave artwork.\n\nDesign details\n- Blue classroom mat artwork shows Welcome To Mrs. Sophia's Classroom with whale, octopus, jellyfish, coral, shells and waves.\n- Product images include doorway scenes, person holding mat, close-up artwork views and size reference graphics.\n- SEO copy stays tied to visible ocean classroom artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-welcome-doormat-notebook-rug-8e1bdf83d4": {
        "title": "Notebook Classroom Welcome Mat with Color Letters",
        "meta": "Personalize a notebook classroom welcome mat with educator name, colorful letters, school doodles and affirmation text.",
        "description": "Personalize a notebook classroom welcome mat with educator name, colorful letters, school doodles and affirmation text.\n\nDesign details\n- White notebook-style mat artwork shows She Is A Teacher text, colorful block letters, school doodles and student affirmation words.\n- Visible scenes include doorway with dog, room mockups, person holding mat and size reference graphics.\n- SEO copy focuses on visible notebook classroom artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-welcome-doormat-teacher-students-7f5a91b5a9": {
        "title": "Speech Pathologist Welcome Mat with Therapy Icons",
        "meta": "Personalize a speech pathologist welcome mat with name, speech bubbles, voice icons, ABC blocks and puzzle piece.",
        "description": "Personalize a speech pathologist welcome mat with name, speech bubbles, voice icons, ABC blocks and puzzle piece.\n\nDesign details\n- Colorful rectangular mat artwork shows Speech Mrs. Sophia Pathologist text with voice, hearing and language icon panels.\n- Visible details include speech bubbles, ABC blocks, puzzle piece, calendar, doorway scenes and size reference graphics.\n- SEO copy stays tied to visible speech-pathologist artwork without unsupported service outcome, material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-welcome-mat-teacher-name-ff4e426bbe": {
        "title": "Notebook Class Welcome Mat with Rainbow Letters",
        "meta": "Personalize a notebook class welcome mat with educator name, rainbow letters, apple, open book, pencil and hearts.",
        "description": "Personalize a notebook class welcome mat with educator name, rainbow letters, apple, open book, pencil and hearts.\n\nDesign details\n- White notebook-lined mat artwork shows Welcome To Our Class Mrs. Elizabeth with rainbow letters and school icons.\n- Visible details include apple, open book, pencil, hearts, stars, clouds, doorway and classroom scenes plus black-background preview.\n- SEO copy focuses on visible class welcome artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-doormat-welcome-class-door-mat-f34045b91d": {
        "title": "Tribal Classroom Welcome Mat with Chalk Letters",
        "meta": "Personalize a tribal classroom welcome mat with educator name, chalk-style letters and colorful geometric border.",
        "description": "Personalize a tribal classroom welcome mat with educator name, chalk-style letters and colorful geometric border.\n\nDesign details\n- Black classroom mat artwork shows Welcome To Mrs. Sophia's Classroom with bold chalk-style lettering.\n- Visible details include colorful tribal geometric border, entryway scenes, dog doorway photo, person holding mat and size reference graphics.\n- SEO copy stays tied to visible tribal classroom artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-doormat-with-custom-teacher-name-675f2e0900": {
        "title": "Happy To See Your Face Mat with Smiley Flowers",
        "meta": "Personalize a Happy To See Your Face mat with educator name, smiley flowers, apples, hearts and pink design.",
        "description": "Personalize a Happy To See Your Face mat with educator name, smiley flowers, apples, hearts and pink design.\n\nDesign details\n- Bright pink classroom mat artwork shows Happy To See Mrs. Sophia Your Face with smiley icons and checker details.\n- Visible details include apples, flowers, hearts, doorway scenes, person holding mat and size reference graphics.\n- SEO copy focuses on visible pink classroom artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-welcome-doormat-3d9f3ec373": {
        "title": "Everyone Welcome Here Classroom Mat with Handprints",
        "meta": "Personalize an Everyone Welcome Here classroom mat with educator name, rainbow handprints, apple, backpack and stars.",
        "description": "Personalize an Everyone Welcome Here classroom mat with educator name, rainbow handprints, apple, backpack and stars.\n\nDesign details\n- Dark colorful classroom mat artwork shows Everyone Is Welcome Here Mrs. Sophia with rainbow handprints and school icons.\n- Visible details include apple, backpack, calculator, book, star doodles, doorway scenes and size reference graphics.\n- SEO copy stays tied to visible welcome artwork without unsupported classroom outcome, material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-doormat-teacher-notebook-a1ca00cc38-a1ca00cc38": {
        "title": "Pastel Notebook Classroom Mat with Large Pink Bow",
        "meta": "Personalize a pastel notebook classroom mat with educator name, pink bow, rainbow, crayons, apple and book art.",
        "description": "Personalize a pastel notebook classroom mat with educator name, pink bow, rainbow, crayons, apple and book art.\n\nDesign details\n- White notebook-lined classroom mat artwork shows Welcome To Mrs. Smith's Classroom with pastel bubble letters.\n- Visible details include large pink bow, rainbow, apple, book, pencil, crayons, hearts, doorway scenes and classroom mockup.\n- SEO copy focuses on visible pastel notebook artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-doormat-welcome-to-class-door-mat-a74c09137b": {
        "title": "School Doodle Classroom Mat with Rainbow Artwork",
        "meta": "Personalize a school doodle classroom mat with educator name, apple, pencil, backpack, rainbow, books and daisies.",
        "description": "Personalize a school doodle classroom mat with educator name, apple, pencil, backpack, rainbow, books and daisies.\n\nDesign details\n- White classroom mat artwork shows Welcome To Mrs. Sophia's Classroom with colorful school doodles and dotted black border.\n- Visible details include apple, pencil, backpack, rainbow, book stack, paper plane, daisies, doorway scenes and size reference graphics.\n- SEO copy stays tied to visible school doodle artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-doormat-with-custom-teacher-name-ae6fef4505-ae6fef4505": {
        "title": "Sunflower Classroom Welcome Mat with Lion Doodles",
        "meta": "Personalize a sunflower classroom welcome mat with educator name, lion faces, daisies, hearts and yellow artwork.",
        "description": "Personalize a sunflower classroom welcome mat with educator name, lion faces, daisies, hearts and yellow artwork.\n\nDesign details\n- Warm yellow classroom mat artwork shows Welcome To Mrs. Sophia's Classroom with sunflowers and lion face doodles.\n- Visible details include daisies, hearts, lightning doodles, room scenes, dog doorway photo, person holding mat and size reference graphics.\n- SEO copy focuses on visible sunflower classroom artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "anti slip", "non-slip", "non slip",
    "non-skid", "non skid", "nonslip", "kid friendly", "pet friendly", "safe durable",
    "machine-washable", "machine washable", "washable", "quick-dry", "quick dry",
    "memory foam", "microfiber", "velvet", "stain", "fade resistant", "easy clean",
    "easy-clean", "waterproof", "absorbent", "absorption", "backing", "rubber",
    "layered construction", "hd printing", "ultra-soft", "soft ", "cushion",
    "support classroom", "support emotional", "learning-space", "thickened",
    "reinforced", "bound edge", "material panel", "therapy result", "use panels",
    "bathroom", "kitchen", "balcony", "location panel",
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
        ws.cell(row_idx, cols["review_reason"], "Revision R031 fixes QA MAJOR issue: unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R031 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
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
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R031"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, _title, _meta in changed:
        log.append([
            "R031", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R031",
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
        "next_step": "Bắt đầu revision R032 or Re-QA revision R031",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R031 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R031.",
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
        "r031_log_rows": counts.get("R031", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

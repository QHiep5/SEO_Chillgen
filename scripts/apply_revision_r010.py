from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R009/SEO_Product_Optimization_revision_R009.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R010"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R010.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R010"
MANIFEST = RUN_DIR / "revision_R010_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R010_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "custom-tree-of-life-celtic-doormat-e8451c7e99-e8451c7e99": {
        "title": "Emerald Celtic Tree of Life Rug with Green Glow Art",
        "meta": "Decorate with an emerald Celtic Tree of Life rug featuring green glow artwork, knot border, roots and mountain scenery.",
        "description": (
            "Decorate with an emerald Celtic Tree of Life rug featuring green glow artwork, knot border, roots and mountain scenery.\n\n"
            "Design details\n"
            "- Round rug artwork shows a bright green Tree of Life with roots, branches and Celtic knot-inspired border.\n"
            "- Product images include room mockups, close-up texture views, size reference graphics and design detail panels.\n"
            "- SEO copy stays tied to the visible Celtic tree artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-tree-of-life-welcome-mat-celtic-design-7a9bd2508d-7a9bd2508d": {
        "title": "Jeweled Celtic Tree of Life Rug with Knotwork Border",
        "meta": "Style a room with a jeweled Celtic Tree of Life rug featuring ornate knotwork, green stones and landscape artwork.",
        "description": (
            "Style a room with a jeweled Celtic Tree of Life rug featuring ornate knotwork, green stones and landscape artwork.\n\n"
            "Design details\n"
            "- Round Celtic tree design shows a decorative border, jewel-like green accents and a mountain landscape scene.\n"
            "- Gallery images show bedroom, living room and close-up views plus size and design reference panels.\n"
            "- SEO copy focuses on visible Tree of Life artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-tree-of-life-celtic-doormat-eaeb60bb6c-eaeb60bb6c": {
        "title": "Rainbow Mosaic Tree of Life Rug with Colorful Hills",
        "meta": "Add a rainbow mosaic Tree of Life rug with glass-like mosaic colors, sweeping hills, branches and roots artwork.",
        "description": (
            "Add a rainbow mosaic Tree of Life rug with glass-like mosaic colors, sweeping hills, branches and roots artwork.\n\n"
            "Design details\n"
            "- Round rug artwork shows a colorful Tree of Life with rainbow glass-like mosaic segments and flowing hill patterns.\n"
            "- Product images include room mockups, close-up panels, size reference graphics and design detail views.\n"
            "- SEO copy stays tied to the visible tree and mosaic artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-reading-corner-welcome-mat-classroom-library-9e72982546-9e72982546": {
        "title": "Custom Reading Corner Rug with Rainbow Book Art",
        "meta": "Customize a reading corner rug with teacher name text, rainbow open book artwork, butterflies and colorful book quote design.",
        "description": (
            "Customize a reading corner rug with teacher name text, rainbow open book artwork, butterflies and colorful book quote design.\n\n"
            "Design details\n"
            "- Black classroom rug design shows an open book, rainbow colors, butterflies and reading-themed text.\n"
            "- Product images show library corner, classroom and playroom-style scenes plus size reference panels.\n"
            "- SEO copy focuses on visible personalized reading artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-reading-corner-classroom-library-rug-f3d87bd3da": {
        "title": "Personalized Reading Gives You Wings Library Rug",
        "meta": "Personalize a Reading Gives You Wings library rug with name text, butterfly artwork, open book and bright classroom colors.",
        "description": (
            "Personalize a Reading Gives You Wings library rug with name text, butterfly artwork, open book and bright classroom colors.\n\n"
            "Design details\n"
            "- Colorful classroom rug design includes open book artwork, butterflies, flowers and reading-themed quote text.\n"
            "- Gallery images show classroom, reading nook and child group mockups plus size and close-up reference panels.\n"
            "- SEO copy stays tied to the visible personalized library artwork without unsupported material, backing or safety claims."
        ),
    },
    "custom-bookshelf-classroom-library-reading-rug-110e79f122": {
        "title": "Custom Bookshelf Classroom Rug with Reader Quote",
        "meta": "Customize a bookshelf classroom rug with teacher name text, colorful books, school icons and Today a Reader quote.",
        "description": (
            "Customize a bookshelf classroom rug with teacher name text, colorful books, school icons and Today a Reader quote.\n\n"
            "Design details\n"
            "- Bookshelf-style rug artwork shows rows of books, classroom icons and Today a Reader Tomorrow a Leader text.\n"
            "- Product images include classroom mockups, child group scenes, personalization examples and size reference views.\n"
            "- SEO copy focuses on the visible personalized bookshelf artwork without unsupported material, backing or safety claims."
        ),
    },
    "custom-book-rug-classroom-reading-nooks-da82e7625c-da82e7625c": {
        "title": "Personalized Good Day to Read a Book Classroom Rug",
        "meta": "Customize a Good Day to Read a Book classroom rug with name text, pastel flowers, book quote and reading nook scenes.",
        "description": (
            "Customize a Good Day to Read a Book classroom rug with name text, pastel flowers, book quote and reading nook scenes.\n\n"
            "Design details\n"
            "- Light classroom rug artwork shows It's a Good Day to Read a Book text with pastel flowers and book-themed layout.\n"
            "- Product images include classroom scenes, child group mockups, close-up detail panels and size reference graphics.\n"
            "- SEO copy stays tied to the visible personalized reading design without unsupported material, backing or safety claims."
        ),
    },
    "personalized-classroom-rules-teacher-area-rug-c01b2868b7": {
        "title": "Custom Rainbow Classroom Rules Rug with Teacher Name",
        "meta": "Customize a rainbow classroom rules rug with teacher name text, checklist-style reminders, pencils and checkerboard border.",
        "description": (
            "Customize a rainbow classroom rules rug with teacher name text, checklist-style reminders, pencils and checkerboard border.\n\n"
            "Design details\n"
            "- Classroom rules rug design includes rainbow stripes, pencil artwork, checklist panels and black checkerboard edging.\n"
            "- Gallery images show classroom, playroom and child group mockups plus size reference panels.\n"
            "- SEO copy focuses on visible personalized classroom rules artwork without unsupported material, backing or safety claims."
        ),
    },
    "custom-octopus-sea-monster-welcome-mat-273aa005c8": {
        "title": "Personalized Red Blue Octopus Shaped Rug with Tentacles",
        "meta": "Personalize a red blue octopus shaped rug with tentacle artwork, dark sea monster styling and custom name text.",
        "description": (
            "Personalize a red blue octopus shaped rug with tentacle artwork, dark sea monster styling and custom name text.\n\n"
            "Design details\n"
            "- Shaped rug artwork shows a dark octopus form with curled red and blue tentacles on a black background.\n"
            "- Product images include floor mockups, close-up panels, personalization examples and size reference graphics.\n"
            "- SEO copy stays tied to the visible octopus shape and artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-octopus-shaped-area-rug-personalized-sea-monster-tentacle-724ae62b5b": {
        "title": "Personalized Skull Octopus Shaped Rug with Tentacles",
        "meta": "Personalize a skull octopus shaped rug with cream tentacle artwork, skeleton-style sea monster design and name text.",
        "description": (
            "Personalize a skull octopus shaped rug with cream tentacle artwork, skeleton-style sea monster design and name text.\n\n"
            "Design details\n"
            "- Shaped rug artwork shows a cream skull-and-octopus design with curling tentacles on a dark background.\n"
            "- Gallery images show floor mockups, close-up views, personalization examples and size reference panels.\n"
            "- SEO copy focuses on visible skull octopus artwork without unsupported material, backing or surface-performance claims."
        ),
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "non-slip", "non-skid",
    "kid friendly", "pet friendly", "safe durable", "machine-washable",
    "machine washable", "washable ", "quick-dry", "memory foam", "microfiber",
    "velvet", "stain", "fade resistant", "easy clean",
]


def headers(ws):
    return {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}


def validate_text(handle: str, item: dict[str, str]) -> None:
    title = item["title"]
    meta = item["meta"]
    blob = " ".join([item["title"], item["meta"], item["description"]]).lower()
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
        ws.cell(row_idx, cols["review_reason"], "Revision R010 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R010 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        found = {c[0] for c in changed}
        raise RuntimeError(f"Did not update all handles. Missing: {sorted(set(REVISIONS) - found)}")

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    else:
        log = wb["Revision_Log"]
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R010"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R010", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R010",
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
        "next_step": "Bắt đầu revision R011 or Re-QA revision R010",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R010 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R010.",
        "- Fixes targeted QA issues: short T1/T2 and unsupported claims.",
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
        blob = " ".join([
            str(title or ""),
            str(meta or ""),
            str(row[out_cols["description_proposed"] - 1] or ""),
        ]).lower()
        if title != meta_title:
            raise RuntimeError(f"{handle} meta title mismatch")
        if not 45 <= len(title) <= 70:
            raise RuntimeError(f"{handle} title len failed after save")
        if len(meta) > 320:
            raise RuntimeError(f"{handle} meta len failed after save")
        if str(row[out_cols["revision"] - 1]) != "2":
            raise RuntimeError(f"{handle} revision not 2")
        if row[out_cols["review_status"] - 1] != "NEEDS_REVIEW":
            raise RuntimeError(f"{handle} review_status incorrect")
        if row[out_cols["content_qa_status"] - 1] != "NOT_RUN":
            raise RuntimeError(f"{handle} content_qa_status incorrect")
        hits = [term for term in FORBIDDEN if term in blob]
        if hits:
            raise RuntimeError(f"{handle} forbidden terms after save: {hits}")

    log = reopened["Revision_Log"]
    counts = {}
    for row in log.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            counts[row[0]] = counts.get(row[0], 0) + 1
    meta_titles = {}
    dup_meta_titles = []
    for row in out_ws.iter_rows(min_row=2, values_only=True):
        mt = row[out_cols["meta_title_seo"] - 1]
        handle = row[out_cols["Handle"] - 1]
        if mt:
            if mt in meta_titles:
                dup_meta_titles.append((mt, meta_titles[mt], handle))
            meta_titles[mt] = handle

    result = {
        "output": str(OUTPUT),
        "updated": seen,
        "revision_log_rows": log.max_row - 1,
        "r010_log_rows": counts.get("R010", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

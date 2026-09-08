from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R011/SEO_Product_Optimization_revision_R011.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R012"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R012.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R012"
MANIFEST = RUN_DIR / "revision_R012_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R012_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "custom-running-horse-welcome-mat-60ae7781a9": {
        "title": "Monochrome Running Horse Rug with Dust Artwork",
        "meta": "Decorate with a monochrome running horse rug featuring galloping horse artwork, dust effects and neutral room mockups.",
        "description": (
            "Decorate with a monochrome running horse rug featuring galloping horse artwork, dust effects and neutral room mockups.\n\n"
            "Design details\n"
            "- Sepia monochrome rug artwork shows a black horse in motion with dust and sketch-style movement effects.\n"
            "- Product images include bedroom, living room and close-up design views plus size reference graphics.\n"
            "- SEO copy stays tied to visible horse artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "custom-wolf-welcome-doormat-entryway-c00c6179a1-c00c6179a1": {
        "title": "Monochrome Forest Wolf Rug with Pine Tree Artwork",
        "meta": "Decorate with a monochrome forest wolf rug featuring a close-up wolf portrait, pine trees and mountain background.",
        "description": (
            "Decorate with a monochrome forest wolf rug featuring a close-up wolf portrait, pine trees and mountain background.\n\n"
            "Design details\n"
            "- Black, white and gray rug artwork shows a wolf face with misty pine forest and mountain scenery.\n"
            "- Gallery images show the design near a sofa, sectional, desk area and neutral living room scenes.\n"
            "- SEO copy focuses on visible wolf and forest artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "custom-galaxy-wolves-rug-boy-room-e0ded989b1": {
        "title": "Blue Moon Mountain Wolf Rug with Icy Peak Artwork",
        "meta": "Decorate with a blue moon mountain wolf rug featuring a howling wolf, icy peaks, moonbeams and pine trees.",
        "description": (
            "Decorate with a blue moon mountain wolf rug featuring a howling wolf, icy peaks, moonbeams and pine trees.\n\n"
            "Design details\n"
            "- Blue and black rug artwork shows a wolf howling on a snowy mountain ledge below a bright full moon.\n"
            "- Product images include sofa, bedroom, sectional and desk-area mockups with the same moonlit wolf scene.\n"
            "- SEO copy stays tied to visible wolf mountain artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "custom-bear-paw-print-area-rug-a5abc100ab": {
        "title": "Blue Daisy Bear Shaped Rug with Cream Flower Art",
        "meta": "Personalize a blue daisy bear shaped rug with cream flower artwork, hand-drawn fur texture and room mockups.",
        "description": (
            "Personalize a blue daisy bear shaped rug with cream flower artwork, hand-drawn fur texture and room mockups.\n\n"
            "Design details\n"
            "- Bear-shaped rug artwork shows a blue bear silhouette covered with cream daisy flowers and sketch-style fur lines.\n"
            "- Product images include living room, nursery-style and close-up views plus personalization and size reference panels.\n"
            "- SEO copy focuses on visible bear shape and floral artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-bear-paw-print-area-rug-7ff5dd63e9": {
        "title": "Teal Southwestern Paw Rug with Geometric Pattern",
        "meta": "Customize a teal southwestern paw rug with bear-paw shape, cream and navy stripes, yellow center motif and room scenes.",
        "description": (
            "Customize a teal southwestern paw rug with bear-paw shape, cream and navy stripes, yellow center motif and room scenes.\n\n"
            "Design details\n"
            "- Paw-shaped rug artwork shows teal paw pads, cream and navy stripes and a yellow geometric center pattern.\n"
            "- Gallery images show room and nursery-style mockups plus size, personalization and close-up design panels.\n"
            "- SEO copy stays tied to visible paw shape and southwestern pattern without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-bear-paw-print-area-rug-bbb48cf1bd-bbb48cf1bd": {
        "title": "Turquoise Wildlife Paw Rug with Southwestern Art",
        "meta": "Customize a turquoise wildlife paw rug with bear-paw shape, animal pattern, teal pads and red blue stripe details.",
        "description": (
            "Customize a turquoise wildlife paw rug with bear-paw shape, animal pattern, teal pads and red blue stripe details.\n\n"
            "Design details\n"
            "- Paw-shaped rug artwork shows turquoise pads, wildlife figures, cream red blue stripes and southwestern-style patterning.\n"
            "- Product images include room and nursery-style mockups plus size, personalization and close-up design views.\n"
            "- SEO copy focuses on visible paw shape and wildlife artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-composition-notebook-classroom-rug-for-kids-8106743a5b": {
        "title": "Science Notebook Classroom Rug with Teacher Name",
        "meta": "Personalize a science notebook classroom rug with teacher name text, composition pattern, globe, rocket and school icons.",
        "description": (
            "Personalize a science notebook classroom rug with teacher name text, composition pattern, globe, rocket and school icons.\n\n"
            "Design details\n"
            "- Black and white composition book rug artwork includes Mrs. Smith text, globe, rocket, pencil, beaker and planet-style doodles.\n"
            "- Product images show classroom, sofa and child group mockups plus close-up and size reference panels.\n"
            "- SEO copy stays tied to visible science classroom artwork without unsupported material, backing or safety claims."
        ),
    },
    "custom-composition-notebook-classroom-rug-7445b86039-7445b86039": {
        "title": "White Doodle Teacher Name Rug with School Icons",
        "meta": "Customize a white doodle teacher name rug with Mrs. Smith text, pastel letters, school icons and classroom scenes.",
        "description": (
            "Customize a white doodle teacher name rug with Mrs. Smith text, pastel letters, school icons and classroom scenes.\n\n"
            "Design details\n"
            "- White classroom rug artwork shows large colorful Mrs. Smith lettering over a gray school doodle background.\n"
            "- Gallery images include classroom scenes, playroom mockups, child group photos and close-up design panels.\n"
            "- SEO copy focuses on visible personalized teacher-name artwork without unsupported material, backing or safety claims."
        ),
    },
    "custom-tree-of-life-welcome-mat-22c37718e3-22c37718e3": {
        "title": "Monochrome Tree of Life Rug with Line Art Roots",
        "meta": "Decorate with a monochrome Tree of Life rug featuring black and white line art, exposed roots and mountain hills.",
        "description": (
            "Decorate with a monochrome Tree of Life rug featuring black and white line art, exposed roots and mountain hills.\n\n"
            "Design details\n"
            "- Round rug artwork shows a detailed black and white Tree of Life canopy, exposed roots and decorative swirl border.\n"
            "- Product images include room mockups, close-up views and size reference graphics.\n"
            "- SEO copy stays tied to visible Tree of Life artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-celtic-tree-of-life-welcome-mat-1e4bffd67c": {
        "title": "Sepia Celtic Knot Tree Rug with Tree of Life Art",
        "meta": "Customize a sepia Celtic knot tree rug with Tree of Life roots, circular knot border and neutral scroll pattern.",
        "description": (
            "Customize a sepia Celtic knot tree rug with Tree of Life roots, circular knot border and neutral scroll pattern.\n\n"
            "Design details\n"
            "- Round rug artwork shows a sepia Tree of Life with exposed roots, circular Celtic knot border and neutral scroll details.\n"
            "- Gallery images show room mockups, close-up panels and size reference graphics.\n"
            "- SEO copy focuses on visible Celtic tree artwork without unsupported material, backing or surface-performance claims."
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
        ws.cell(row_idx, cols["review_reason"], "Revision R012 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R012 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        found = {c[0] for c in changed}
        raise RuntimeError(f"Did not update all handles. Missing: {sorted(set(REVISIONS) - found)}")

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    else:
        log = wb["Revision_Log"]
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R012"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R012", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R012",
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
        "next_step": "Bắt đầu revision R013 or Re-QA revision R012",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R012 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R012.",
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
        "r012_log_rows": counts.get("R012", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

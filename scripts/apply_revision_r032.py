from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R031/SEO_Product_Optimization_revision_R031.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R032"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R032.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R032"
MANIFEST = RUN_DIR / "revision_R032_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R032_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-classroom-welcome-doormat-notebook-rug-508fa38fe3": {
        "title": "Better With You Classroom Mat with Checkerboard Art",
        "meta": "Personalize a Better With You classroom mat with educator name, pink checkerboard artwork, flowers, globe and school doodles.",
        "description": "Personalize a Better With You classroom mat with educator name, pink checkerboard artwork, flowers, globe and school doodles.\n\nDesign details\n- Pink and black classroom mat artwork shows This Classroom Is Better With You In It with Mrs. Sophia name.\n- Visible details include flowers, globe, school doodles, classroom doorway scene, dog doorway photo and close-up views.\n- SEO copy stays tied to visible classroom welcome artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-doormat-with-custom-teacher-name-956ffe5ee5": {
        "title": "Yay You're Here Teacher Mat with First Grade Badge",
        "meta": "Personalize a Yay You're Here teacher mat with educator name, grade badge, rainbow, books, pencil and school bus.",
        "description": "Personalize a Yay You're Here teacher mat with educator name, grade badge, rainbow, books, pencil and school bus.\n\nDesign details\n- Bright classroom mat artwork shows Yay You're Here with Mrs. Sophia name and 1st grade badge.\n- Visible details include rainbow, books, pencil, school bus, flower doodles, doorway scenes and size reference graphics.\n- SEO copy focuses on visible teacher welcome artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-welcome-doormat-b404579ec3-b404579ec3": {
        "title": "School Supplies Classroom Mat with Music Icons",
        "meta": "Personalize a school supplies classroom mat with educator name, welcome text, pencil border, apple and music icons.",
        "description": "Personalize a school supplies classroom mat with educator name, welcome text, pencil border, apple and music icons.\n\nDesign details\n- Bright rectangular mat artwork shows Welcome To Mrs. Smith's Classroom with colorful pencil border.\n- Visible details include apple, music note, lightning bolt, drum, guitar, art icons, entrance scene and hand-held photo.\n- SEO copy stays tied to visible school-supplies artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-welcome-doormat-77047e9c34-77047e9c34": {
        "title": "Math Teacher Classroom Mat with Chalkboard Numbers",
        "meta": "Personalize a math teacher classroom mat with educator name, chalkboard numbers, equations, ruler, calculator and pencils.",
        "description": "Personalize a math teacher classroom mat with educator name, chalkboard numbers, equations, ruler, calculator and pencils.\n\nDesign details\n- Dark chalkboard-style classroom mat artwork shows Mrs. Sophia Math She Has Problems with large math letters.\n- Visible details include numbers, ruler, calculator, equations, pencils, apple, doorway scenes and size reference graphics.\n- SEO copy focuses on visible math classroom artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-welcome-doormat-2f0308422b-2f0308422b": {
        "title": "Dinosaur Classroom Welcome Mat with Jungle Art",
        "meta": "Personalize a dinosaur classroom welcome mat with educator name, purple orange green dinosaurs, leaves and playful lettering.",
        "description": "Personalize a dinosaur classroom welcome mat with educator name, purple orange green dinosaurs, leaves and playful lettering.\n\nDesign details\n- White and colorful classroom mat artwork shows Welcome To Mrs. Sophia's Classroom with purple, orange and green dinosaurs.\n- Visible details include tropical leaves, playful letters, room and doorway mockups, person holding mat and size reference graphics.\n- SEO copy stays tied to visible dinosaur classroom artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-classroom-welcome-doormat-b5fe59f765-b5fe59f765": {
        "title": "Affirmation Pencil Classroom Mat with Floral Art",
        "meta": "Personalize an affirmation pencil classroom mat with educator name, black floral design and pencil-shaped word labels.",
        "description": "Personalize an affirmation pencil classroom mat with educator name, black floral design and pencil-shaped word labels.\n\nDesign details\n- Black floral classroom mat artwork shows In This Classroom You Are with Mrs. Sophia name and pencil-shaped words.\n- Visible words include kind, valued, important, trusted, creative, loved and smart, with room mockups and size graphics.\n- SEO copy focuses on visible affirmation-pencil artwork without unsupported classroom outcome, material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-soccer-2026-area-rug-73a22990b0-73a22990b0": {
        "title": "USA Road To Glory Soccer Rug with Eagle Wings",
        "meta": "Decorate with a USA Road To Glory soccer rug featuring eagle wings, soccer ball, stars, stripes and World Games text.",
        "description": "Decorate with a USA Road To Glory soccer rug featuring eagle wings, soccer ball, stars, stripes and World Games text.\n\nDesign details\n- Patriotic soccer rug artwork shows USA Road To Glory text, eagle wings, soccer ball, stars and stripe border.\n- Product images include living room and fireside mockups, close-up artwork views and size reference graphics.\n- SEO copy stays tied to visible USA soccer artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-soccer-trophy-world-flags-doormat-00f2f9bbfa": {
        "title": "World Flags Soccer Stadium Rug with Pitch Center",
        "meta": "Decorate with a world flags soccer stadium rug featuring green pitch center, soccer ball, crowd stands and flag accents.",
        "description": "Decorate with a world flags soccer stadium rug featuring green pitch center, soccer ball, crowd stands and flag accents.\n\nDesign details\n- Soccer stadium rug artwork shows a green pitch center, soccer ball, circular crowd stands and many world flag accents.\n- Visible scenes include living room and fireside mockups, close-up design panels and size reference graphics.\n- SEO copy focuses on visible stadium flag artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-soccer-2026-area-rug-c2d07261b5": {
        "title": "USA Canada Mexico Soccer Rug with Stadium Lights",
        "meta": "Decorate with a USA Canada Mexico soccer rug featuring soccer ball center, host flags, stadium lights and green field art.",
        "description": "Decorate with a USA Canada Mexico soccer rug featuring soccer ball center, host flags, stadium lights and green field art.\n\nDesign details\n- Bright soccer rug artwork shows a soccer ball center with USA, Canada and Mexico flag colors and stadium light beams.\n- Product images include living room and fireside mockups, close-up artwork panels and size reference graphics.\n- SEO copy stays tied to visible host-flag artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
    "personalized-soccer-2026-area-rug-a60ad6dcf9-a60ad6dcf9": {
        "title": "North America Soccer Crest Rug with Medal Border",
        "meta": "Decorate with a North America soccer crest rug featuring center ball, USA Canada Mexico flags and ornate medal border.",
        "description": "Decorate with a North America soccer crest rug featuring center ball, USA Canada Mexico flags and ornate medal border.\n\nDesign details\n- Ornate soccer crest rug artwork shows center ball, USA Canada Mexico flag elements, circular compass details and medal-style border.\n- Visible scenes include living room and fireside mockups, close-up design panels and size reference graphics.\n- SEO copy focuses on visible soccer crest artwork without unsupported material, cleaning, reverse-side or surface-performance claims.",
    },
}


FORBIDDEN = [
    "indoor/outdoor", " indoor floor", " outdoor ", "anti-slip", "anti slip",
    "non-slip", "non slip", "non-skid", "non skid", "nonslip", "kid friendly",
    "pet friendly", "safe durable", "machine-washable", "machine washable",
    "washable", "quick-dry", "quick dry", "memory foam", "microfiber", "velvet",
    "stain", "fade resistant", "fade", "easy clean", "easy-clean", "waterproof",
    "absorbent", "absorption", "backing", "rubber", "layered construction",
    "hd printing", "ultra-soft", "soft ", "cushion", "support classroom",
    "support emotional", "learning-space", "thickened", "reinforced", "bound edge",
    "material panel", "bring home",
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
        ws.cell(row_idx, cols["review_reason"], "Revision R032 fixes QA MAJOR issue: unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R032 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
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
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R032"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, _title, _meta in changed:
        log.append([
            "R032", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R032",
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
        "next_step": "Bắt đầu revision R033 or Re-QA revision R032",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R032 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R032.",
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
        "r032_log_rows": counts.get("R032", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

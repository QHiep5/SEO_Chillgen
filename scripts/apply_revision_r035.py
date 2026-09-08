from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R034/SEO_Product_Optimization_revision_R034.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R035"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R035.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R035"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"
NOW = datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-christmas-doormat-with-custom-family-name-315531dc75-315531dc75": {
        "title": "Buffalo Plaid Christmas Family Mat with Monogram",
        "meta": "Personalize a buffalo plaid Christmas family mat with monogram initial, wreath, red bow, snowflakes and family name.",
        "description": "Personalize a buffalo plaid Christmas family mat with monogram initial, wreath, red bow, snowflakes and family name.\n\nDesign details\n- Red and black buffalo plaid Christmas mat artwork shows green wreath, red bow, monogram R and The Robinson Family text.\n- Visible scenes include front door, white doorway, wood floor, bedside mockups, person holding mat and close-up design views.\n- SEO copy stays tied to visible holiday artwork without extra care, construction, reverse-side or surface-performance claims.",
    },
    "feelings-wheel-emotions-round-rug-0f979e072d-0f979e072d": {
        "title": "Black Feelings Flower Round Rug with Daisy Faces",
        "meta": "Decorate with a black feelings flower round rug featuring daisy face center, emotion words and colorful face icons.",
        "description": "Decorate with a black feelings flower round rug featuring daisy face center, emotion words and colorful face icons.\n\nDesign details\n- Round black and white feelings rug artwork shows How Are You Feeling text around a center daisy face.\n- Visible details include outer emotion words, colorful face icons, classroom group scene, child reading and toddler photo mockups.\n- SEO copy stays tied to visible emotion artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
    },
    "wheel-of-feelings-and-emotions-round-rug-697b1d04c9": {
        "title": "Colorful Crayon Feelings Wheel Rug with Emotion Words",
        "meta": "Decorate with a colorful crayon feelings wheel rug featuring You Are center text, crayon ring and emotion words.",
        "description": "Decorate with a colorful crayon feelings wheel rug featuring You Are center text, crayon ring and emotion words.\n\nDesign details\n- Round black classroom rug artwork shows You Are center text with colorful crayons arranged around the circle.\n- Visible border words include happy, loved, stressed, nervous and frustrated, with classroom and child photo mockups.\n- SEO copy focuses on visible crayon feelings artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
    },
    "custom-composition-notebook-classroom-kids-rug-4e7e4d40cf-4e7e4d40cf": {
        "title": "Black Emotion Classroom Rug with I Feel Prompts",
        "meta": "Personalize a black emotion classroom rug with educator name, I Feel and I Can sections, colorful faces and prompt circles.",
        "description": "Personalize a black emotion classroom rug with educator name, I Feel and I Can sections, colorful faces and prompt circles.\n\nDesign details\n- Black classroom rug artwork shows I Feel and I Can sections with colorful emotion faces and circular prompt graphics.\n- Visible details include Mrs. Smith's Classroom text, classroom and playroom mockups, children sitting on rug and size reference graphics.\n- SEO copy stays tied to visible emotion prompt artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
    },
    "custom-composition-notebook-classroom-rug-for-kids-64af3c74f1": {
        "title": "First Grade Writing Classroom Rug with Rainbow Art",
        "meta": "Personalize a first grade writing classroom rug with educator name, rainbow, pencil, stars, apples and green Write badge.",
        "description": "Personalize a first grade writing classroom rug with educator name, rainbow, pencil, stars, apples and green Write badge.\n\nDesign details\n- Black and white doodle classroom rug artwork reads Welcome To Mrs. Brown's 1st Grade Class.\n- Visible details include rainbow, pencil, stars, apples, smiley face, green Write badge, classroom scenes and group photos.\n- SEO copy focuses on visible first-grade writing artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
    },
    "custom-composition-notebook-classroom-rugs-for-kids-88021773b3": {
        "title": "Crayon Affirmation Classroom Rug with You Are Words",
        "meta": "Personalize a crayon affirmation classroom rug with educator name, rainbow crayons, You Are message and hearts.",
        "description": "Personalize a crayon affirmation classroom rug with educator name, rainbow crayons, You Are message and hearts.\n\nDesign details\n- Black classroom rug artwork reads In This Classroom You Are with rainbow crayons and affirmation words on each crayon.\n- Visible details include Mrs. Naomi name, hearts, classroom floor scenes, children sitting on rug and close-up views.\n- SEO copy stays tied to visible crayon affirmation artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
    },
    "custom-composition-notebook-classroom-rug-for-kids-95c57fb741": {
        "title": "Pencil Notebook Classroom Rug with Composition Art",
        "meta": "Personalize a pencil notebook classroom rug with educator name, lined paper, pink spine, ABC text, apple and books.",
        "description": "Personalize a pencil notebook classroom rug with educator name, lined paper, pink spine, ABC text, apple and books.\n\nDesign details\n- Lined composition notebook style rug artwork shows Composition Book label, Mrs. Brown name and pink spine.\n- Visible details include ABC text, pencil, apple, books, crayons, A+ mark, classroom mockups and group photos.\n- SEO copy focuses on visible notebook classroom artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
    },
    "custom-composition-notebook-classroom-rug-for-kids-592c6d1fa7": {
        "title": "Piano Educator Classroom Rug with Keyboard Artwork",
        "meta": "Personalize a piano educator classroom rug with name, colorful keyboard, music notes, stars and pastel letters.",
        "description": "Personalize a piano educator classroom rug with name, colorful keyboard, music notes, stars and pastel letters.\n\nDesign details\n- Black music-themed classroom rug artwork shows colorful piano keyboard, music notes, stars and Mrs. Aisha name.\n- Product images include classroom and playroom mockups, children group photos, close-up design panels and size reference graphics.\n- SEO copy stays tied to visible piano classroom artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
    },
    "celtic-tree-of-life-welcome-mat-5798e850b0": {
        "title": "Sunset Celtic Tree Of Life Rug with Knot Border",
        "meta": "Decorate with a sunset Celtic Tree of Life rug featuring orange glow, green knot border, roots and branch artwork.",
        "description": "Decorate with a sunset Celtic Tree of Life rug featuring orange glow, green knot border, roots and branch artwork.\n\nDesign details\n- Rectangular Celtic rug artwork shows a circular orange sunset center with spreading tree roots and branches.\n- Visible details include green Celtic border, knot and symbol panels, living room mockups, close-up design views and size reference graphics.\n- SEO copy focuses on visible Celtic tree artwork without extra care, construction, reverse-side or surface-performance claims.",
    },
    "custom-tree-of-life-welcome-mat-1cdb86df37-1cdb86df37": {
        "title": "Color Glass Yggdrasil Rug with Mountain Scene",
        "meta": "Decorate with a color glass Yggdrasil rug featuring mountain landscape, Tree of Life roots and jewel-tone border.",
        "description": "Decorate with a color glass Yggdrasil rug featuring mountain landscape, Tree of Life roots and jewel-tone border.\n\nDesign details\n- Tree of Life rug artwork shows a mountain and sunset landscape with exposed roots and branching tree center.\n- Visible details include green, blue, yellow and purple panel styling, circular border, room mockups and size reference graphics.\n- SEO copy stays tied to visible Yggdrasil artwork without extra care, construction, reverse-side or surface-performance claims.",
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "anti slip", "non-slip", "non slip",
    "non-skid", "non skid", "nonslip", "machine-washable", "machine washable",
    "washable", "vacuum", "quick-dry", "quick dry", "memory foam", "microfiber",
    "velvet", "plush", "stain", "fade", "easy clean", "easy-clean", "waterproof",
    "absorbent", "backing", "rubber", "layer construction", "layered construction",
    "ultra-soft", "soft ", "cushion", "thickness", "reinforced", "bound edge",
    "material panel", "learning-pattern", "learning pattern", "support classroom",
    "support emotional", "encourage students", "use a black", "classroom sel",
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
        ws.cell(row_idx, cols["review_reason"], "Revision R035 fixes QA MAJOR issue: unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R035 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        missing = sorted(set(REVISIONS) - {c[0] for c in changed})
        raise RuntimeError(f"Did not update all handles. Missing: {missing}")

    log = wb["Revision_Log"] if "Revision_Log" in wb.sheetnames else wb.create_sheet("Revision_Log")
    if log.max_row == 1 and log.cell(1, 1).value != "revision_batch_id":
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R035"]
    if log.max_row > 1:
        log.delete_rows(2, log.max_row - 1)
    for r in keep:
        log.append(list(r))
    for handle, row_idx, old_revision, new_revision, _title, _meta in changed:
        log.append([
            "R035", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R035", "generated_at": NOW,
        "source_workbook": str(SOURCE), "output_workbook": str(OUTPUT),
        "source_revision_plan": str(PLAN), "product_count": len(changed),
        "handles": [c[0] for c in changed], "status": "COMPLETE_AWAITING_REQA",
        "review_status": "NEEDS_REVIEW", "content_qa_status": "NOT_RUN",
        "not_approved_not_deployed": True,
        "next_step": "Bắt đầu revision R036 or Re-QA revision R035",
    }
    (RUN_DIR / "revision_R035_manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R035 Summary", "", f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`", f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R035.",
        "- Fixes targeted QA issue: unsupported claims.",
        "- Kept `review_status=NEEDS_REVIEW` and `content_qa_status=NOT_RUN`.",
        "- No Shopify deploy, no approval.", "", "## Updated products", "",
        "| Handle | New title | Title chars | Meta chars |", "|---|---|---:|---:|",
    ]
    for handle, _row_idx, _old, _new, title, meta in changed:
        lines.append(f"| `{handle}` | {title} | {len(title)} | {len(meta)} |")
    (OUTPUT_DIR / "REVISION_R035_SUMMARY.md").write_text("\n".join(lines) + "\n")

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
        "output": str(OUTPUT), "updated": seen,
        "revision_log_rows": log.max_row - 1,
        "r035_log_rows": counts.get("R035", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

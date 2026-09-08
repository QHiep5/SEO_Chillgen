from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R033/SEO_Product_Optimization_revision_R033.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R034"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R034.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R034"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"
NOW = datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()


REVISIONS = {
    "custom-running-horse-welcome-mat-b4cfa05f9b-b4cfa05f9b": {
        "title": "Sunset Horse Silhouette Rug with Pine Forest Art",
        "meta": "Decorate with a sunset horse silhouette rug featuring pink sky, pine forest, mountain shapes and room mockups.",
        "description": "Decorate with a sunset horse silhouette rug featuring pink sky, pine forest, mountain shapes and room mockups.\n\nDesign details\n- Rectangular horse rug artwork shows a dark running horse silhouette against a pink and blue sunset sky.\n- Visible details include pine forest, mountain shapes, patio and living room mockups, close-up artwork panels and size reference graphics.\n- SEO copy stays tied to visible horse artwork without extra care, construction, reverse-side or surface-performance claims.",
    },
    "custom-western-running-horse-doormat-c4ecf71d0f-c4ecf71d0f": {
        "title": "White Horse Golden Halo Rug with Starry Sky Art",
        "meta": "Decorate with a white horse golden halo rug featuring starry sky, glowing circle, glitter trail and room mockups.",
        "description": "Decorate with a white horse golden halo rug featuring starry sky, glowing circle, glitter trail and room mockups.\n\nDesign details\n- Dark blue and gold horse rug artwork shows a white running horse inside a glowing circular halo.\n- Visible details include starry sky, golden glitter trail, room mockups, close-up artwork panels and size reference graphics.\n- SEO copy stays tied to visible horse artwork without extra care, construction, reverse-side or surface-performance claims.",
    },
    "custom-galaxy-wolves-rug-boy-room-a0375b4217-a0375b4217": {
        "title": "Tribal Medallion Wolf Rug with Amber Portrait",
        "meta": "Decorate with a tribal medallion wolf rug featuring amber wolf portrait, ornate circular frame and branch details.",
        "description": "Decorate with a tribal medallion wolf rug featuring amber wolf portrait, ornate circular frame and branch details.\n\nDesign details\n- Tan and brown wolf rug artwork shows a large wolf face inside an ornate circular tribal-style medallion.\n- Visible details include root or branch accents, amber highlights, sofa scenes, living room mockups and desk-area mockup.\n- SEO copy stays tied to visible wolf medallion artwork without extra care, construction, reverse-side or surface-performance claims.",
    },
    "wolf-galaxy-carpet-boy-room-decor-rug-4c471ed92f-4c471ed92f": {
        "title": "Purple Dreamcatcher Wolf Rug with Feather Accents",
        "meta": "Decorate with a purple dreamcatcher wolf rug featuring dual wolf portrait, bead details, feather accents and cloudy sky.",
        "description": "Decorate with a purple dreamcatcher wolf rug featuring dual wolf portrait, bead details, feather accents and cloudy sky.\n\nDesign details\n- Purple and gold wolf rug artwork shows two close-up wolf faces with dreamcatcher bead and feather accents.\n- Visible scenes include sofa, wood floor, living room, sectional and desk mockups.\n- SEO copy stays tied to visible wolf dreamcatcher artwork without extra care, construction, reverse-side or surface-performance claims.",
    },
    "custom-bear-paw-print-area-rug-4e14d548d7-4e14d548d7": {
        "title": "Black Floral Bear Shaped Rug with Color Flowers",
        "meta": "Decorate with a black floral bear shaped rug featuring bear silhouette, colorful flowers and custom outline artwork.",
        "description": "Decorate with a black floral bear shaped rug featuring bear silhouette, colorful flowers and custom outline artwork.\n\nDesign details\n- Bear-shaped rug artwork shows a black bear silhouette filled with colorful floral pattern and shaped outline.\n- Product images include living room and nursery mockups, custom shape views, close-up artwork panels and size reference graphics.\n- SEO copy stays tied to visible bear-and-flower artwork without extra care, construction, reverse-side or surface-performance claims.",
    },
    "custom-bear-paw-print-area-rug-3821bbd1dc-3821bbd1dc": {
        "title": "Southwestern Bear Paw Rug with Geometric Pattern",
        "meta": "Decorate with a southwestern bear paw rug featuring brown paw pads, cream blue pattern and yellow diamond center.",
        "description": "Decorate with a southwestern bear paw rug featuring brown paw pads, cream blue pattern and yellow diamond center.\n\nDesign details\n- Paw-shaped rug artwork shows brown paw pads with cream and blue southwestern pattern and yellow diamond center.\n- Visible scenes include living room and nursery mockups, custom shape views, close-up artwork panels and size reference graphics.\n- SEO copy stays tied to visible paw artwork without extra care, construction, reverse-side or surface-performance claims.",
    },
    "custom-road-map-area-rug-play-mat-baa948f510-baa948f510": {
        "title": "Liam Alphabet Town Road Rug with City Streets",
        "meta": "Personalize a Liam alphabet town road rug with city streets, buildings, vehicles, traffic icons and alphabet border.",
        "description": "Personalize a Liam alphabet town road rug with city streets, buildings, vehicles, traffic icons and alphabet border.\n\nDesign details\n- Colorful rectangular road-map rug artwork shows Liam's Town, alphabet border, city streets, buildings and vehicles.\n- Visible details include traffic icons, classroom and playroom mockups, group scene, close-up map panels and size reference graphics.\n- SEO copy stays tied to visible road-map artwork without extra care, construction, reverse-side or surface-performance claims.",
    },
    "custom-abc-educational-classroom-round-rug-33de96c620-33de96c620": {
        "title": "White ABC Flower Classroom Round Rug with Name",
        "meta": "Personalize a white ABC flower classroom round rug with educator name, alphabet border, smiling flower and school icons.",
        "description": "Personalize a white ABC flower classroom round rug with educator name, alphabet border, smiling flower and school icons.\n\nDesign details\n- White round classroom rug artwork shows Mrs. Smith's Classroom with colorful alphabet letters around the edge.\n- Visible details include smiling pink flower center, school icons, group activity scene, child reading and baby photo mockups.\n- SEO copy stays tied to visible ABC flower artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
    },
    "classroom-round-rug-educational-feelings-carpet-6f4ceb729b": {
        "title": "Rainbow Affirmation Classroom Round Rug with Name",
        "meta": "Personalize a rainbow affirmation classroom round rug with educator name, pastel rainbow text and I Am Loved center.",
        "description": "Personalize a rainbow affirmation classroom round rug with educator name, pastel rainbow text and I Am Loved center.\n\nDesign details\n- Cream round classroom rug artwork shows pastel rainbow affirmations and large I Am Loved text with Mrs. Smith name.\n- Visible phrases include I am confident, grateful, enough, smart and kind, with classroom and child photo mockups.\n- SEO copy stays tied to visible affirmation artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
    },
    "custom-abc-educational-welcome-mat-bed98ba4ec-bed98ba4ec": {
        "title": "Black ABC Flower Classroom Round Rug with Name",
        "meta": "Personalize a black ABC flower classroom round rug with educator name, colorful alphabet border, stars and flower center.",
        "description": "Personalize a black ABC flower classroom round rug with educator name, colorful alphabet border, stars and flower center.\n\nDesign details\n- Black round classroom rug artwork shows Mrs. Smith's Classroom text with colorful alphabet letters around the edge.\n- Visible details include smiling pink flower center, stars, school icons, classroom group scene, child reading and baby photo mockups.\n- SEO copy stays tied to visible ABC flower artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "anti slip", "non-slip", "non slip",
    "non-skid", "non skid", "nonslip", "machine-washable", "machine washable",
    "washable", "vacuum", "quick-dry", "quick dry", "memory foam", "microfiber",
    "velvet", "plush", "stain", "fade", "easy clean", "easy-clean", "waterproof",
    "absorbent", "backing", "rubber", "layer construction", "layered construction",
    "ultra-soft", "soft ", "cushion", "reinforced", "bound edge", "material panel",
    "learning border", "learning-space", "calm colors", "teach", "support",
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
        ws.cell(row_idx, cols["review_reason"], "Revision R034 fixes QA MAJOR issue: unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R034 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        missing = sorted(set(REVISIONS) - {c[0] for c in changed})
        raise RuntimeError(f"Did not update all handles. Missing: {missing}")

    log = wb["Revision_Log"] if "Revision_Log" in wb.sheetnames else wb.create_sheet("Revision_Log")
    if log.max_row == 1:
        log.cell(1, 1, "revision_batch_id")
        log.cell(1, 2, "generated_at")
        log.cell(1, 3, "handle")
        log.cell(1, 4, "source_row")
        log.cell(1, 5, "old_revision")
        log.cell(1, 6, "new_revision")
        log.cell(1, 7, "changed_fields")
        log.cell(1, 8, "qa_issue_fields")
        log.cell(1, 9, "recheck_condition")
    keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R034"]
    if log.max_row > 1:
        log.delete_rows(2, log.max_row - 1)
    for r in keep:
        log.append(list(r))
    for handle, row_idx, old_revision, new_revision, _title, _meta in changed:
        log.append([
            "R034", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R034",
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
        "next_step": "Bắt đầu revision R035 or Re-QA revision R034",
    }
    (RUN_DIR / "revision_R034_manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R034 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R034.",
        "- Fixes targeted QA issue: unsupported claims.",
        "- Kept `review_status=NEEDS_REVIEW` and `content_qa_status=NOT_RUN`.",
        "- No Shopify deploy, no approval.", "",
        "## Updated products", "",
        "| Handle | New title | Title chars | Meta chars |",
        "|---|---|---:|---:|",
    ]
    for handle, _row_idx, _old, _new, title, meta in changed:
        lines.append(f"| `{handle}` | {title} | {len(title)} | {len(meta)} |")
    (OUTPUT_DIR / "REVISION_R034_SUMMARY.md").write_text("\n".join(lines) + "\n")

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
        "r034_log_rows": counts.get("R034", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

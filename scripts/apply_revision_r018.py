from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R017/SEO_Product_Optimization_revision_R017.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R018"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R018.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R018"
MANIFEST = RUN_DIR / "revision_R018_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R018_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-teachers-classroom-rug-custom-teacher-name-c-design-09": {
        "title": "Personalized Classroom Jobs Rug with Helper Icons",
        "meta": "Customize a classroom jobs rug with teacher name text, rainbow artwork, checkerboard border and helper icons.",
        "description": (
            "Customize a classroom jobs rug with teacher name text, rainbow artwork, checkerboard border and helper icons.\n\n"
            "Design details\n"
            "- Black-bordered classroom rug artwork shows Classroom Jobs title, rainbow header and teacher-name classroom text.\n"
            "- Visible job icons include line leader, door holder, lights helper, calendar helper, weather reporter and librarian.\n"
            "- SEO copy focuses on visible classroom jobs artwork without unsupported material, cleaning, reverse-side or surface-performance claims."
        ),
    },
    "personalized-teachers-classroom-rug-custom-teacher-name-c-design-07": {
        "title": "Personalized Classroom Rules Rug with Rainbow Art",
        "meta": "Customize a classroom rules rug with teacher name text, rainbow artwork, school icons and listen be kind prompts.",
        "description": (
            "Customize a classroom rules rug with teacher name text, rainbow artwork, school icons and listen be kind prompts.\n\n"
            "Design details\n"
            "- Black doodle-border rug artwork shows Classroom Rules title with rainbow, school icons and teacher-name personalization.\n"
            "- Visible prompt text includes listen, be kind, share, follow directions, helping and helping hands.\n"
            "- SEO copy describes visible classroom rules artwork only, without safety, material, cleaning or surface-performance claims."
        ),
    },
    "personalized-teachers-classroom-rug-custom-teacher-name-c-design-05": {
        "title": "Personalized Voice Level Rug with Classroom Chart",
        "meta": "Customize a voice level classroom rug with teacher name text, 0 to 4 noise level chart and colorful icons.",
        "description": (
            "Customize a voice level classroom rug with teacher name text, 0 to 4 noise level chart and colorful icons.\n\n"
            "Design details\n"
            "- Black classroom rug artwork shows Voice Level title, teacher-name text and color-coded levels 0 through 4.\n"
            "- Visible chart labels include no voice, whisper, table talk, groups and outside.\n"
            "- SEO copy focuses on visible voice-level chart artwork without unsupported material, cleaning, reverse-side or surface-performance claims."
        ),
    },
    "personalized-teachers-classroom-rug-custom-teacher-name-c-design-06": {
        "title": "Personalized Behavior Champs Rug with School Icons",
        "meta": "Customize a Behavior Champs classroom rug with teacher name text, red black artwork and kindness prompt icons.",
        "description": (
            "Customize a Behavior Champs classroom rug with teacher name text, red black artwork and kindness prompt icons.\n\n"
            "Design details\n"
            "- Bold red and black classroom rug artwork shows Behavior Champs title with teacher-name classroom rules text.\n"
            "- Visible prompts include be a great listener, show kindness, share and take turns, follow directions and helping hands.\n"
            "- SEO copy describes visible classroom prompt artwork only, without safety, therapeutic, material or surface-performance claims."
        ),
    },
    "personalized-teacher-classroom-rug-custom-back-to-school-design-101": {
        "title": "Personalized Neon Classroom Rug with School Icons",
        "meta": "Customize a neon classroom welcome rug with teacher name text, chalkboard styling, globe, backpack and pencil icons.",
        "description": (
            "Customize a neon classroom welcome rug with teacher name text, chalkboard styling, globe, backpack and pencil icons.\n\n"
            "Design details\n"
            "- Black chalkboard-style rug artwork shows neon Welcome to teacher-name classroom wording with bright school icons.\n"
            "- Visible icons include globe, backpack, pencils, books, notebook and science flask.\n"
            "- SEO copy focuses on visible classroom welcome artwork without unsupported material, cleaning, reverse-side or surface-performance claims."
        ),
    },
    "personalized-teacher-classroom-rug-custom-back-to-school-design-102": {
        "title": "Personalized Pastel Kindness Rug with Bow Artwork",
        "meta": "Customize a pastel kindness classroom rug with teacher name text, bows, daisies, crayon words and checkered border.",
        "description": (
            "Customize a pastel kindness classroom rug with teacher name text, bows, daisies, crayon words and checkered border.\n\n"
            "Design details\n"
            "- Pastel classroom rug artwork shows aqua checkered border, pink bows, daisies and teacher-name classroom text.\n"
            "- Visible crayon words include kind, brave, special, smart and loved, with Color the World with Kindness wording.\n"
            "- SEO copy describes visible kindness-themed artwork only, without safety, therapeutic, material or surface-performance claims."
        ),
    },
    "personalized-teacher-classroom-rug-custom-back-to-school-design-103": {
        "title": "Personalized Rainbow Classroom Rug with Chalkboard Art",
        "meta": "Customize a rainbow classroom welcome rug with teacher name text, chalkboard style, apple, books and pencil icons.",
        "description": (
            "Customize a rainbow classroom welcome rug with teacher name text, chalkboard style, apple, books and pencil icons.\n\n"
            "Design details\n"
            "- Black chalkboard-style rug artwork shows Welcome to teacher-name classroom wording with rainbow and school icons.\n"
            "- Visible icons include apple, books, pencils, science flask, clouds and colorful striped border.\n"
            "- SEO copy focuses on visible classroom welcome artwork without unsupported material, cleaning, reverse-side or surface-performance claims."
        ),
    },
    "personalized-teacher-classroom-rug-custom-back-to-school-design-104": {
        "title": "Personalized Teacher Welcome Rug with Chalkboard Art",
        "meta": "Customize a teacher welcome rug with classroom name text, chalkboard styling, bunting flags, apple and books.",
        "description": (
            "Customize a teacher welcome rug with classroom name text, chalkboard styling, bunting flags, apple and books.\n\n"
            "Design details\n"
            "- Black chalkboard-style rug artwork shows Welcome to teacher-name classroom wording with blue border.\n"
            "- Visible decorative elements include bunting flags, apple, pencil cup, books and heart icons.\n"
            "- SEO copy stays tied to visible teacher welcome artwork without unsupported material, cleaning, reverse-side or surface-performance claims."
        ),
    },
    "personalized-teacher-classroom-rug-custom-back-to-school-design-106": {
        "title": "Personalized Crayon Classroom Rug with Affirmation Words",
        "meta": "Customize a crayon classroom rug with teacher name text, In This Classroom wording, rainbow border and school icons.",
        "description": (
            "Customize a crayon classroom rug with teacher name text, In This Classroom wording, rainbow border and school icons.\n\n"
            "Design details\n"
            "- Black classroom rug artwork shows In This Classroom You Are wording with large teacher name at the bottom.\n"
            "- Visible crayon words include brave, helpful, kind, unique, smart and loved, with rainbow border and school icons.\n"
            "- SEO copy describes visible classroom artwork only, without safety, therapeutic, material or surface-performance claims."
        ),
    },
    "personalized-teacher-classroom-rug-custom-back-to-school-design-107": {
        "title": "Personalized Learning Day Rug with Chalk School Icons",
        "meta": "Customize a Good Day for Learning classroom rug with teacher name strip, chalk-style school icons and rainbow art.",
        "description": (
            "Customize a Good Day for Learning classroom rug with teacher name strip, chalk-style school icons and rainbow art.\n\n"
            "Design details\n"
            "- Black chalk-style classroom rug artwork shows It's a Good Day for Learning wording with a large pencil name strip.\n"
            "- Visible panel icons include sun, apple, rainbow, crayons, scissors, globe, book and flowers.\n"
            "- SEO copy focuses on visible learning-day artwork without unsupported material, cleaning, reverse-side or surface-performance claims."
        ),
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "anti slip", "non-slip",
    "non slip", "non-skid", "non skid", "kid friendly", "pet friendly",
    "safe durable", "machine-washable", "machine washable", "washable ",
    "quick-dry", "quick dry", "memory foam", "microfiber", "velvet",
    "stain", "fade resistant", "easy clean", "soft fabric", "backing",
    "cleaning visuals", "cleaning graphic", "thickened", "bound edge",
    "support ", "supports ", "safe space", "calm classroom",
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
        ws.cell(row_idx, cols["review_reason"], "Revision R018 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R018 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        found = {c[0] for c in changed}
        raise RuntimeError(f"Did not update all handles. Missing: {sorted(set(REVISIONS) - found)}")

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    else:
        log = wb["Revision_Log"]
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R018"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R018", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R018",
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
        "next_step": "Bắt đầu revision R019 or Re-QA revision R018",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R018 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R018.",
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
        "r018_log_rows": counts.get("R018", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

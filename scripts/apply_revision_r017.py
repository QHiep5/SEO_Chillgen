from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R016/SEO_Product_Optimization_revision_R016.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R017"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R017.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R017"
MANIFEST = RUN_DIR / "revision_R017_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R017_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-tropical-beach-rug-custom-family-name-tropic-design-02": {
        "title": "Personalized Tropical Beach Rug with Family Name",
        "meta": "Customize a tropical beach rug with family name oasis text, palm leaves, bright flowers, flamingos and ocean colors.",
        "description": (
            "Customize a tropical beach rug with family name oasis text, palm leaves, bright flowers, flamingos and ocean colors.\n\n"
            "Design details\n"
            "- Rectangular tropical rug artwork shows Welcome to the Browns' Oasis text with palm leaves, flowers, flamingos and ocean background.\n"
            "- Product images include living room, bedroom and close-up design views plus size reference graphics.\n"
            "- SEO copy stays tied to visible tropical family-name artwork without unsupported material, cleaning, reverse-side or surface-performance claims."
        ),
    },
    "personalized-tropical-beach-rug-custom-family-name-tropic-design-05": {
        "title": "Personalized Surfboard Beach Rug with Family Name",
        "meta": "Customize a surfboard beach rug with family name text, palm trees, flip-flops, wood sign artwork and striped colors.",
        "description": (
            "Customize a surfboard beach rug with family name text, palm trees, flip-flops, wood sign artwork and striped colors.\n\n"
            "Design details\n"
            "- Rectangular beach rug artwork shows surfboards, palm trees, flip-flops and personalized family-name welcome text.\n"
            "- Gallery images include coastal room mockups, close-up views, size reference graphics and lifestyle scenes.\n"
            "- SEO copy focuses on visible surfboard beach artwork without unsupported material, cleaning, reverse-side or surface-performance claims."
        ),
    },
    "personalized-tropical-beach-rug-custom-family-name-tropic-design-01": {
        "title": "Personalized Flip Flop Beach Rug with Happy Place Text",
        "meta": "Customize a flip flop beach rug with family name text, Happy Place wording, sand, ocean, starfish and seashell art.",
        "description": (
            "Customize a flip flop beach rug with family name text, Happy Place wording, sand, ocean, starfish and seashell art.\n\n"
            "Design details\n"
            "- Rectangular beach rug artwork shows colorful flip-flops, palm leaves, starfish, seashells and turquoise ocean colors.\n"
            "- Product images include coastal room mockups, size reference graphics, close-up views and lifestyle scenes.\n"
            "- SEO copy stays tied to visible flip flop beach artwork without unsupported material, cleaning, reverse-side or surface-performance claims."
        ),
    },
    "personalized-tropical-beach-rug-custom-family-name-tropic-design-08": {
        "title": "Personalized Flamingo Paradise Rug with Family Name",
        "meta": "Customize a flamingo paradise rug with family name text, pink flamingo, hibiscus flowers, palm leaves and beach sign.",
        "description": (
            "Customize a flamingo paradise rug with family name text, pink flamingo, hibiscus flowers, palm leaves and beach sign.\n\n"
            "Design details\n"
            "- Rectangular tropical rug artwork shows a pink flamingo, red Paradise sign, family name text, flowers and palm leaves.\n"
            "- Gallery images show coastal room mockups, close-up panels, size reference graphics and lifestyle scenes.\n"
            "- SEO copy focuses on visible flamingo paradise artwork without unsupported material, cleaning, reverse-side or surface-performance claims."
        ),
    },
    "personalized-calming-corner-rug-for-classroom-custom-tea-design-101": {
        "title": "Classroom Mindfulness Rug with Positive Prompt Art",
        "meta": "Decorate a classroom area with a mindfulness rug featuring positive-thinking text, breathe prompt and doodle icons.",
        "description": (
            "Decorate a classroom area with a mindfulness rug featuring positive-thinking text, breathe prompt and doodle icons.\n\n"
            "Design details\n"
            "- White classroom rug artwork includes positive thinking, relax, breathe, cloud gazing, plant, cook and listening prompts.\n"
            "- Product images include classroom scenes, child group photos, close-up panels and room mockups.\n"
            "- SEO copy describes visible prompt artwork only, without therapeutic, safety, material, cleaning or surface-performance claims."
        ),
    },
    "personalized-calming-corner-rug-for-classroom-custom-tea-design-102": {
        "title": "Classroom Feelings Check In Rug with Battery Meter",
        "meta": "Decorate a classroom area with a feelings check in rug featuring battery meter, mood faces and coping prompt text.",
        "description": (
            "Decorate a classroom area with a feelings check in rug featuring battery meter, mood faces and coping prompt text.\n\n"
            "Design details\n"
            "- Classroom rug artwork reads Check Your Battery with colored battery levels, feeling faces and prompt text panels.\n"
            "- Gallery images include classroom scenes, playroom mockups, child group photos and close-up design panels.\n"
            "- SEO copy describes visible SEL-style artwork only, without therapeutic, safety, material, cleaning or surface-performance claims."
        ),
    },
    "personalized-calming-corner-rug-for-classroom-custom-tea-design-103": {
        "title": "Personalized Be Who You Are Classroom Rug with Hearts",
        "meta": "Customize a classroom rug with teacher name text, Be Who You Are wording, colorful hearts and confetti artwork.",
        "description": (
            "Customize a classroom rug with teacher name text, Be Who You Are wording, colorful hearts and confetti artwork.\n\n"
            "Design details\n"
            "- White classroom rug artwork reads This Is a Space to Be Who You Are with teacher name text, hearts and confetti.\n"
            "- Product images include classroom and playroom scenes, child group photos and close-up design panels.\n"
            "- SEO copy describes visible classroom artwork only, without safety, therapeutic, material, cleaning or surface-performance claims."
        ),
    },
    "personalized-teachers-classroom-rug-custom-teacher-name-c-design-03": {
        "title": "Personalized Coping Tools Rug with Monster Prompts",
        "meta": "Customize a coping tools classroom rug with teacher name text, colorful monster icons and breathe move focus prompts.",
        "description": (
            "Customize a coping tools classroom rug with teacher name text, colorful monster icons and breathe move focus prompts.\n\n"
            "Design details\n"
            "- Pastel classroom rug artwork shows Coping Tools title with monster characters and prompts such as breathe, move, focus, squeeze and talk.\n"
            "- Gallery images include classroom mockups, size reference graphics and close-up views.\n"
            "- SEO copy describes visible prompt artwork only, without therapeutic, safety, material, cleaning or surface-performance claims."
        ),
    },
    "personalized-teachers-classroom-rug-custom-teacher-name-c-design-04": {
        "title": "Personalized Kindness Rules Rug with Checker Border",
        "meta": "Customize a kindness classroom rules rug with teacher name text, illustrated children, school icons and checker border.",
        "description": (
            "Customize a kindness classroom rules rug with teacher name text, illustrated children, school icons and checker border.\n\n"
            "Design details\n"
            "- Classroom rules rug artwork shows positive behavior prompts, illustrated children, school icons and black white checker border.\n"
            "- Product images include classroom mockups, child group scenes, close-up views and size reference panels.\n"
            "- SEO copy focuses on visible classroom rules artwork without safety, material, cleaning, reverse-side or surface-performance claims."
        ),
    },
    "personalized-teachers-classroom-rug-custom-teacher-name-c-design-01": {
        "title": "Classroom Zone of Tolerance Rug with Three Panels",
        "meta": "Decorate a classroom area with a Zone of Tolerance rug featuring overwhelmed, just right and shut down panels.",
        "description": (
            "Decorate a classroom area with a Zone of Tolerance rug featuring overwhelmed, just right and shut down panels.\n\n"
            "Design details\n"
            "- Classroom rug artwork shows My Zone of Tolerance headline with panels for overwhelmed, just right and shut down.\n"
            "- Gallery images include classroom mockups, size reference panels, close-up views and feeling reminder artwork.\n"
            "- SEO copy describes visible SEL-style artwork only, without therapeutic, safety, material, cleaning or surface-performance claims."
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
        ws.cell(row_idx, cols["review_reason"], "Revision R017 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R017 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        found = {c[0] for c in changed}
        raise RuntimeError(f"Did not update all handles. Missing: {sorted(set(REVISIONS) - found)}")

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    else:
        log = wb["Revision_Log"]
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R017"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R017", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R017",
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
        "next_step": "Bắt đầu revision R018 or Re-QA revision R017",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R017 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R017.",
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
        "r017_log_rows": counts.get("R017", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

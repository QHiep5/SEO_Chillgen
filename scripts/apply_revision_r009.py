from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R008/SEO_Product_Optimization_revision_R008.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R009"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R009.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R009"
MANIFEST = RUN_DIR / "revision_R009_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R009_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "wolf-galaxy-runner-rug-a1fc801242": {
        "title": "Fire And Ice Wolf Runner Rug with Flaming Artwork",
        "meta": "Style a hallway or room with a fire and ice wolf runner rug featuring a dark wolf silhouette, flame edges and icy blue contrast.",
        "description": (
            "Style a hallway or room with a fire and ice wolf runner rug featuring a dark wolf silhouette, flame edges and icy blue contrast.\n\n"
            "Design details\n"
            "- Rectangular runner-style rug artwork shows a wolf shape split between orange fire and cool blue ice effects.\n"
            "- Gallery images show the design in bedroom, sofa and hallway-style room mockups.\n"
            "- SEO copy stays tied to visible wolf, fire and ice artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-road-map-area-rug-81bfb9cda4-81bfb9cda4": {
        "title": "Personalized Road Map Play Rug with Alphabet Border",
        "meta": "Customize a colorful road map play rug with child name text, town roads, buildings, vehicles and an alphabet border.",
        "description": (
            "Customize a colorful road map play rug with child name text, town roads, buildings, vehicles and an alphabet border.\n\n"
            "Design details\n"
            "- Bright kids road map design includes winding streets, houses, school-style scenes, vehicles and letter icons.\n"
            "- Product images show classroom and playroom-style mockups plus a size chart and close-up views.\n"
            "- SEO copy focuses on the visible personalized play mat artwork without unsupported cleaning, material, backing or safety claims."
        ),
    },
    "galaxy-wolves-wild-wolf-print-area-rug-c951e785f7-c951e785f7": {
        "title": "Blue Moon Wolves Lake Rug with Night Forest Artwork",
        "meta": "Decorate with a blue moon wolves lake rug featuring howling wolves, moonlit water, rocks and dark forest scenery.",
        "description": (
            "Decorate with a blue moon wolves lake rug featuring howling wolves, moonlit water, rocks and dark forest scenery.\n\n"
            "Design details\n"
            "- Blue and black rug artwork shows wolf silhouettes howling near a lake under a bright moon.\n"
            "- Gallery images show the design in bedroom, sofa and living room-style mockups.\n"
            "- SEO copy stays tied to the visible wolf lake scene without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-galaxy-wolves-welcome-mat-5ce46e57d9-5ce46e57d9": {
        "title": "Snow Wolf Close Up Rug with Gray Portrait Artwork",
        "meta": "Decorate with a snow wolf close up rug featuring a gray wolf face, falling snow effect and neutral room mockups.",
        "description": (
            "Decorate with a snow wolf close up rug featuring a gray wolf face, falling snow effect and neutral room mockups.\n\n"
            "Design details\n"
            "- Gray wolf portrait artwork fills the rug with snowy texture and soft neutral color tones.\n"
            "- Gallery images show sofa, bedroom and living room-style scenes with the same close-up wolf design.\n"
            "- SEO copy focuses on the visible snow wolf portrait without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-bear-paw-print-area-rug-a43b3e3594": {
        "title": "Custom Bear Paw Wood Slice Rug with Family Name",
        "meta": "Personalize a bear paw wood slice rug with family name text, rustic wood ring artwork and paw print centerpiece.",
        "description": (
            "Personalize a bear paw wood slice rug with family name text, rustic wood ring artwork and paw print centerpiece.\n\n"
            "Design details\n"
            "- Round wood-slice style rug shows a dark bear paw print, rustic ring texture and custom family name text.\n"
            "- Product images include room mockups, personalization examples, size visuals and close-up detail panels.\n"
            "- SEO copy stays tied to the visible bear paw and wood-slice design without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-road-map-area-rug-play-mat-9cf6abf61b": {
        "title": "City Road Map Play Rug with Cars and Alphabet Icons",
        "meta": "Customize a city road map play rug with name text, circular streets, cars, buildings, alphabet icons and classroom scenes.",
        "description": (
            "Customize a city road map play rug with name text, circular streets, cars, buildings, alphabet icons and classroom scenes.\n\n"
            "Design details\n"
            "- Colorful road map artwork includes looping roads, town icons, vehicles, signs and alphabet border elements.\n"
            "- Gallery images show playroom and classroom-style scenes plus close-up and size reference panels.\n"
            "- SEO copy focuses on the visible personalized road map design without unsupported cleaning, material, backing or safety claims."
        ),
    },
    "wheel-of-feelings-and-emotions-round-rug-ec03e55978": {
        "title": "I Feel Emotions Round Rug with Classroom Faces",
        "meta": "Decorate a classroom area with an I Feel emotions round rug featuring colorful feeling words, face icons and circle layout.",
        "description": (
            "Decorate a classroom area with an I Feel emotions round rug featuring colorful feeling words, face icons and circle layout.\n\n"
            "Design details\n"
            "- Round classroom rug design shows I Feel center text with illustrated faces and feelings around the circle.\n"
            "- Product images include classroom-style mockups, close-up views and size reference graphics.\n"
            "- SEO copy stays tied to the visible emotions chart artwork without unsupported material, backing, therapeutic or safety claims."
        ),
    },
    "wheel-of-feelings-and-emotions-round-rug-01636a663b": {
        "title": "Teaching Clock Round Rug with Hour and Minute Marks",
        "meta": "Teach time visually with a round clock rug featuring hour numbers, minute marks, color segments and classroom mockups.",
        "description": (
            "Teach time visually with a round clock rug featuring hour numbers, minute marks, color segments and classroom mockups.\n\n"
            "Design details\n"
            "- Round educational clock design includes large hour numbers, minute labels and bright segmented colors.\n"
            "- Product images show classroom-style scenes, close-up details and size reference panels.\n"
            "- SEO copy focuses on the visible clock-learning artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "wheel-of-feelings-and-emotions-round-rug-8a745f7f62": {
        "title": "Color Emotions Wheel Rug with Feelings Face Chart",
        "meta": "Decorate a learning space with a color emotions wheel rug featuring feelings words, face icons and rainbow segments.",
        "description": (
            "Decorate a learning space with a color emotions wheel rug featuring feelings words, face icons and rainbow segments.\n\n"
            "Design details\n"
            "- Round emotions chart artwork includes colorful feeling zones, expressive face icons and classroom-style layout.\n"
            "- Gallery images show the design in classroom mockups, close-up panels and size reference graphics.\n"
            "- SEO copy stays tied to visible SEL-style artwork without unsupported material, backing, therapeutic or safety claims."
        ),
    },
    "custom-composition-notebook-classroom-rug-kids-850c621ca3-850c621ca3": {
        "title": "Personalized Composition Notebook Rug with Teacher Name",
        "meta": "Customize a composition notebook classroom rug with teacher name text, pencil bow, apple icon and crayon border artwork.",
        "description": (
            "Customize a composition notebook classroom rug with teacher name text, pencil bow, apple icon and crayon border artwork.\n\n"
            "Design details\n"
            "- Notebook paper design includes lined page artwork, teacher name personalization, pencil bow and apple illustration.\n"
            "- Product images show classroom-style scenes, child group mockups, round preview panels and rug size references.\n"
            "- SEO copy focuses on the visible personalized classroom design without unsupported material, backing or surface-performance claims."
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
        ws.cell(row_idx, cols["review_reason"], "Revision R009 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R009 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        found = {c[0] for c in changed}
        raise RuntimeError(f"Did not update all handles. Missing: {sorted(set(REVISIONS) - found)}")

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    else:
        log = wb["Revision_Log"]
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R009"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R009", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R009",
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
        "next_step": "Bắt đầu revision R010 or Re-QA revision R009",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R009 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R009.",
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
        "r009_log_rows": counts.get("R009", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

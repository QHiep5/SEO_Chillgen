from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R014/SEO_Product_Optimization_revision_R014.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R015"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R015.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R015"
MANIFEST = RUN_DIR / "revision_R015_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R015_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-halloween-3d-optical-illusion-ghost-round-rug-design-12": {
        "title": "Lava Skeleton Halloween Round Rug with Pit Illusion",
        "meta": "Decorate with a lava skeleton Halloween round rug featuring cracked fire pit artwork, crawling skeleton and skull details.",
        "description": (
            "Decorate with a lava skeleton Halloween round rug featuring cracked fire pit artwork, crawling skeleton and skull details.\n\n"
            "Design details\n"
            "- Round Halloween rug artwork shows a skeleton crawling from a glowing lava pit with cracked black rock texture.\n"
            "- Product images include living room, dining area and seasonal room mockups plus close-up and size reference panels.\n"
            "- SEO copy stays tied to visible Halloween illusion artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-dragon-shaped-area-rug-personalized-library-wester-design-02": {
        "title": "Personalized Blue Dragon Book Rug with Name Art",
        "meta": "Personalize a blue dragon book rug with custom name text, curled dragon artwork, open book shape and moonlit accents.",
        "description": (
            "Personalize a blue dragon book rug with custom name text, curled dragon artwork, open book shape and moonlit accents.\n\n"
            "Design details\n"
            "- Custom-shaped fantasy rug artwork shows a blue dragon curled across an open book with personalized name text.\n"
            "- Product images include floor mockups, size reference panels, close-up views and custom-name examples.\n"
            "- SEO copy focuses on visible personalized dragon book artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-dragon-shaped-area-rug-personalized-library-wester-design-08": {
        "title": "Personalized Purple Dragon Rug with Stacked Books",
        "meta": "Personalize a purple dragon rug with custom name text, stacked book artwork, flowers, vines and fantasy styling.",
        "description": (
            "Personalize a purple dragon rug with custom name text, stacked book artwork, flowers, vines and fantasy styling.\n\n"
            "Design details\n"
            "- Custom-shaped rug artwork shows a purple dragon resting on stacked books with purple flowers and green vines.\n"
            "- Gallery images show room mockups, close-up panels, size reference graphics and personalization examples.\n"
            "- SEO copy stays tied to visible dragon library artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-dragon-shaped-area-rug-personalized-library-wester-design-05": {
        "title": "Personalized Dragon Open Book Rug with Teal Wings",
        "meta": "Personalize a dragon open book rug with custom name text, teal wings, orange fantasy accents and illustrated pages.",
        "description": (
            "Personalize a dragon open book rug with custom name text, teal wings, orange fantasy accents and illustrated pages.\n\n"
            "Design details\n"
            "- Custom-shaped rug artwork shows a teal dragon with orange wings lying across an open illustrated book.\n"
            "- Product images include room mockups, close-up panels, size reference graphics and personalization examples.\n"
            "- SEO copy focuses on visible personalized dragon book artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-dragon-shaped-area-rug-personalized-library-wester-design-07": {
        "title": "Personalized Moon Dragon Book Rug with Rose Art",
        "meta": "Personalize a moon dragon book rug with custom name text, blue gold dragon, full moon, stacked books and roses.",
        "description": (
            "Personalize a moon dragon book rug with custom name text, blue gold dragon, full moon, stacked books and roses.\n\n"
            "Design details\n"
            "- Custom-shaped rug artwork shows a blue and gold dragon before a full moon with stacked books and rose details.\n"
            "- Gallery images include room mockups, size reference panels, close-up views and personalization examples.\n"
            "- SEO copy stays tied to visible moon dragon artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-dragon-shaped-area-rug-personalized-library-wester-design-03": {
        "title": "Personalized Dragon Library Rug with Orange Flowers",
        "meta": "Personalize a dragon library rug with custom name text, purple dragon, teal wings, stacked books and orange flowers.",
        "description": (
            "Personalize a dragon library rug with custom name text, purple dragon, teal wings, stacked books and orange flowers.\n\n"
            "Design details\n"
            "- Custom-shaped rug artwork shows a purple dragon with teal wings beside stacked books, orange flowers and green leaves.\n"
            "- Product images include floor mockups, close-up panels, size reference graphics and personalized-name examples.\n"
            "- SEO copy focuses on visible dragon library artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-dragon-shaped-area-rug-personalized-library-wester-design-10": {
        "title": "Personalized Green Dragon Book Rug with Name Art",
        "meta": "Personalize a green dragon book rug with custom name text, large wings, stacked books and fantasy library artwork.",
        "description": (
            "Personalize a green dragon book rug with custom name text, large wings, stacked books and fantasy library artwork.\n\n"
            "Design details\n"
            "- Custom-shaped rug artwork shows a green dragon standing and curling across stacked books with personalized name text.\n"
            "- Gallery images include room mockups, size reference panels, close-up views and personalization examples.\n"
            "- SEO copy stays tied to visible green dragon book artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-dragon-shaped-area-rug-personalized-library-wester-design-01": {
        "title": "Personalized Red Dragon Book Rug with Red Roses",
        "meta": "Personalize a red dragon book rug with custom name text, curled dragon artwork, stacked books and red roses.",
        "description": (
            "Personalize a red dragon book rug with custom name text, curled dragon artwork, stacked books and red roses.\n\n"
            "Design details\n"
            "- Custom-shaped rug artwork shows a red dragon curled around stacked books with roses and personalized name text.\n"
            "- Product images include room mockups, close-up panels, size reference graphics and custom-name examples.\n"
            "- SEO copy focuses on visible red dragon book artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-octopus-welcome-mat-for-front-door-cf228fd538": {
        "title": "Octopus Mandala Coastal Rug with Turquoise Art",
        "meta": "Decorate with an octopus mandala coastal rug featuring turquoise nautical artwork, white octopus and ornate line details.",
        "description": (
            "Decorate with an octopus mandala coastal rug featuring turquoise nautical artwork, white octopus and ornate line details.\n\n"
            "Design details\n"
            "- Rectangular rug artwork shows a white octopus over turquoise nautical mandala-style details and circular line patterns.\n"
            "- Product images include room mockups, close-up panels, family scenes and size reference graphics.\n"
            "- SEO copy stays tied to visible octopus mandala artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-octopus-welcome-mat-for-front-door-and-entryway-6c5a0747da": {
        "title": "Red Octopus Mosaic Rug with Dark Teal Artwork",
        "meta": "Decorate with a red octopus mosaic rug featuring dark teal background, high-contrast tentacles and patterned sea art.",
        "description": (
            "Decorate with a red octopus mosaic rug featuring dark teal background, high-contrast tentacles and patterned sea art.\n\n"
            "Design details\n"
            "- Rectangular rug artwork shows a red octopus with bold patterned tentacles over a dark teal ocean-style background.\n"
            "- Gallery images include room mockups, close-up views, family scenes and size reference panels.\n"
            "- SEO copy focuses on visible red octopus artwork without unsupported material, backing or surface-performance claims."
        ),
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "anti slip", "non-slip",
    "non slip", "non-skid", "non skid", "kid friendly", "pet friendly",
    "safe durable", "machine-washable", "machine washable", "washable ",
    "quick-dry", "quick dry", "memory foam", "microfiber", "velvet",
    "stain", "fade resistant", "easy clean", "soft bound", "bound edges",
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
        ws.cell(row_idx, cols["review_reason"], "Revision R015 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R015 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        found = {c[0] for c in changed}
        raise RuntimeError(f"Did not update all handles. Missing: {sorted(set(REVISIONS) - found)}")

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    else:
        log = wb["Revision_Log"]
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R015"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R015", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R015",
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
        "next_step": "Bắt đầu revision R016 or Re-QA revision R015",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R015 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R015.",
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
        "r015_log_rows": counts.get("R015", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

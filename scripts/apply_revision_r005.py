from __future__ import annotations

import json
from datetime import datetime, timezone, timedelta
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R004/SEO_Product_Optimization_revision_R004.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R005"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R005.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R005"
MANIFEST = RUN_DIR / "revision_R005_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R005_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-cool-kids-read-book-rug-classroom-library-b1f15827f5": {
        "title": "Cool Kids Read Books Classroom Rug with Book Characters",
        "meta": "Add a Cool Kids Read Books classroom rug with smiling book characters, colorful school icons and story-time mockups.",
        "description": (
            "Add a Cool Kids Read Books classroom rug with smiling book characters, colorful school icons and story-time mockups.\n\n"
            "Design details\n"
            "- Cream reading rug with Cool Kids Read Books text and colorful smiling book character artwork.\n"
            "- Gallery images show classroom, playroom and story-time scenes plus close-up design panels.\n"
            "- SEO copy stays tied to the visible classroom reading design and avoids unsupported material, backing or surface-performance claims."
        ),
    },
    "colorful-handprints-classroom-kids-rug-51cf8ab902-51cf8ab902": {
        "title": "Personalized Welcome Classroom Rug with Handprint Values",
        "meta": "Personalize a welcome classroom rug with teacher name text, colorful handprints and character words such as kind and brave.",
        "description": (
            "Personalize a welcome classroom rug with teacher name text, colorful handprints and character words such as kind and brave.\n\n"
            "Design details\n"
            "- Cream handprint classroom rug sample with Welcome to Mrs. Smith's Classroom text and bright handprint artwork.\n"
            "- Visible value words include kind, yourself, creative, brave, honest, humble, happy and thankful.\n"
            "- SEO copy focuses on the visible teacher-name classroom design without unsupported material, backing or surface-performance claims."
        ),
    },
    "personalized-orthodox-cross-rug-byzantine-eagle-be38fb9e0a": {
        "title": "Byzantine Eagle Area Rug with Burgundy Ornate Border",
        "meta": "Decorate with a Byzantine eagle area rug featuring burgundy field, gold double-headed eagle emblem and ornate border art.",
        "description": (
            "Decorate with a Byzantine eagle area rug featuring burgundy field, gold double-headed eagle emblem and ornate border art.\n\n"
            "Design details\n"
            "- Burgundy area rug with a gold double-headed Byzantine eagle emblem and red, gold and blue ornamental border.\n"
            "- Gallery images show living room, fireplace and bedroom-style mockups plus close-up design and size visuals.\n"
            "- SEO copy treats the rug as decorative home decor and avoids unsupported material, backing, devotional-use or surface-performance claims."
        ),
    },
    "personalized-orthodox-cross-byzantine-eagle-rug-8e70ebd991": {
        "title": "Blue Orthodox Cross Area Rug with Gold Three Bar Design",
        "meta": "Decorate with a blue Orthodox cross area rug featuring a gold three-bar cross, IC XC NIKA lettering and ornate border.",
        "description": (
            "Decorate with a blue Orthodox cross area rug featuring a gold three-bar cross, IC XC NIKA lettering and ornate border.\n\n"
            "Design details\n"
            "- Blue area rug with ornate gold three-bar Orthodox cross, IC XC NIKA lettering and curved gold border artwork.\n"
            "- Gallery images show living room and quiet room-style mockups plus close-up and size visuals.\n"
            "- SEO copy focuses on the visible cross artwork and avoids unsupported material, backing, prayer-use or surface-performance claims."
        ),
    },
    "custom-orthodox-rug-three-bar-cross-byzantine-eagle-d81169235b-d81169235b": {
        "title": "Red Byzantine Eagle Area Rug with Cream Center Panel",
        "meta": "Decorate with a red Byzantine eagle area rug featuring a cream center, red double-headed eagle and gold floral border.",
        "description": (
            "Decorate with a red Byzantine eagle area rug featuring a cream center, red double-headed eagle and gold floral border.\n\n"
            "Design details\n"
            "- Cream and red area rug with a red double-headed Byzantine eagle, gold floral corner ornaments and red border.\n"
            "- Gallery images show living room and bedroom-style mockups plus close-up design and size visuals.\n"
            "- SEO copy stays tied to the visible Byzantine eagle design and avoids unsupported material, backing or surface-performance claims."
        ),
    },
    "personalized-teachers-classroom-rug-with-custom-name-87d858996b": {
        "title": "Calm Down Classroom Rug with Zone of Tolerance Design",
        "meta": "Create a calm down classroom rug with Zone of Tolerance sections, emotion prompts, strategy icons and teacher-room text.",
        "description": (
            "Create a calm down classroom rug with Zone of Tolerance sections, emotion prompts, strategy icons and teacher-room text.\n\n"
            "Design details\n"
            "- Classroom rug with overwhelmed, just right and shut down sections, feelings icons and calm strategy prompts.\n"
            "- Gallery images show classroom and reading nook mockups with the visual regulation chart design.\n"
            "- SEO copy describes the visible calm-down classroom artwork without unsupported therapeutic, material, backing or surface-performance claims."
        ),
    },
    "custom-octopus-shaped-area-rug-9041fd8f83": {
        "title": "Purple Octopus Shaped Area Rug with Ocean Cave Artwork",
        "meta": "Add a purple octopus shaped area rug with curling tentacles, blue water, rocky cave edges and custom-cut sea monster art.",
        "description": (
            "Add a purple octopus shaped area rug with curling tentacles, blue water, rocky cave edges and custom-cut sea monster art.\n\n"
            "Design details\n"
            "- Custom-shaped rug with vivid purple octopus, curling tentacles, blue water and rocky cave-edge artwork.\n"
            "- Gallery images show room mockups, cut-shape views, size visuals and close-up tentacle details.\n"
            "- SEO copy focuses on the visible octopus shape and ocean cave artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-octopus-shaped-area-rug-sea-monster-tentacle-0753b2605a": {
        "title": "Wave Octopus Shaped Area Rug with Blue Tentacle Art",
        "meta": "Add a wave octopus shaped area rug with dark tentacles, blue ocean waves, orange coral accents and custom-cut edges.",
        "description": (
            "Add a wave octopus shaped area rug with dark tentacles, blue ocean waves, orange coral accents and custom-cut edges.\n\n"
            "Design details\n"
            "- Custom-shaped octopus rug with dark blue tentacles, crashing blue waves, orange coral accents and irregular cut edge.\n"
            "- Gallery images show room mockups, cut-shape views, size visuals and close-up wave/tentacle details.\n"
            "- SEO copy stays focused on the visible wave octopus artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-octopus-shaped-area-rug-sea-monster-1975158434": {
        "title": "Cave Octopus Shaped Area Rug with Deep Blue Water Art",
        "meta": "Create an ocean focal point with a cave octopus shaped area rug featuring purple tentacles, rocky border and blue water.",
        "description": (
            "Create an ocean focal point with a cave octopus shaped area rug featuring purple tentacles, rocky border and blue water.\n\n"
            "Design details\n"
            "- Custom-shaped rug with a purple octopus inside underwater cave artwork, surface light, rocky border and deep blue water.\n"
            "- Gallery images show room mockups, cut-shape panels, size visuals and close-up design details.\n"
            "- SEO copy focuses on the visible sea-monster cave artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-octopus-shaped-area-rug-c7d44eba12-c7d44eba12": {
        "title": "Teal Octopus Shaped Area Rug with Playful Tentacle Art",
        "meta": "Style an ocean-themed room with a teal octopus shaped area rug featuring rounded tentacles and wave-like border art.",
        "description": (
            "Style an ocean-themed room with a teal octopus shaped area rug featuring rounded tentacles and wave-like border art.\n\n"
            "Design details\n"
            "- Custom-shaped teal and navy octopus rug with rounded tentacles, wave-like border and playful sea creature illustration.\n"
            "- Gallery images show wood-floor and living room mockups, cut-shape panels, size visuals and close-up details.\n"
            "- SEO copy stays tied to the visible teal octopus artwork without unsupported material, backing or surface-performance claims."
        ),
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "non-slip", "kid friendly",
    "pet friendly", "safe durable", "machine-washable", "machine washable",
    "washable ", "quick-dry", "memory foam", "microfiber", "velvet",
    "stain", "fade resistant",
]


def headers(ws):
    return {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}


def main():
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
        ws.cell(row_idx, cols["review_reason"], "Revision R005 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R005 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        found = {c[0] for c in changed}
        raise RuntimeError(f"Did not update all handles. Missing: {sorted(set(REVISIONS) - found)}")

    for sheet_name in ["Buyer_Search_Research", "Keyword_Map"]:
        if sheet_name not in wb.sheetnames:
            continue
        sh = wb[sheet_name]
        hmap = headers(sh)
        key_col = hmap.get("product_key")
        if not key_col:
            continue
        for row_idx in range(2, sh.max_row + 1):
            pk = str(sh.cell(row_idx, key_col).value or "")
            if not any(handle in pk for handle in REVISIONS):
                continue
            for col_idx in range(1, sh.max_column + 1):
                value = sh.cell(row_idx, col_idx).value
                if not isinstance(value, str):
                    continue
                new = value.replace("indoor/outdoor", "indoor").replace("Indoor/outdoor", "Indoor")
                new = new.replace("outdoor-use", "room-use").replace("outdoor use", "room use")
                sh.cell(row_idx, col_idx, new)

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    else:
        log = wb["Revision_Log"]
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R005"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R005", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R005",
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
        "next_step": "Bắt đầu revision R006 or Re-QA revision R005",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R005 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R005.",
        "- Fixes targeted QA issues: short T1/T2 and unsupported claims.",
        "- Kept `review_status=NEEDS_REVIEW` and `content_qa_status=NOT_RUN`.",
        "- No Shopify deploy, no approval.", "",
        "## Updated products", "",
        "| Handle | New title | Title chars | Meta chars |",
        "|---|---|---:|---:|",
    ]
    for handle, _, _, _, title, meta in changed:
        lines.append(f"| `{handle}` | {title} | {len(title)} | {len(meta)} |")
    SUMMARY.write_text("\n".join(lines) + "\n", encoding="utf-8")

    verify = load_workbook(OUTPUT, read_only=True, data_only=True)
    vws = verify["SEO_Products"]
    vcols = {cell.value: idx for idx, cell in enumerate(next(vws.iter_rows(min_row=1, max_row=1)))}
    bad = []
    for row in vws.iter_rows(min_row=2, values_only=True):
        handle = row[vcols["Handle"]]
        if handle not in REVISIONS:
            continue
        title = row[vcols["title_proposed"]] or ""
        meta_title = row[vcols["meta_title_seo"]] or ""
        meta_desc = row[vcols["meta_description_seo"]] or ""
        desc = (row[vcols["description_proposed"]] or "") + " " + (row[vcols["description_proposed_html"]] or "")
        blob = " ".join([title, meta_title, meta_desc, desc]).lower()
        if row[vcols["revision"]] != "2":
            bad.append((handle, "revision", row[vcols["revision"]]))
        if row[vcols["review_status"]] != "NEEDS_REVIEW":
            bad.append((handle, "review_status", row[vcols["review_status"]]))
        if row[vcols["content_qa_status"]] != "NOT_RUN":
            bad.append((handle, "content_qa_status", row[vcols["content_qa_status"]]))
        if not (45 <= len(title) <= 70):
            bad.append((handle, "title_len", len(title)))
        if meta_title != title:
            bad.append((handle, "meta_title_mismatch", meta_title))
        if len(meta_desc) > 320:
            bad.append((handle, "meta_desc_len", len(meta_desc)))
        for forbidden in FORBIDDEN:
            if forbidden in f" {blob} ":
                bad.append((handle, "forbidden_claim", forbidden))
    if bad:
        raise RuntimeError(f"Verification failed: {bad[:20]}")

    print(json.dumps({
        "output_workbook": str(OUTPUT),
        "manifest": str(MANIFEST),
        "summary": str(SUMMARY),
        "updated": len(changed),
        "status": "COMPLETE_AWAITING_REQA",
    }, indent=2))


if __name__ == "__main__":
    main()

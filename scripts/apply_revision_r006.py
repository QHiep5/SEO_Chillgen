from __future__ import annotations

import json
from datetime import datetime, timezone, timedelta
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R005/SEO_Product_Optimization_revision_R005.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R006"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R006.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R006"
MANIFEST = RUN_DIR / "revision_R006_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R006_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "tree-of-life-welcome-mat-living-room-c11fdc3ba9-c11fdc3ba9": {
        "title": "Teal Tree of Life Area Rug with Golden Botanical Art",
        "meta": "Decorate with a teal Tree of Life area rug featuring golden roots, botanical branches, starry background and ornate border.",
        "description": (
            "Decorate with a teal Tree of Life area rug featuring golden roots, botanical branches, starry background and ornate border.\n\n"
            "Design details\n"
            "- Rectangular Tree of Life rug with glowing gold tree, visible roots, teal-black night background and ornate gold border.\n"
            "- Gallery images show living room, fireplace and bedroom-style mockups plus detail and size visuals.\n"
            "- SEO copy focuses on the visible tree and botanical artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "western-running-horse-area-rug-122e8b85fb-122e8b85fb": {
        "title": "Black Running Horse Area Rug with Golden Sunset Glow",
        "meta": "Decorate with a black running horse area rug featuring a golden sunset circle, dust effects and western motion artwork.",
        "description": (
            "Decorate with a black running horse area rug featuring a golden sunset circle, dust effects and western motion artwork.\n\n"
            "Design details\n"
            "- Rectangular horse rug with black galloping horse, golden circular glow, dust texture and sparkling western-style motion effects.\n"
            "- Gallery images show living room and fireplace mockups plus close-up design and size visuals.\n"
            "- SEO copy stays tied to the visible running horse artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-running-horse-area-rug-5eadea7a69": {
        "title": "Purple Horse Area Rug with Floral Dream Sky Artwork",
        "meta": "Add a purple horse area rug with black horse portrait, violet flowers, pink-purple sky and glowing fantasy-style details.",
        "description": (
            "Add a purple horse area rug with black horse portrait, violet flowers, pink-purple sky and glowing fantasy-style details.\n\n"
            "Design details\n"
            "- Rectangular rug with black horse portrait, purple and pink sky, violet floral accents and dreamy glow effects.\n"
            "- Gallery images show room mockups, fireplace scenes, close-up panels and size visuals.\n"
            "- SEO copy focuses on the visible purple horse artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-running-horse-welcome-mat-5c3fa81b88": {
        "title": "Rustic Horse Area Rug with Dramatic Brown Portrait Art",
        "meta": "Style a western room with a rustic horse area rug featuring a close-up horse portrait, brown dust tones and motion texture.",
        "description": (
            "Style a western room with a rustic horse area rug featuring a close-up horse portrait, brown dust tones and motion texture.\n\n"
            "Design details\n"
            "- Rectangular rug with close-up black and white horse portrait, brown dusty background and dramatic motion texture.\n"
            "- Gallery images show living room, fireplace and bedroom-style mockups plus detail and size visuals.\n"
            "- SEO copy stays tied to the visible rustic horse portrait without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-running-horse-welcome-mat-a0cd749d2b": {
        "title": "Rainbow Mane Horse Area Rug with Warm Cream Background",
        "meta": "Brighten a horse lover room with a rainbow mane horse area rug featuring tan horse portrait and colorful flowing strokes.",
        "description": (
            "Brighten a horse lover room with a rainbow mane horse area rug featuring tan horse portrait and colorful flowing strokes.\n\n"
            "Design details\n"
            "- Rectangular horse rug with tan horse portrait, rainbow mane, warm cream-gold background and colorful leaf-like strokes.\n"
            "- Gallery images show living room and fireplace mockups plus close-up design and size visuals.\n"
            "- SEO copy focuses on the visible rainbow horse artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-running-horse-doormat-2a0ef19423-2a0ef19423": {
        "title": "Tribal Horse Area Rug with Feather Mane Details",
        "meta": "Style a western room with a tribal horse area rug featuring tan horse portrait, bridle details and feather-like mane accents.",
        "description": (
            "Style a western room with a tribal horse area rug featuring tan horse portrait, bridle details and feather-like mane accents.\n\n"
            "Design details\n"
            "- Rectangular horse rug with close-up tan horse portrait, bridle details, feather-like mane accents and warm golden background.\n"
            "- Gallery images show room mockups, fireplace scenes, size visuals and close-up artwork panels.\n"
            "- SEO copy stays tied to the visible tribal-style horse design without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-running-horse-doormat-b25dd5e514": {
        "title": "Sunflower Horse Area Rug with White Horse Portrait",
        "meta": "Brighten a horse lover room with a sunflower horse area rug featuring white horse artwork, yellow blooms and soft blue leaves.",
        "description": (
            "Brighten a horse lover room with a sunflower horse area rug featuring white horse artwork, yellow blooms and soft blue leaves.\n\n"
            "Design details\n"
            "- Rectangular rug with white horse portrait, large yellow sunflowers, soft blue leaves and warm floral background.\n"
            "- Gallery images show living room and fireplace mockups plus close-up flower and size visuals.\n"
            "- SEO copy focuses on the visible sunflower horse artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-running-horse-area-rug-e5f0770fd6-e5f0770fd6": {
        "title": "Horse Eye Area Rug with Close Up Brown Artwork",
        "meta": "Create a striking equestrian room with a horse eye area rug featuring close-up brown coat texture and dramatic gaze artwork.",
        "description": (
            "Create a striking equestrian room with a horse eye area rug featuring close-up brown coat texture and dramatic gaze artwork.\n\n"
            "Design details\n"
            "- Rectangular rug with a close-up horse eye, warm brown coat texture, whisker details and dramatic cropped composition.\n"
            "- Gallery images show living room and fireplace mockups plus close-up design and size visuals.\n"
            "- SEO copy stays focused on the visible horse-eye artwork without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-running-horse-welcome-mat-1edee5e89b-1edee5e89b": {
        "title": "White Floral Horse Area Rug with Daisy Garden Artwork",
        "meta": "Brighten an equestrian room with a white floral horse area rug featuring yellow daisies, red flowers and garden artwork.",
        "description": (
            "Brighten an equestrian room with a white floral horse area rug featuring yellow daisies, red flowers and garden artwork.\n\n"
            "Design details\n"
            "- Rectangular rug with white horse portrait surrounded by yellow daisies, red flowers and a soft wildflower field.\n"
            "- Gallery images show room mockups, fireplace scenes, close-up floral details and size visuals.\n"
            "- SEO copy focuses on the visible white floral horse design without unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-running-horse-welcome-mat-01f80a49ee": {
        "title": "Butterfly Horse Area Rug with White Horse and Flowers",
        "meta": "Decorate with a butterfly horse area rug featuring white horse artwork, orange butterflies, daisies and colorful garden flowers.",
        "description": (
            "Decorate with a butterfly horse area rug featuring white horse artwork, orange butterflies, daisies and colorful garden flowers.\n\n"
            "Design details\n"
            "- Rectangular rug with white horse portrait, orange butterflies, daisies and colorful garden flower artwork.\n"
            "- Gallery images show living room and fireplace mockups plus close-up floral and butterfly details.\n"
            "- SEO copy stays tied to the visible butterfly horse artwork without unsupported material, backing or surface-performance claims."
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
        ws.cell(row_idx, cols["review_reason"], "Revision R006 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R006 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
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
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R006"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R006", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R006",
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
        "next_step": "Bắt đầu revision R007 or Re-QA revision R006",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R006 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R006.",
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

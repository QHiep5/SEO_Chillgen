from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R012/SEO_Product_Optimization_revision_R012.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R013"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R013.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R013"
MANIFEST = RUN_DIR / "revision_R013_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R013_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "tree-of-life-celtic-indoor-rug-89f894eb42": {
        "title": "Abstract Rainbow Tree of Life Rug with Blue Artwork",
        "meta": "Style a room with an abstract rainbow Tree of Life rug featuring colorful leaves, flowing trunk and deep blue background.",
        "description": (
            "Style a room with an abstract rainbow Tree of Life rug featuring colorful leaves, flowing trunk and deep blue background.\n\n"
            "Design details\n"
            "- Round rug artwork shows a vivid Tree of Life with red, orange, yellow and blue leaves over a navy turquoise scene.\n"
            "- Product images include room mockups, close-up design panels and size reference graphics.\n"
            "- SEO copy stays tied to visible Tree of Life artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "tree-of-life-celtic-indoor-rug-e74720f30a": {
        "title": "Teal Gold Tree of Life Rug with Celtic Knot Border",
        "meta": "Decorate with a teal gold Tree of Life rug featuring a Celtic knot border, rolling mountains, roots and branch artwork.",
        "description": (
            "Decorate with a teal gold Tree of Life rug featuring a Celtic knot border, rolling mountains, roots and branch artwork.\n\n"
            "Design details\n"
            "- Round rug artwork shows a golden Tree of Life, exposed roots, rolling green mountains and decorative knotwork border.\n"
            "- Gallery images show bedroom and living room-style mockups plus close-up and size reference panels.\n"
            "- SEO copy focuses on visible Celtic tree artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "custom-pink-flamingo-welcome-mat-98ad4e675e": {
        "title": "Pink Flamingo Moonlight Rug with Tropical Water Art",
        "meta": "Decorate with a pink flamingo moonlight rug featuring tropical water reflections, purple sky and large moon artwork.",
        "description": (
            "Decorate with a pink flamingo moonlight rug featuring tropical water reflections, purple sky and large moon artwork.\n\n"
            "Design details\n"
            "- Rectangular rug artwork shows bright pink flamingos standing in reflective water below a large moon.\n"
            "- Product images include bedroom, living room and close-up views with dark blue, purple and pink tropical colors.\n"
            "- SEO copy stays tied to visible flamingo artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "custom-pink-flamingo-welcome-mat-for-entrance-and-living-room-c2beb658ca": {
        "title": "Flamingo Heart Sunset Rug with Tropical Water Art",
        "meta": "Decorate with a flamingo heart sunset rug featuring two pink birds, orange sky, turquoise water and palm silhouettes.",
        "description": (
            "Decorate with a flamingo heart sunset rug featuring two pink birds, orange sky, turquoise water and palm silhouettes.\n\n"
            "Design details\n"
            "- Rectangular rug artwork shows two flamingos forming a heart shape in a bright sunset water scene.\n"
            "- Product images include room mockups, close-up design views and warm pink orange tropical colors.\n"
            "- SEO copy focuses on visible flamingo sunset artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "custom-japanese-koi-fish-shaped-area-rug-940fcac253": {
        "title": "Japanese Koi Pond Shaped Rug with Lily Pad Border",
        "meta": "Decorate with a Japanese koi pond shaped rug featuring orange fish, dark teal water and dense green lily pad border.",
        "description": (
            "Decorate with a Japanese koi pond shaped rug featuring orange fish, dark teal water and dense green lily pad border.\n\n"
            "Design details\n"
            "- Irregular pond-shaped rug artwork shows several orange koi fish swimming through dark teal water.\n"
            "- Product images include floor mockups, size reference panels, close-up views and green lily pad border details.\n"
            "- SEO copy stays tied to visible koi pond artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "custom-japanese-koi-fish-shaped-area-rug-925e5d0767": {
        "title": "Koi Pond Stepping Stone Rug with Orange Fish Art",
        "meta": "Decorate with a koi pond stepping stone rug featuring orange koi fish, dark water and moss green stone artwork.",
        "description": (
            "Decorate with a koi pond stepping stone rug featuring orange koi fish, dark water and moss green stone artwork.\n\n"
            "Design details\n"
            "- Oval irregular rug artwork shows orange koi fish swimming between moss green stepping stones over dark water.\n"
            "- Gallery images include floor mockups, close-up panels, custom-name examples and size reference graphics.\n"
            "- SEO copy focuses on visible koi pond artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "custom-koi-fish-shaped-area-rug-027c867fc1": {
        "title": "Pink Lotus Koi Fish Shaped Rug with Pond Artwork",
        "meta": "Decorate with a pink lotus koi fish shaped rug featuring pale koi, lily pads, lotus flowers and green border.",
        "description": (
            "Decorate with a pink lotus koi fish shaped rug featuring pale koi, lily pads, lotus flowers and green border.\n\n"
            "Design details\n"
            "- Irregular shaped rug artwork shows a soft pink pond scene with pale koi fish, lotus flowers and lily pads.\n"
            "- Product images include floor mockups, close-up detail panels, custom-name examples and size reference graphics.\n"
            "- SEO copy stays tied to visible koi and lotus artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "custom-japanese-koi-fish-shaped-area-rug-3e3d07aa66": {
        "title": "Zen Koi Pond Shaped Rug with Sand Garden Artwork",
        "meta": "Decorate with a zen koi pond shaped rug featuring blue water, koi fish, rocks, greenery and raked sand lines.",
        "description": (
            "Decorate with a zen koi pond shaped rug featuring blue water, koi fish, rocks, greenery and raked sand lines.\n\n"
            "Design details\n"
            "- Irregular shaped rug artwork combines beige raked-sand lines, blue pond water, koi fish, rocks and green plants.\n"
            "- Gallery images show floor mockups, close-up panels, size reference graphics and personalization examples.\n"
            "- SEO copy focuses on visible zen pond artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "personalized-vinyl-record-rug-with-photo-and-song-name-19f23f2bc0": {
        "title": "Custom CD Record Rug with Personalized Image Area",
        "meta": "Personalize a custom CD record rug with image placeholder, bold song-style lettering, round black disc and color rim.",
        "description": (
            "Personalize a custom CD record rug with image placeholder, bold song-style lettering, round black disc and color rim.\n\n"
            "Design details\n"
            "- Round record-style rug artwork reads Your Custom CD and Your Image Here with bright circular trim.\n"
            "- Product images include room mockups, personalization examples, picture illustration panels and close-up views.\n"
            "- SEO copy stays tied to visible custom record artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
    "personalized-dragon-round-rug-3d-effect-dcdcc98f68": {
        "title": "White Dragon Forest Round Rug with 3D Pit Artwork",
        "meta": "Decorate with a white dragon forest round rug featuring fantasy dragon artwork, evergreen trees and dark pit illusion.",
        "description": (
            "Decorate with a white dragon forest round rug featuring fantasy dragon artwork, evergreen trees and dark pit illusion.\n\n"
            "Design details\n"
            "- Round rug artwork shows a white dragon flying above a dark forest pit with evergreen trees and stone border.\n"
            "- Product images include living room and seasonal room mockups, close-up panels and size reference graphics.\n"
            "- SEO copy focuses on visible dragon fantasy artwork without unsupported material, cleaning, backing or surface-performance claims."
        ),
    },
}


FORBIDDEN = [
    "indoor/outdoor", " outdoor ", "anti-slip", "anti slip", "non-slip",
    "non slip", "non-skid", "non skid", "kid friendly", "pet friendly",
    "safe durable", "machine-washable", "machine washable", "washable ",
    "quick-dry", "quick dry", "memory foam", "microfiber", "velvet",
    "stain", "fade resistant", "easy clean",
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
        ws.cell(row_idx, cols["review_reason"], "Revision R013 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R013 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        found = {c[0] for c in changed}
        raise RuntimeError(f"Did not update all handles. Missing: {sorted(set(REVISIONS) - found)}")

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    else:
        log = wb["Revision_Log"]
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R013"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R013", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R013",
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
        "next_step": "Bắt đầu revision R014 or Re-QA revision R013",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R013 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R013.",
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
        "r013_log_rows": counts.get("R013", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

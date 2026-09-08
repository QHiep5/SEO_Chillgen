from __future__ import annotations

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R020/SEO_Product_Optimization_revision_R020.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R021"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R021.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R021"
MANIFEST = RUN_DIR / "revision_R021_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R021_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-103": {
        "title": "Raven Autumn Halloween Rug with Gothic Leaf Art",
        "meta": "Decorate with a raven autumn Halloween rug featuring dark bird artwork, red maple leaves and warm moonlit colors.",
        "description": (
            "Decorate with a raven autumn Halloween rug featuring dark bird artwork, red maple leaves and warm moonlit colors.\n\n"
            "Design details\n"
            "- Round gothic fall rug artwork shows a large dark raven or crow with red and orange maple leaves.\n"
            "- Product images include seasonal room mockups, close-up views and size reference graphics.\n"
            "- SEO copy stays tied to visible raven autumn artwork without unsupported material, reverse-side or surface-performance claims."
        ),
    },
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-104": {
        "title": "Checkerboard Ghost Illusion Rug with Pumpkin Art",
        "meta": "Decorate with a checkerboard ghost illusion rug featuring floating ghosts, pumpkins, spiderwebs and vortex artwork.",
        "description": (
            "Decorate with a checkerboard ghost illusion rug featuring floating ghosts, pumpkins, spiderwebs and vortex artwork.\n\n"
            "Design details\n"
            "- Round Halloween rug artwork shows a black white checkerboard vortex with floating ghosts and orange pumpkins.\n"
            "- Visible details include purple and green border squares, spiderwebs, leaves and seasonal room mockups.\n"
            "- SEO copy focuses on visible ghost illusion artwork without unsupported material, reverse-side or surface-performance claims."
        ),
    },
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-105": {
        "title": "Ghost Campfire Halloween Rug with Autumn Lake Art",
        "meta": "Decorate with a ghost campfire Halloween rug featuring blanket ghost, glowing fire, pumpkins and autumn lake scene.",
        "description": (
            "Decorate with a ghost campfire Halloween rug featuring blanket ghost, glowing fire, pumpkins and autumn lake scene.\n\n"
            "Design details\n"
            "- Round Halloween fall rug artwork shows a blanket-wrapped ghost beside a glowing campfire with pumpkins and lanterns.\n"
            "- Visible background details include autumn leaves, lake or cabin scenery and warm sunset colors.\n"
            "- SEO copy stays tied to visible ghost campfire artwork without unsupported material, reverse-side or surface-performance claims."
        ),
    },
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-106": {
        "title": "Crawling Skeleton Pit Rug with Stone Well Art",
        "meta": "Decorate with a crawling skeleton pit rug featuring stone well illusion, glowing skull eyes and dark Halloween art.",
        "description": (
            "Decorate with a crawling skeleton pit rug featuring stone well illusion, glowing skull eyes and dark Halloween art.\n\n"
            "Design details\n"
            "- Round Halloween rug artwork shows a skeleton crawling from a dark stone pit with glowing eyes and brick rim.\n"
            "- Product images include living room, dining and fireplace scenes plus close-up and size reference panels.\n"
            "- SEO copy focuses on visible skeleton pit artwork without unsupported material, reverse-side or surface-performance claims."
        ),
    },
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-108": {
        "title": "Haunted Path Ghost Pumpkin Rug with House Art",
        "meta": "Decorate with a haunted path ghost pumpkin rug featuring haunted house, cobblestone walkway, bats and pumpkins.",
        "description": (
            "Decorate with a haunted path ghost pumpkin rug featuring haunted house, cobblestone walkway, bats and pumpkins.\n\n"
            "Design details\n"
            "- Round Halloween rug artwork shows a haunted house, cobblestone path, white ghosts and jack-o-lantern pumpkins.\n"
            "- Visible details include bats, dark trees and purple orange sky in seasonal room mockups.\n"
            "- SEO copy stays tied to visible haunted path artwork without unsupported material, reverse-side or surface-performance claims."
        ),
    },
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-109": {
        "title": "Corgi Haunted Castle Halloween Rug with Witch Hats",
        "meta": "Decorate with a corgi haunted castle Halloween rug featuring witch hat dogs, pumpkins, bats and moonlit castle.",
        "description": (
            "Decorate with a corgi haunted castle Halloween rug featuring witch hat dogs, pumpkins, bats and moonlit castle.\n\n"
            "Design details\n"
            "- Round Halloween rug artwork shows corgi dogs wearing witch hats before a haunted castle silhouette.\n"
            "- Visible details include a large moon, bats, pumpkins and dark blue night background.\n"
            "- SEO copy focuses on visible corgi Halloween artwork without unsupported material, reverse-side or surface-performance claims."
        ),
    },
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-111": {
        "title": "Purple Haunted Village Rug with Ghost Vortex Art",
        "meta": "Decorate with a purple haunted village Halloween rug featuring ghostly vortex, pumpkins, lanterns and gothic arches.",
        "description": (
            "Decorate with a purple haunted village Halloween rug featuring ghostly vortex, pumpkins, lanterns and gothic arches.\n\n"
            "Design details\n"
            "- Round purple Halloween rug artwork shows haunted village or gothic city vortex imagery with small ghosts.\n"
            "- Visible details include pumpkins, glowing lanterns, dark arches and layered depth effect.\n"
            "- SEO copy stays tied to visible haunted village artwork without unsupported material, reverse-side or surface-performance claims."
        ),
    },
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-112": {
        "title": "Witch Hat Pumpkin Halloween Rug with Web Border",
        "meta": "Decorate with a witch hat pumpkin Halloween rug featuring jack-o-lantern art, orange leaves and spiderweb border.",
        "description": (
            "Decorate with a witch hat pumpkin Halloween rug featuring jack-o-lantern art, orange leaves and spiderweb border.\n\n"
            "Design details\n"
            "- Round Halloween rug artwork shows a smiling jack-o-lantern wearing a black witch hat.\n"
            "- Visible details include orange fall leaves, web border and teal orange background in seasonal room mockups.\n"
            "- SEO copy focuses on visible pumpkin Halloween artwork without unsupported material, reverse-side or surface-performance claims."
        ),
    },
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-113": {
        "title": "Grim Reaper Halloween Rug with Cemetery Artwork",
        "meta": "Decorate with a grim reaper Halloween rug featuring hooded figure, cemetery scene, pumpkins, candles and bats.",
        "description": (
            "Decorate with a grim reaper Halloween rug featuring hooded figure, cemetery scene, pumpkins, candles and bats.\n\n"
            "Design details\n"
            "- Round rug artwork shows a hooded grim reaper holding a candle in a misty cemetery scene.\n"
            "- Visible details include tombstones, bats, pumpkins and blue gray background in room mockups.\n"
            "- SEO copy stays tied to visible grim reaper artwork without unsupported material, reverse-side or surface-performance claims."
        ),
    },
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-114": {
        "title": "Haunted House Halloween Rug with Twilight Artwork",
        "meta": "Decorate with a haunted house Halloween rug featuring glowing windows, pumpkin walkway, bats and purple twilight sky.",
        "description": (
            "Decorate with a haunted house Halloween rug featuring glowing windows, pumpkin walkway, bats and purple twilight sky.\n\n"
            "Design details\n"
            "- Round rug artwork shows a gothic haunted house with glowing orange windows and pumpkin-lined walkway.\n"
            "- Visible details include bare trees, bats and purple twilight sky in seasonal room mockups.\n"
            "- SEO copy focuses on visible haunted house artwork without unsupported material, reverse-side or surface-performance claims."
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
    "reinforced", "absorbent", "absorption", "water absorption",
    "support ", "supports ", "use a ", "foldable material",
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
        ws.cell(row_idx, cols["review_reason"], "Revision R021 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R021 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        found = {c[0] for c in changed}
        raise RuntimeError(f"Did not update all handles. Missing: {sorted(set(REVISIONS) - found)}")

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    else:
        log = wb["Revision_Log"]
        keep = [r for r in log.iter_rows(min_row=2, values_only=True) if r and r[0] != "R021"]
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R021", NOW, handle, row_idx, old_revision, new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R021",
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
        "next_step": "Bắt đầu revision R022 or Re-QA revision R021",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R021 Summary", "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R021.",
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
        "r021_log_rows": counts.get("R021", 0),
        "duplicate_meta_titles": len(dup_meta_titles),
        "sheets": reopened.sheetnames,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

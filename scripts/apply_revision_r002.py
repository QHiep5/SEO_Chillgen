from __future__ import annotations

import json
from datetime import datetime, timezone, timedelta
from pathlib import Path

from openpyxl import load_workbook


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R001/SEO_Product_Optimization_revision_R001.xlsx"
OUTPUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R002"
OUTPUT = OUTPUT_DIR / "SEO_Product_Optimization_revision_R002.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R002"
MANIFEST = RUN_DIR / "revision_R002_manifest.json"
SUMMARY = OUTPUT_DIR / "REVISION_R002_SUMMARY.md"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"

TZ = timezone(timedelta(hours=7))
NOW = datetime.now(TZ).replace(microsecond=0).isoformat()


REVISIONS = {
    "custom-japanese-koi-fish-shaped-area-rug-personalized-3d-design-08": {
        "title": "Swirl Water Japanese Koi Pond Area Rug for Room Decor",
        "meta": "Decorate with a swirl water Japanese koi pond area rug featuring orange koi, lily pads, dark teal water and room mockups.",
        "description": (
            "Decorate with a swirl water Japanese koi pond area rug featuring orange koi, lily pads, dark teal water and room mockups.\n\n"
            "Design details\n"
            "- Irregular pond-shaped rug with dark teal swirling water, clustered green lily pads and orange-and-white koi fish.\n"
            "- Gallery images show the rug in indoor room scenes, plus close-up, size and design-detail visuals.\n"
            "- SEO copy stays focused on the visible koi pond artwork and avoids unsupported material, backing or surface-performance claims."
        ),
    },
    "custom-japanese-koi-fish-shaped-area-rug-personalized-3d-optical": {
        "title": "3D Optical Japanese Koi Pond Rug with Stone Border",
        "meta": "Add a 3D optical Japanese koi pond rug with swimming koi, green-blue water, stone edging and pink flower details.",
        "description": (
            "Add a 3D optical Japanese koi pond rug with swimming koi, green-blue water, stone edging and pink flower details.\n\n"
            "Design details\n"
            "- Pond-shaped rug with a light stone border, multiple red and yellow koi fish and small pink flower accents.\n"
            "- Gallery images include room mockups, close-up design views, size chart and lifestyle scenes.\n"
            "- SEO copy stays tied to the visible Japanese koi pond motif without adding unsupported backing, material or surface-performance claims."
        ),
    },
    "boat-rug-04-boat-rug-04": {
        "title": "Personalized Navy Compass Boat Rug with Welcome Aboard Text",
        "meta": "Personalize a navy compass boat rug with welcome aboard wording, boat name, location text and striped nautical styling.",
        "description": (
            "Personalize a navy compass boat rug with welcome aboard wording, boat name, location text and striped nautical styling.\n\n"
            "Design details\n"
            "- Rectangular navy and white nautical rug shown in a living room setting with a compass rose and framed border.\n"
            "- Sample artwork includes Welcome Aboard text with boat name and location fields visible in the design.\n"
            "- SEO copy focuses on the personalized boat-name and compass artwork without adding unsupported material or surface-performance claims."
        ),
    },
    "core-vocabulary-rug-communication-rugs-for-kids-sped-cla-design-102": {
        "title": "Core Vocabulary AAC Classroom Rug with Icon Word Grid",
        "meta": "Use a core vocabulary AAC classroom rug with icon-word cells for yes, no, feelings, help, stop, go and classroom prompts.",
        "description": (
            "Use a core vocabulary AAC classroom rug with icon-word cells for yes, no, feelings, help, stop, go and classroom prompts.\n\n"
            "Design details\n"
            "- Light classroom rug with a grid of simple icons and words for communication prompts, emotions, actions and questions.\n"
            "- Gallery shows classroom, playroom and home-style scenes with the communication-board artwork clearly visible.\n"
            "- SEO copy describes the visible AAC-style vocabulary layout and avoids unsupported therapeutic or performance claims."
        ),
    },
    "wheel-of-feelings-and-emotions-round-rug-educational-mental-health": {
        "title": "Round Feelings Wheel Classroom Rug with Emotion Faces",
        "meta": "Use a round feelings wheel classroom rug with How Are You Feeling text, emotion faces, flower center and colorful labels.",
        "description": (
            "Use a round feelings wheel classroom rug with How Are You Feeling text, emotion faces, flower center and colorful labels.\n\n"
            "Design details\n"
            "- Round classroom rug with a smiling flower center, colorful emotion faces and readable feeling labels around the ring.\n"
            "- Gallery images show group-time classroom scenes, close-up design views, room mockups and size visuals.\n"
            "- SEO copy stays focused on the visible feelings-wheel artwork without adding unsupported backing, safety or material claims."
        ),
    },
    "wheel-of-feelings-and-emotions-round-rug-educational-men-design-100": {
        "title": "Round Learn Shapes Classroom Rug with Colorful Shape Labels",
        "meta": "Use a round learn shapes classroom rug with bright labeled panels for circle, square, triangle, rectangle, star and more.",
        "description": (
            "Use a round learn shapes classroom rug with bright labeled panels for circle, square, triangle, rectangle, star and more.\n\n"
            "Design details\n"
            "- Round educational rug with Learn Shapes center text and colorful shape panels including circle, square, triangle, oval, star and heart.\n"
            "- Gallery images show classroom scenes, child reading visuals, size guide and close-up design panels.\n"
            "- SEO copy describes the visible early-learning shape design and removes unsupported backing, safety or material claims."
        ),
    },
    "wheel-of-feelings-and-emotions-round-rug-educational-men-design-101": {
        "title": "Emoji Feelings Wheel Classroom Rug with Emotion Labels",
        "meta": "Use an emoji feelings wheel classroom rug with How Are You Feeling text, colorful faces and emotion labels.",
        "description": (
            "Use an emoji feelings wheel classroom rug with How Are You Feeling text, colorful faces and emotion labels.\n\n"
            "Design details\n"
            "- Dark blue round rug with a large emoji face center, colorful emotion faces and readable labels around the circular design.\n"
            "- Gallery includes classroom group scenes, close-up design views, room mockups, size guide and lifestyle visuals.\n"
            "- SEO copy stays tied to the visible emoji feelings-wheel artwork without unsupported backing, safety or material claims."
        ),
    },
    "custom-composition-notebook-classroom-rug-for-kids-d5b3be8848-d5b3be8848": {
        "title": "Personalized Pink Classroom Rug with Teacher Name Doodles",
        "meta": "Personalize a pink classroom rug with teacher name text, I Am prompt, school doodles, rainbow, pencils and apple artwork.",
        "description": (
            "Personalize a pink classroom rug with teacher name text, I Am prompt, school doodles, rainbow, pencils and apple artwork.\n\n"
            "Design details\n"
            "- Pink classroom rug sample shows Mrs. Brown's Class, I Am wording, school icons, rainbow, pencils, apple, hearts and doodle border.\n"
            "- Gallery includes classroom/playroom scenes, children photos, close-up panels and feature-style visuals.\n"
            "- SEO copy focuses on the visible custom teacher-name classroom design and avoids unsupported material or performance claims."
        ),
    },
    "custom-composition-notebook-classroom-rug-kids-f215fce7cb": {
        "title": "Personalized Pink Doodle Teacher Rug with School Icons",
        "meta": "Personalize a pink doodle teacher rug with Mrs. Smith name text, apple, daisies, pencils, stars and classroom scenes.",
        "description": (
            "Personalize a pink doodle teacher rug with Mrs. Smith name text, apple, daisies, pencils, stars and classroom scenes.\n\n"
            "Design details\n"
            "- Pink and peach classroom rug sample with large Mrs. Smith text, apple, daisies, pencils, paper plane, stars and black doodle lines.\n"
            "- Gallery images show classroom scenes, playroom-style mockups, children photos and close-up feature panels.\n"
            "- SEO copy stays tied to the visible custom teacher-name artwork without adding unsupported material, backing or surface-performance claims."
        ),
    },
    "tree-of-life-celtic-area-rug-d16ac94729": {
        "title": "Celtic Tree of Life Area Rug with Knotwork Roots",
        "meta": "Decorate with a Celtic Tree of Life area rug showing golden knotwork roots, a round tree design and earthy room mockups.",
        "description": (
            "Decorate with a Celtic Tree of Life area rug showing golden knotwork roots, a round tree design and earthy room mockups.\n\n"
            "Design details\n"
            "- Earthy Tree of Life area rug with golden Celtic knot roots forming a circle, large tree canopy and dark green-brown palette.\n"
            "- Gallery images show room scenes, detail panels, texture close-ups and size/mockup visuals.\n"
            "- SEO copy focuses on the visible Celtic knotwork and tree artwork without unsupported material, backing or surface-performance claims."
        ),
    },
}


def headers(ws):
    return {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(SOURCE)
    ws = wb["SEO_Products"]
    cols = headers(ws)

    required_cols = [
        "Handle",
        "title_proposed",
        "meta_title_seo",
        "meta_title_chars",
        "meta_description_seo",
        "meta_description_chars",
        "description_proposed",
        "description_proposed_html",
        "revision",
        "review_status",
        "review_reason",
        "processing_status",
        "content_qa_status",
        "issues",
    ]
    missing = [c for c in required_cols if c not in cols]
    if missing:
        raise RuntimeError(f"Missing columns: {missing}")

    changed = []
    for row_idx in range(2, ws.max_row + 1):
        handle = ws.cell(row_idx, cols["Handle"]).value
        if handle not in REVISIONS:
            continue
        item = REVISIONS[handle]
        title = item["title"]
        meta = item["meta"]
        desc = item["description"]
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
        ws.cell(row_idx, cols["review_reason"], "Revision R002 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row_idx, cols["processing_status"], "DRAFTED")
        ws.cell(row_idx, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row_idx, cols["issues"], "Revision R002 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")
        changed.append((handle, row_idx, old_revision, "2", title, meta))

    if len(changed) != len(REVISIONS):
        found = {c[0] for c in changed}
        missing = sorted(set(REVISIONS) - found)
        raise RuntimeError(f"Did not update all handles. Missing: {missing}")

    # Sanitize internal research/mapping phrases for scoped products: replace broad risky
    # claim wording in internal sheets without changing unrelated products.
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
                new = (
                    value.replace("indoor/outdoor", "indoor")
                    .replace("Indoor/outdoor", "Indoor")
                    .replace("outdoor-use", "room-use")
                    .replace("outdoor use", "room use")
                )
                sh.cell(row_idx, col_idx, new)

    if "Revision_Log" not in wb.sheetnames:
        log = wb.create_sheet("Revision_Log")
        log.append(["revision_batch_id", "generated_at", "handle", "source_row", "old_revision", "new_revision", "changed_fields", "qa_issue_fields", "recheck_condition"])
    else:
        log = wb["Revision_Log"]
        # Avoid double-appending if rerun.
        keep = []
        for r in log.iter_rows(min_row=2, values_only=True):
            if r and r[0] != "R002":
                keep.append(r)
        if log.max_row > 1:
            log.delete_rows(2, log.max_row - 1)
        for r in keep:
            log.append(list(r))

    for handle, row_idx, old_revision, new_revision, title, meta in changed:
        log.append([
            "R002",
            NOW,
            handle,
            row_idx,
            old_revision,
            new_revision,
            "title_proposed, meta_title_seo, meta_title_chars, meta_description_seo, meta_description_chars, description_proposed, description_proposed_html, revision, review_reason, issues",
            "T1; T2; unsupported_claims",
            "Run evidence-driven re-QA on revision 2 for this product before approval.",
        ])

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R002",
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
        "next_step": "Re-QA revision R002",
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")

    lines = [
        "# Revision R002 Summary",
        "",
        f"- Generated at: `{NOW}`",
        f"- Source workbook: `{SOURCE}`",
        f"- Output workbook: `{OUTPUT}`",
        "- Scope: 10 products from revision plan R002.",
        "- Fixes targeted QA issues: short T1/T2 and unsupported claims.",
        "- Kept `review_status=NEEDS_REVIEW` and `content_qa_status=NOT_RUN`.",
        "- No Shopify deploy, no approval.",
        "",
        "## Updated products",
        "",
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
        for forbidden in ["indoor/outdoor", " outdoor ", "anti-slip", "non-slip", "kid friendly", "pet friendly", "safe durable"]:
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

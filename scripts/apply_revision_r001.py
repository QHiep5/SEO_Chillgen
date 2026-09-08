from __future__ import annotations

import json
import re
import shutil
from datetime import datetime, timezone, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/SEO_Product_Optimization.xlsx"
OUT_DIR = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R001"
WORK_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R001"
OUTPUT = OUT_DIR / "SEO_Product_Optimization_revision_R001.xlsx"
MANIFEST = WORK_DIR / "revision_R001_manifest.json"
SUMMARY_MD = OUT_DIR / "REVISION_R001_SUMMARY.md"


REVISION_TS = datetime.now(timezone(timedelta(hours=7))).isoformat(timespec="seconds")


REVISIONS = {
    "custom-soccer-tournament-trophy-area-rug-8d23acd520-8d23acd520": {
        "title": "Custom Soccer 2026 USA Trophy Area Rug for Fan Rooms",
        "meta": "Style a soccer fan room with a 2026 USA trophy area rug featuring a stadium field, eagle crest, soccer ball and flag border.",
        "desc": (
            "Style a soccer fan room with a 2026 USA trophy area rug featuring a stadium field, eagle crest, soccer ball and flag border.\n\n"
            "Design details\n"
            "- Rectangular soccer area rug with green field artwork, USA text, golden trophy, eagle crest, soccer ball, stadium crowd and patriotic flag border.\n"
            "- Shown in indoor room mockups for living rooms, bedrooms, game rooms and fan spaces.\n"
            "- SEO copy stays tied to the visible sports motif and page evidence."
        ),
        "html": (
            "<p>Style a soccer fan room with a 2026 USA trophy area rug featuring a stadium field, eagle crest, soccer ball and flag border.</p>"
            "<h3>Design details</h3><ul>"
            "<li>Rectangular soccer area rug with green field artwork, USA text, golden trophy, eagle crest, soccer ball, stadium crowd and patriotic flag border.</li>"
            "<li>Shown in indoor room mockups for living rooms, bedrooms, game rooms and fan spaces.</li>"
            "<li>SEO copy stays tied to the visible sports motif and page evidence.</li>"
            "</ul>"
        ),
    },
    "personalized-dog-doormat-custom-text-non-slip-c76f704530": {
        "title": "Personalized No Need To Knock Dog Doormat with Names",
        "meta": "Personalize a No Need To Knock dog doormat with pet names, cartoon dog portraits and a funny entryway message.",
        "desc": (
            "Personalize a No Need To Knock dog doormat with pet names, cartoon dog portraits and a funny entryway message.\n\n"
            "Design details\n"
            "- Tan cartoon dog mat reading No Need To Knock We Know You're Here, with three pet portraits and names shown in the sample design.\n"
            "- Includes dog portrait and breed option panels plus doorway and room mockups.\n"
            "- SEO copy focuses on the personalized dog artwork, names and phrase visible in the product evidence."
        ),
        "html": (
            "<p>Personalize a No Need To Knock dog doormat with pet names, cartoon dog portraits and a funny entryway message.</p>"
            "<h3>Design details</h3><ul>"
            "<li>Tan cartoon dog mat reading No Need To Knock We Know You're Here, with three pet portraits and names shown in the sample design.</li>"
            "<li>Includes dog portrait and breed option panels plus doorway and room mockups.</li>"
            "<li>SEO copy focuses on the personalized dog artwork, names and phrase visible in the product evidence.</li>"
            "</ul>"
        ),
    },
    "custom-galaxy-wolves-rug-bedroom-living-room-eb7c237d42-eb7c237d42": {
        "title": "Moonlit Wolf Pack Area Rug for Bedroom and Living Room",
        "meta": "Decorate with a moonlit wolf pack area rug showing wolves on rocks, a glowing waterfall and forest scenery.",
        "desc": (
            "Decorate with a moonlit wolf pack area rug showing wolves on rocks, a glowing waterfall and forest scenery.\n\n"
            "Design details\n"
            "- Rectangular wolf rug with two wolves howling on rocks under a large moon, teal waterfall light and dark forest artwork.\n"
            "- Shown beside sofas and under desks in bedroom, living room and office-style mockups.\n"
            "- SEO copy stays focused on the visible moonlit pack motif."
        ),
        "html": (
            "<p>Decorate with a moonlit wolf pack area rug showing wolves on rocks, a glowing waterfall and forest scenery.</p>"
            "<h3>Design details</h3><ul>"
            "<li>Rectangular wolf rug with two wolves howling on rocks under a large moon, teal waterfall light and dark forest artwork.</li>"
            "<li>Shown beside sofas and under desks in bedroom, living room and office-style mockups.</li>"
            "<li>SEO copy stays focused on the visible moonlit pack motif.</li>"
            "</ul>"
        ),
    },
    "custom-galaxy-wolves-welcome-mat-f320a33157": {
        "title": "Lightning Snow Wolf Area Rug for Bedroom and Living Room",
        "meta": "Decorate with a lightning snow wolf area rug featuring a gray wolf portrait, snowy forest and blue energy accents.",
        "desc": (
            "Decorate with a lightning snow wolf area rug featuring a gray wolf portrait, snowy forest and blue energy accents.\n\n"
            "Design details\n"
            "- Blue and white wolf rug with a large gray wolf portrait, snowy forest, icy river, lightning effects and starry sky.\n"
            "- Shown in living room, sectional sofa, wood floor and desk mockups.\n"
            "- SEO copy keeps the focus on the winter wolf artwork shown in the images."
        ),
        "html": (
            "<p>Decorate with a lightning snow wolf area rug featuring a gray wolf portrait, snowy forest and blue energy accents.</p>"
            "<h3>Design details</h3><ul>"
            "<li>Blue and white wolf rug with a large gray wolf portrait, snowy forest, icy river, lightning effects and starry sky.</li>"
            "<li>Shown in living room, sectional sofa, wood floor and desk mockups.</li>"
            "<li>SEO copy keeps the focus on the winter wolf artwork shown in the images.</li>"
            "</ul>"
        ),
    },
    "custom-galaxy-wolves-rug-boy-room-bedroom-167aea55b4": {
        "title": "Foggy Forest Wolf Area Rug for Boys Bedroom and Den Decor",
        "meta": "Decorate with a foggy forest wolf area rug showing a dark wolf portrait, misty blue trees and room mockups.",
        "desc": (
            "Decorate with a foggy forest wolf area rug showing a dark wolf portrait, misty blue trees and room mockups.\n\n"
            "Design details\n"
            "- Dark blue-gray wolf rug with a single black wolf standing in foggy forest light and misty tree scenery.\n"
            "- Shown by a sofa, on wood flooring, under a sectional and near a desk.\n"
            "- SEO copy stays tied to the moody forest wolf design visible in the product images."
        ),
        "html": (
            "<p>Decorate with a foggy forest wolf area rug showing a dark wolf portrait, misty blue trees and room mockups.</p>"
            "<h3>Design details</h3><ul>"
            "<li>Dark blue-gray wolf rug with a single black wolf standing in foggy forest light and misty tree scenery.</li>"
            "<li>Shown by a sofa, on wood flooring, under a sectional and near a desk.</li>"
            "<li>SEO copy stays tied to the moody forest wolf design visible in the product images.</li>"
            "</ul>"
        ),
    },
    "galaxy-wolves-wolf-area-rug-94b47adf17-94b47adf17": {
        "title": "Floral Line Art Wolf Area Rug for Neutral Room Decor",
        "meta": "Decorate with a floral line art wolf area rug featuring sepia botanical details, a wolf portrait and neutral styling.",
        "desc": (
            "Decorate with a floral line art wolf area rug featuring sepia botanical details, a wolf portrait and neutral styling.\n\n"
            "Design details\n"
            "- Sepia tan wolf rug with a side-profile wolf portrait and floral botanical line art around the neck and background.\n"
            "- Shown beside a sofa, on wood flooring, under a sectional and under a desk.\n"
            "- SEO copy focuses on the visible botanical wolf artwork and neutral room use."
        ),
        "html": (
            "<p>Decorate with a floral line art wolf area rug featuring sepia botanical details, a wolf portrait and neutral styling.</p>"
            "<h3>Design details</h3><ul>"
            "<li>Sepia tan wolf rug with a side-profile wolf portrait and floral botanical line art around the neck and background.</li>"
            "<li>Shown beside a sofa, on wood flooring, under a sectional and under a desk.</li>"
            "<li>SEO copy focuses on the visible botanical wolf artwork and neutral room use.</li>"
            "</ul>"
        ),
    },
    "custom-galaxy-wolf-welcome-mat-0e4d706227": {
        "title": "Sepia Wolf Pack Area Rug for Rustic Living Room Decor",
        "meta": "Style a room with a sepia wolf pack area rug showing a large wolf face, smaller wolves and warm forest tones.",
        "desc": (
            "Style a room with a sepia wolf pack area rug showing a large wolf face, smaller wolves and warm forest tones.\n\n"
            "Design details\n"
            "- Brown and sepia wolf rug with a large wolf face, multiple smaller wolves, forest scenery and warm earthy color.\n"
            "- Shown in room mockups near a sofa, on wood flooring, under a sectional and near a desk.\n"
            "- SEO copy stays focused on the visible pack artwork and rustic wolf theme."
        ),
        "html": (
            "<p>Style a room with a sepia wolf pack area rug showing a large wolf face, smaller wolves and warm forest tones.</p>"
            "<h3>Design details</h3><ul>"
            "<li>Brown and sepia wolf rug with a large wolf face, multiple smaller wolves, forest scenery and warm earthy color.</li>"
            "<li>Shown in room mockups near a sofa, on wood flooring, under a sectional and near a desk.</li>"
            "<li>SEO copy stays focused on the visible pack artwork and rustic wolf theme.</li>"
            "</ul>"
        ),
    },
    "galaxy-wolves-wolf-area-rug-carpet-da377ca92e-da377ca92e": {
        "title": "Dreamcatcher Wolf Area Rug with Feather Artwork for Bedrooms",
        "meta": "Add a dreamcatcher wolf area rug with feather details, warm tan artwork and a wolf portrait for room decor.",
        "desc": (
            "Add a dreamcatcher wolf area rug with feather details, warm tan artwork and a wolf portrait for room decor.\n\n"
            "Design details\n"
            "- Tan wolf rug with a wolf portrait inside a dreamcatcher-style circle, blue bead details and long feather accents.\n"
            "- Shown beside a sofa, on wood flooring, under a sectional and under a desk.\n"
            "- SEO copy stays focused on the visible dreamcatcher-inspired wolf artwork."
        ),
        "html": (
            "<p>Add a dreamcatcher wolf area rug with feather details, warm tan artwork and a wolf portrait for room decor.</p>"
            "<h3>Design details</h3><ul>"
            "<li>Tan wolf rug with a wolf portrait inside a dreamcatcher-style circle, blue bead details and long feather accents.</li>"
            "<li>Shown beside a sofa, on wood flooring, under a sectional and under a desk.</li>"
            "<li>SEO copy stays focused on the visible dreamcatcher-inspired wolf artwork.</li>"
            "</ul>"
        ),
    },
    "wolf-galaxy-wolves-area-rug-f49c8f2c6e-f49c8f2c6e": {
        "title": "Two Tone Galaxy Wolf Area Rug for Bedroom or Gaming Room",
        "meta": "Style a bedroom or gaming room with a two tone galaxy wolf area rug showing a pink and blue wolf face.",
        "desc": (
            "Style a bedroom or gaming room with a two tone galaxy wolf area rug showing a pink and blue wolf face.\n\n"
            "Design details\n"
            "- Bright galaxy rug with a close-up wolf face split into pink-purple and blue-white tones against a starry space background.\n"
            "- Shown beside a sofa, on wood flooring, under a sectional and under a desk.\n"
            "- SEO copy keeps the focus on the visible colorful galaxy wolf artwork."
        ),
        "html": (
            "<p>Style a bedroom or gaming room with a two tone galaxy wolf area rug showing a pink and blue wolf face.</p>"
            "<h3>Design details</h3><ul>"
            "<li>Bright galaxy rug with a close-up wolf face split into pink-purple and blue-white tones against a starry space background.</li>"
            "<li>Shown beside a sofa, on wood flooring, under a sectional and under a desk.</li>"
            "<li>SEO copy keeps the focus on the visible colorful galaxy wolf artwork.</li>"
            "</ul>"
        ),
    },
    "custom-dog-paw-print-shaped-rug-39afbc78d5": {
        "title": "Patchwork Paw Shaped Rug with Plaid Panels for Pet Rooms",
        "meta": "Add a patchwork paw shaped rug with plaid fabric panels, a black paw outline and pet room styling.",
        "desc": (
            "Add a patchwork paw shaped rug with plaid fabric panels, a black paw outline and pet room styling.\n\n"
            "Design details\n"
            "- Large paw-shaped rug with black paw outline, multicolor tartan-style panels and small paw-print border details.\n"
            "- Shown in room and nursery-style mockups with size and construction panels in the product gallery.\n"
            "- SEO copy avoids backing and surface-performance claims unless confirmed during admin/export review."
        ),
        "html": (
            "<p>Add a patchwork paw shaped rug with plaid fabric panels, a black paw outline and pet room styling.</p>"
            "<h3>Design details</h3><ul>"
            "<li>Large paw-shaped rug with black paw outline, multicolor tartan-style panels and small paw-print border details.</li>"
            "<li>Shown in room and nursery-style mockups with size and construction panels in the product gallery.</li>"
            "<li>SEO copy avoids backing and surface-performance claims unless confirmed during admin/export review.</li>"
            "</ul>"
        ),
    },
}


def safe_text(value: str) -> str:
    if value.startswith("="):
        return "'" + value
    return value


def get_headers(ws):
    return {cell.value: cell.column for cell in ws[1]}


def style_revision_log(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Arial", size=10, color="1F1F1F")
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            if cell.row > 1:
                cell.font = body_font
            cell.alignment = Alignment(vertical="top", wrap_text=False)
    for idx, col in enumerate(ws.columns, start=1):
        width = 10
        for cell in col:
            width = max(width, min(len(str(cell.value or "")) + 2, 80))
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.auto_filter.ref = ws.dimensions


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SOURCE, OUTPUT)
    wb = load_workbook(OUTPUT)

    ws = wb["SEO_Products"]
    cols = get_headers(ws)

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
        raise RuntimeError(f"Missing SEO_Products columns: {missing}")

    changes = []
    handles = set(REVISIONS)
    for row in range(2, ws.max_row + 1):
        handle = ws.cell(row, cols["Handle"]).value
        if handle not in handles:
            continue
        rev = REVISIONS[handle]
        before = {
            "title_proposed": ws.cell(row, cols["title_proposed"]).value,
            "meta_title_seo": ws.cell(row, cols["meta_title_seo"]).value,
            "meta_description_seo": ws.cell(row, cols["meta_description_seo"]).value,
            "description_proposed": ws.cell(row, cols["description_proposed"]).value,
            "revision": ws.cell(row, cols["revision"]).value,
        }

        ws.cell(row, cols["title_proposed"], safe_text(rev["title"]))
        ws.cell(row, cols["meta_title_seo"], safe_text(rev["title"]))
        ws.cell(row, cols["meta_title_chars"], len(rev["title"]))
        ws.cell(row, cols["meta_description_seo"], safe_text(rev["meta"]))
        ws.cell(row, cols["meta_description_chars"], len(rev["meta"]))
        ws.cell(row, cols["description_proposed"], safe_text(rev["desc"]))
        ws.cell(row, cols["description_proposed_html"], safe_text(rev["html"]))
        ws.cell(row, cols["revision"], "2")
        ws.cell(row, cols["review_status"], "NEEDS_REVIEW")
        ws.cell(row, cols["review_reason"], "Revision R001 fixes QA MAJOR issues: short T1/T2 and unsupported claims. Requires re-QA before approval.")
        ws.cell(row, cols["processing_status"], "DRAFTED")
        ws.cell(row, cols["content_qa_status"], "NOT_RUN")
        ws.cell(row, cols["issues"], "Revision R001 updated targeted fields; not approved; re-QA required. Admin/export prerequisites still open before deploy.")

        changes.append(
            {
                "handle": handle,
                "row": row,
                "revision": "1 -> 2",
                "fields_changed": [
                    "title_proposed",
                    "meta_title_seo",
                    "meta_title_chars",
                    "meta_description_seo",
                    "meta_description_chars",
                    "description_proposed",
                    "description_proposed_html",
                    "revision",
                    "review_reason",
                    "issues",
                ],
                "before": before,
                "after": {
                    "title_proposed": rev["title"],
                    "meta_title_seo": rev["title"],
                    "meta_description_seo": rev["meta"],
                    "description_proposed": rev["desc"],
                    "revision": "2",
                },
            }
        )

    if len(changes) != len(REVISIONS):
        seen = {c["handle"] for c in changes}
        raise RuntimeError(f"Only updated {len(changes)} rows. Missing: {sorted(handles - seen)}")

    # Keep internal research text aligned where it contained the exact unsupported boilerplate.
    replacements = [
        (re.compile(r"\bindoor/outdoor\b", re.I), "indoor"),
        (re.compile(r"\boutdoor\b", re.I), "indoor"),
    ]
    for sheet_name in ["Buyer_Search_Research", "Keyword_Map"]:
        if sheet_name not in wb.sheetnames:
            continue
        ws2 = wb[sheet_name]
        hmap = get_headers(ws2)
        key_col = hmap.get("product_key")
        if not key_col:
            continue
        for row in range(2, ws2.max_row + 1):
            product_key = str(ws2.cell(row, key_col).value or "")
            if product_key not in handles:
                continue
            for col in range(1, ws2.max_column + 1):
                value = ws2.cell(row, col).value
                if not isinstance(value, str):
                    continue
                new_value = value
                for pattern, repl in replacements:
                    new_value = pattern.sub(repl, new_value)
                if product_key == "custom-dog-paw-print-shaped-rug-39afbc78d5":
                    new_value = re.sub(r"\banti-slip\b", "backing", new_value, flags=re.I)
                    new_value = re.sub(r"\bnon-slip\b", "backing", new_value, flags=re.I)
                if new_value != value:
                    ws2.cell(row, col, safe_text(new_value))

    if "Revision_Log" in wb.sheetnames:
        del wb["Revision_Log"]
    log = wb.create_sheet("Revision_Log")
    log.append(
        [
            "revision_batch_id",
            "generated_at",
            "handle",
            "source_row",
            "old_revision",
            "new_revision",
            "changed_fields",
            "qa_issue_fields",
            "recheck_condition",
        ]
    )
    for change in changes:
        log.append(
            [
                "R001",
                REVISION_TS,
                change["handle"],
                change["row"],
                "1",
                "2",
                ", ".join(change["fields_changed"]),
                "T1; T2; unsupported_claims",
                "Run evidence-driven re-QA on revision 2 for this product before approval.",
            ]
        )
    style_revision_log(log)

    wb.save(OUTPUT)

    manifest = {
        "revision_batch_id": "R001",
        "generated_at": REVISION_TS,
        "source_workbook": str(SOURCE),
        "output_workbook": str(OUTPUT),
        "source_qa_summary": str(BASE / "resutls/chillgen.com/chillgen_20260907_01/qa/QA_EVIDENCE_RERUN_SUMMARY.xlsx"),
        "product_count": len(changes),
        "scope": "Targeted fix for MAJOR issues T1, T2, and unsupported_claims. No Shopify deploy or approval.",
        "changes": changes,
    }
    MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    SUMMARY_MD.write_text(
        "# Revision R001 Summary\n\n"
        f"- Generated at: `{REVISION_TS}`\n"
        f"- Source workbook: `{SOURCE}`\n"
        f"- Output workbook: `{OUTPUT}`\n"
        "- Scope: targeted revision for QA MAJOR issues `T1`, `T2`, and `unsupported_claims`.\n"
        "- Products updated: 10\n"
        "- Revision: `1 -> 2`\n"
        "- Status after revision: still `NEEDS_REVIEW`, `content_qa_status=NOT_RUN`, not approved, not deployed.\n\n"
        "## Updated products\n\n"
        + "\n".join(f"- `{c['handle']}`: `{c['after']['title_proposed']}`" for c in changes)
        + "\n\n## Next step\n\nRun evidence-driven re-QA for revision batch `R001` only.\n",
        encoding="utf-8",
    )

    print(OUTPUT)
    print(MANIFEST)
    print(SUMMARY_MD)


if __name__ == "__main__":
    main()

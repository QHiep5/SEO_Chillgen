from pathlib import Path
import csv
import hashlib
import json

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path.cwd()
RUN_ID = "chillgen_20260907_01"
SHOP = "chillgen.com"
BATCH = "B033"
REVISION = "R109"
RUN = BASE / f"seo_runs/{SHOP}/{RUN_ID}"
OUT_DIR = BASE / f"resutls/{SHOP}/{RUN_ID}/revisions/{REVISION}"
OUT = OUT_DIR / f"SEO_Product_Optimization_revision_{REVISION}.xlsx"

COPY = {
    "personalized-soccer-area-rug-trophy-football-868088cbf4-868088cbf4": ("Game Of Champions Soccer Rug", "Game Of Champions soccer rug", "soccer 2026 area rug, championship football rug, soccer fan room rug", "a Game Of Champions soccer and football design with a ball, trophy, and championship details for a fan room or game space"),
    "custom-soccer-tournament-trophy-area-rug-8d23acd520-8d23acd520": ("USA World Game 2026 Rug", "USA World Game 2026 rug", "custom soccer 2026 area rug, USA soccer rug, tournament trophy rug", "a USA World Game 2026 design with a soccer ball, trophy, and patriotic details for a fan room or game space"),
    "personalized-soccer-2026-area-rug-95231047da-95231047da": ("Host Flags Swirl Soccer Rug", "host flags swirl soccer rug", "soccer 2026 area rug, USA Canada Mexico soccer rug, football fan room rug", "a host flags swirl soccer design with USA, Canada, and Mexico colors for a fan room or game space"),
    "custom-soccer-round-rug-trophy-championship-043ff71141": ("World Game 2026 Round Rug", "World Game 2026 round rug", "soccer trophy round rug, championship soccer rug, round football fan rug", "a round World Game 2026 soccer design with a trophy and tournament details for a fan room or game space"),
    "custom-soccer-tournament-round-rug-a7f6df2bd9-a7f6df2bd9": ("World Flags Trophy Round Rug", "world flags trophy round rug", "personalized soccer trophy round rug, tournament soccer rug, football playroom rug", "a round soccer trophy design with world flags and tournament details for a fan room or game space"),
    "personalized-dog-welcome-mat-front-door-36fdce22f9": ("Humans Live Here Pet Mat", "Humans Live Here pet mat", "personalized dog welcome mat, custom pet doormat, dog cat entryway rug", "a Humans Live Here message with pet-themed artwork for a front door, entryway, or pet-loving home"),
    "personalized-dog-welcome-mat-custom-pet-photo-7a71ecf37d": ("Approved By Dog Photo Mat", "Approved By Dog photo mat", "custom pet photo doormat, personalized dog welcome mat, dog name entryway rug", "an Approved By Dog message with a custom pet photo concept for a front door, entryway, or pet-loving home"),
    "personalized-dog-welcome-mat-entryway-front-door-5385d11bb7": ("Three Pet Welcome Mat", "three pet welcome mat", "custom dog cat welcome mat, personalized pet doormat, multi pet entryway rug", "a Three Pet welcome design for a front door, entryway, or home with multiple pets"),
    "personalized-dog-welcome-doormat-front-door-entryway-d819eb0e7c": ("Loki's House Dog Mat", "Loki's House dog mat", "personalized single pet welcome mat, custom dog doormat, dog house entryway rug", "a Loki's House message with a dog-themed welcome design for a front door or entryway"),
    "personalized-dog-welcome-doormat-front-door-entryway-e2ef212662-e2ef212662": ("Four Pet Welcome Home Mat", "four pet welcome home mat", "custom multi pet doormat, personalized dog cat mat, pet lover entryway rug", "a Four Pet Welcome Home design for a front door, entryway, or home with multiple pets"),
}


def read_csv(path):
    with path.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return list(reader), reader.fieldnames or []


def sha256(path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def write_sheet(wb, name, rows, headers):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = fill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            widths[cell.column] = max(widths.get(cell.column, 0), min(70, len(str(cell.value or "")) + 2))
    for idx, width in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = max(12, width)


def main():
    seo_rows, seo_headers = read_csv(RUN / f"batches/{BATCH}_SEO_Products.csv")
    img_rows, img_headers = read_csv(RUN / f"batches/{BATCH}_Image_Audit.csv")
    evidence_rows, evidence_headers = read_csv(RUN / f"batches/{BATCH}_Product_Evidence.csv")
    keyword_rows, keyword_headers = read_csv(RUN / "keyword_research.csv")
    with (RUN / "buyer_search_research.jsonl").open(encoding="utf-8") as f:
        buyer_all = [json.loads(line) for line in f if line.strip()]
    buyer_headers = list(buyer_all[0])
    handles = [row["Handle"] for row in seo_rows]
    if handles != list(COPY):
        raise RuntimeError("B033 scope does not match the R109 copy map")
    for row in seo_rows:
        title, primary, secondary, motif = COPY[row["Handle"]]
        article = "a" if primary.startswith("USA") else ("an" if primary.lower()[0] in "aeiou" else "a")
        meta = f"Create {article} {primary} with {motif}."
        description = f"Create {article} {primary} with {motif}. The design adds a distinctive focal point to a fan room, game space, front door, entryway, or pet-loving home."
        row.update({
            "title_proposed": title, "meta_title_seo": title, "meta_title_chars": str(len(title)),
            "meta_description_seo": meta, "meta_description_chars": str(len(meta)),
            "description_proposed": description, "description_proposed_html": f"<p>{description}</p>",
            "primary_keyword": primary, "secondary_keywords": secondary,
            "long_tail_candidates": f"{primary}; {secondary}", "revision": "2",
            "review_status": "NEEDS_REVIEW", "content_qa_status": "NOT_RUN",
            "review_reason": f"{REVISION} removes internal evidence-note language and unsupported claim wording from {BATCH} copy.",
            "issues": f"{REVISION} requires hardened re-QA; admin/export before-state and direct demand data remain limitations.",
        })
    handles_set = set(handles)
    evidence_ids = {r.get("evidence_id") for r in seo_rows}
    evidence_scope = [r for r in evidence_rows if r.get("evidence_id") in evidence_ids]
    keyword_scope = [r for r in keyword_rows if r.get("product_key") in handles_set]
    buyer_rows = [r for r in buyer_all if r.get("product_key") in handles_set]
    img_scope = [r for r in img_rows if r.get("Handle") in handles_set]
    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "SEO_Products", seo_rows, seo_headers)
    write_sheet(wb, "Image_Audit", img_scope, img_headers)
    write_sheet(wb, "Product_Evidence", evidence_scope, evidence_headers)
    write_sheet(wb, "Keyword_Map", keyword_scope, keyword_headers)
    write_sheet(wb, "Buyer_Search_Research", buyer_rows, buyer_headers)
    write_sheet(wb, "README_QA", [{
        "metric": "revision", "value": REVISION, "definition": f"Canonical {BATCH} cleaned-copy revision for hardened re-QA."
    }, {
        "metric": "scope", "value": f"{BATCH}: {len(handles)} products, {len(img_scope)} images", "definition": "Scope is exactly frozen to the original research batch."
    }, {
        "metric": "approval_status", "value": "NOT_APPROVED_NOT_DEPLOYED", "definition": "QA pass is not human approval or Shopify deployment."
    }], ["metric", "value", "definition"])
    log = wb.create_sheet("Revision_Log")
    log.append(["revision_batch_id", "revision", "batch_id", "handle", "scope_status", "review_status", "content_qa_status", "source_workbook", "evidence_id"])
    for row in seo_rows:
        if row.get("evidence_id") not in evidence_ids:
            raise RuntimeError(f"Missing evidence row for {row['Handle']}: {row.get('evidence_id')}")
        log.append([REVISION, 2, BATCH, row["Handle"], "LOCKED", "NEEDS_REVIEW", "NOT_RUN", f"seo_runs/{SHOP}/{RUN_ID}/batches/{BATCH}_SEO_Products.csv", row.get("evidence_id")])
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    manifest = {"revision": REVISION, "batch_id": BATCH, "scope_count": len(handles), "image_count": len(img_scope), "canonical_workbook": str(OUT), "sha256": sha256(OUT), "status": "CREATED_FOR_HARDENED_REQA", "approval_status": "NOT_APPROVED_NOT_DEPLOYED"}
    (OUT_DIR / f"revision_{REVISION}_manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(manifest, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

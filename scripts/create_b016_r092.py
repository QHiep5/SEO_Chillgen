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
BATCH = "B016"
REVISION = "R092"
RUN = BASE / f"seo_runs/{SHOP}/{RUN_ID}"
OUT_DIR = BASE / f"resutls/{SHOP}/{RUN_ID}/revisions/{REVISION}"
OUT = OUT_DIR / f"SEO_Product_Optimization_revision_{REVISION}.xlsx"

COPY = {
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-112": ("Witch Hat Pumpkin Halloween Rug", "witch hat pumpkin Halloween rug", "pumpkin Halloween rug, witch hat round rug, spooky seasonal floor decor", "witch hat and pumpkin artwork, a deep optical-illusion effect, autumn colors, and a playful Halloween scene"),
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-113": ("Grim Reaper Halloween Rug", "grim reaper Halloween rug", "grim reaper rug, cemetery Halloween rug, spooky optical illusion floor decor", "grim reaper artwork, cemetery details, a dramatic tunnel effect, and dark Halloween colors"),
    "custom-halloween-3d-optical-illusion-round-rug-skeleton-h-design-114": ("Haunted House Halloween Rug", "haunted house Halloween rug", "haunted house rug, Halloween mansion rug, spooky round floor decor", "haunted house artwork, pumpkins, moonlit details, and a dramatic seasonal scene"),
    "personalized-orthodox-christian-area-rug-custom-eastern-orthodox": ("Blue Orthodox Cross Area Rug", "blue Orthodox cross area rug", "Orthodox Christian rug, blue cross area rug, religious home decor", "a blue Orthodox cross, decorative border details, and a calm religious design for a living space"),
    "personalized-orthodox-christian-area-rug-custom-eastern-design-100": ("Red Gold Orthodox Cross Rug", "red gold Orthodox cross rug", "red Christian cross rug, Orthodox area rug, religious room decor", "a red and gold Orthodox cross, ornate border details, and a warm religious design for home decor"),
    "personalized-orthodox-christian-area-rug-custom-eastern-design-101": ("Book Style Orthodox Cross Rug", "book style Orthodox cross rug", "book style Christian rug, Orthodox cross area rug, religious home decor", "a book-style Orthodox cross, red and gold details, and a distinctive religious design for a room"),
    "personalized-orthodox-christian-area-rug-custom-eastern-design-102": ("Black Gold Orthodox Cross Rug", "black gold Orthodox cross rug", "black Christian cross rug, black gold Orthodox rug, religious room decor", "a black and gold Orthodox cross, ornate border details, and a dramatic religious design for home decor"),
    "personalized-orthodox-christian-area-rug-custom-eastern-design-103": ("Burgundy Orthodox Cross Rug", "burgundy Orthodox cross rug", "burgundy Christian rug, red floral Orthodox rug, religious area rug", "a burgundy Orthodox cross, red floral border details, and a rich religious design for a living space"),
    "personalized-orthodox-christian-area-rug-custom-eastern-design-104": ("Black Orthodox Cross Area Rug", "black Orthodox cross area rug", "black Christian rug, Orthodox cross area rug, religious home decor", "a black Orthodox cross, gold border details, and a refined religious design for a room"),
    "personalized-orthodox-christian-area-rug-custom-eastern-design-105": ("Navy Gold Orthodox Cross Rug", "navy gold Orthodox cross rug", "navy Christian rug, Orthodox cross area rug, religious room decor", "a navy and gold Orthodox cross, decorative border details, and a classic religious design for home decor"),
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
        raise RuntimeError("B016 scope does not match the R092 copy map")
    for row in seo_rows:
        title, primary, secondary, motif = COPY[row["Handle"]]
        meta = f"Create a {primary} with {motif}."
        description = f"Create a {primary} with {motif}. The design adds a distinctive seasonal or religious focal point to a living room, entryway, bedroom, or display space."
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
    write_sheet(wb, "README_QA", [
        {"metric": "revision", "value": REVISION, "definition": f"Canonical {BATCH} cleaned-copy revision for hardened re-QA."},
        {"metric": "scope", "value": f"{BATCH}: {len(handles)} products, {len(img_scope)} images", "definition": "Scope is exactly frozen to the original research batch."},
        {"metric": "approval_status", "value": "NOT_APPROVED_NOT_DEPLOYED", "definition": "QA pass is not human approval or Shopify deployment."},
    ], ["metric", "value", "definition"])
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

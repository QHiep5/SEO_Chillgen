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
BATCH = "B023"
REVISION = "R099"
RUN = BASE / f"seo_runs/{SHOP}/{RUN_ID}"
OUT_DIR = BASE / f"resutls/{SHOP}/{RUN_ID}/revisions/{REVISION}"
OUT = OUT_DIR / f"SEO_Product_Optimization_revision_{REVISION}.xlsx"

COPY = {
    "custom-road-map-area-rug-play-mat-9aed6a6ee9-9aed6a6ee9": ("Adventure Town Road Map Play Rug", "adventure town road map play rug", "town road map kids rug, city streets playroom mat, classroom road rug", "an adventure town map with streets and colorful play details for a nursery, classroom, or kids playroom"),
    "custom-road-map-area-rug-personalized-city-street-car-rug-6313a0fee7": ("Construction City Road Map Rug", "construction city road map rug", "construction road play rug, truck city map rug, personalized classroom mat", "a construction city map with trucks, streets, and play details for a classroom, bedroom, nursery, or playroom"),
    "personalized-christmas-doormat-custom-family-name-7358e56962-7358e56962": ("Personalized Santa Stop Here Doormat", "personalized Santa Stop Here doormat", "custom Santa Christmas mat, family name holiday doormat, Santa entryway decor", "Santa Stop Here wording with personalized family details for a front door, porch, or holiday entryway"),
    "personalized-christmas-doormat-80a01c19f6-80a01c19f6": ("Personalized Pet Stocking Christmas Doormat", "personalized pet stocking Christmas doormat", "pet Christmas welcome mat, custom dog stocking doormat, personalized holiday entryway mat", "pet stocking artwork with personalized pet names and Merry Christmas wording for a holiday entryway"),
    "personalized-christmas-doormat-black-bear-family-e3c9d6a891-e3c9d6a891": ("Personalized Black Bear Family Doormat", "personalized black bear family Christmas doormat", "bear family Christmas mat, custom holiday welcome doormat, family entryway decor", "black bear family artwork with personalized family details for a front door, porch, or holiday display"),
    "personalized-christmas-doormat-family-name-5fb80e7372-5fb80e7372": ("Personalized Red Monogram Christmas Doormat", "personalized red monogram Christmas doormat", "red Christmas family mat, custom monogram holiday doormat, personalized entryway decor", "a red monogram and family-name design for a front door, porch, or Christmas entryway"),
    "personalized-christmas-doormat-with-custom-family-name-83c3ea451f": ("Personalized Family Stocking Doormat", "personalized family stocking Christmas doormat", "custom stocking Christmas mat, family name holiday doormat, personalized entryway decor", "family stocking artwork with custom names for a front door, porch, or holiday entryway"),
    "personalized-christmas-doormat-with-custom-family-name-d66e920e14": ("Custom 3D Snowman Christmas Doormat", "custom 3D snowman Christmas doormat", "snowman holiday welcome mat, personalized Christmas entryway mat, custom family doormat", "a 3D snowman design with personalized family details for a front door, porch, or Christmas display"),
    "personalized-christmas-welcome-doormat-28640346cc": ("Personalized Pet Family Christmas Doormat", "personalized pet family Christmas doormat", "custom pet Christmas mat, dog cat holiday welcome doormat, personalized entryway decor", "pet family artwork with customizable pet details for a front door, porch, or holiday entryway"),
    "personalized-christmas-doormat-1a037c0533": ("Custom Santa Letter Christmas Doormat", "custom Santa letter Christmas doormat", "Santa letter holiday mat, personalized Christmas welcome doormat, family entryway decor", "a Santa letter-inspired design with personalized family details for a front door, porch, or Christmas entryway"),
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
        raise RuntimeError("B023 scope does not match the R099 copy map")
    for row in seo_rows:
        title, primary, secondary, motif = COPY[row["Handle"]]
        meta = f"Create a {primary} with {motif}."
        description = f"Create a {primary} with {motif}. The design adds a welcoming seasonal focal point to a front door, porch, holiday entryway, nursery, classroom, or kids playroom."
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

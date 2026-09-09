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
BATCH = "B022"
REVISION = "R098"
RUN = BASE / f"seo_runs/{SHOP}/{RUN_ID}"
OUT_DIR = BASE / f"resutls/{SHOP}/{RUN_ID}/revisions/{REVISION}"
OUT = OUT_DIR / f"SEO_Product_Optimization_revision_{REVISION}.xlsx"

COPY = {
    "personalized-album-cover-rug-abb1eeca32": ("Pink Personalized Album Cover Rug", "pink personalized album cover rug", "pink music rug, custom album cover area rug, personalized bedroom decor", "a pink album-cover design with personalized song or artist details for a bedroom, living room, studio, or music space"),
    "custom-road-map-door-mat-town-design-e3598e955a-e3598e955a": ("Personalized Town Road Map Play Rug", "personalized town road map play rug", "custom city streets kids rug, town map playroom rug, personalized road rug", "a personalized town map with streets, cars, and playroom details for a classroom, nursery, or kids play space"),
    "custom-road-map-area-rug-8262be4ef1-8262be4ef1": ("Alphabet Road Map Kids Rug", "alphabet road map kids rug", "ABC road rug, alphabet playroom rug, educational city map rug", "an alphabet road map with letters, streets, and colorful play elements for a classroom, nursery, or playroom"),
    "custom-road-map-area-rug-town-play-mat-587b1148b8-587b1148b8": ("Blue Town Road Map Play Rug", "blue town road map play rug", "blue city street kids rug, town play mat, road map classroom rug", "a blue town map with city streets and vehicle-themed play details for a classroom, toy room, or playroom"),
    "custom-road-map-area-rug-e4dbf01f05": ("Pastel City Road Map Rug", "pastel city road map rug", "pastel city play rug, road map classroom mat, kids streets rug", "a pastel city road map with streets and play details for a classroom, bedroom, nursery, or playroom"),
    "custom-road-map-welcome-mat-708eca28bb": ("Personalized City Street Map Rug", "personalized city street map rug", "custom city street rug, personalized road map mat, classroom play rug", "a personalized city street map with named-road details for a classroom, bedroom, nursery, or playroom"),
    "custom-road-map-doormat-city-street-play-mat-0e3ba7dda0-0e3ba7dda0": ("Work Zone Construction Road Rug", "work zone construction road rug", "construction road play rug, truck kids rug, work zone classroom mat", "a work-zone road map with construction vehicles and street details for a classroom, nursery, or playroom"),
    "custom-road-map-area-rug-town-play-mat-a27d8ff428": ("Gray Construction Road Map Rug", "gray construction road map rug", "gray truck play rug, construction city map rug, road playroom mat", "a gray construction road map with truck-themed details for a bedroom, classroom, or playroom"),
    "custom-road-map-area-rug-d39d38a4a8": ("Soft City Road Play Rug", "soft city road play rug", "soft city map rug, road playroom rug, classroom street mat", "a soft-toned city road map with streets and play details for a classroom, reading corner, bedroom, or playroom"),
    "custom-road-map-area-rug-60aa327740": ("ABC Town Road Map Rug", "ABC town road map rug", "ABC town play rug, alphabet city map rug, kids road classroom mat", "an ABC town road map with letters, streets, and colorful play details for a classroom, nursery, or playroom"),
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
        raise RuntimeError("B022 scope does not match the R098 copy map")
    for row in seo_rows:
        title, primary, secondary, motif = COPY[row["Handle"]]
        meta = f"Create a {primary} with {motif}."
        description = f"Create a {primary} with {motif}. The design adds a colorful play and learning focal point to a classroom, nursery, bedroom, or kids playroom."
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

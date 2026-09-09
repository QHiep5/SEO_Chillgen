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
BATCH = "B021"
REVISION = "R097"
RUN = BASE / f"seo_runs/{SHOP}/{RUN_ID}"
OUT_DIR = BASE / f"resutls/{SHOP}/{RUN_ID}/revisions/{REVISION}"
OUT = OUT_DIR / f"SEO_Product_Optimization_revision_{REVISION}.xlsx"

COPY = {
    "personalized-album-cover-rug-4af618eed5": ("Black Personalized Album Cover Rug", "black personalized album cover rug", "custom black music rug, personalized album art rug, music room floor decor", "a black album-cover design with personalized song or artist details for a studio, living room, bedroom, or music space"),
    "personalized-album-cover-welcome-mat-26099b1534": ("Teal Personalized Album Cover Rug", "teal personalized album cover rug", "teal music rug, custom album cover mat, personalized music room decor", "a teal album-cover design with personalized music details for an entryway, living room, bedroom, or music space"),
    "personalized-album-cover-area-rug-648fd8e2ff": ("Orange Personalized Music Album Rug", "orange personalized music album rug", "orange album art rug, custom music area rug, personalized bedroom decor", "an orange album-cover design with personalized music details for a bedroom, living room, studio, or music space"),
    "personalized-album-cover-area-rug-0587b91810-0587b91810": ("Beige Personalized Album Cover Rug", "beige personalized album cover rug", "beige music rug, custom album area rug, personalized home decor", "a beige album-cover design with personalized song or artist details for a bedroom, living room, studio, or music space"),
    "personalized-album-cover-rug-3bfef01c49-3bfef01c49": ("Brown Personalized Album Cover Rug", "brown personalized album cover rug", "brown music rug, custom album cover area rug, personalized dining room decor", "a brown album-cover design with personalized music details for a dining room, living room, bedroom, or music space"),
    "personalized-album-cover-area-rug-52d1b2412d-52d1b2412d": ("Burgundy Personalized Music Album Rug", "burgundy personalized music album rug", "burgundy album art rug, custom music area rug, personalized bedroom decor", "a burgundy album-cover design with personalized music details for a bedroom, living room, studio, or music space"),
    "personalized-album-cover-rug-75f5a15f3e-75f5a15f3e": ("Gray Personalized Album Cover Rug", "gray personalized album cover rug", "gray music rug, custom album cover rug, personalized bedroom decor", "a gray album-cover design with personalized song or artist details for a bedroom, living room, studio, or music space"),
    "custom-album-cover-rug-music-decor-f2f905d19b-f2f905d19b": ("Music Studio Album Art Rug", "music studio album art rug", "guitar music rug, custom album art area rug, music room decor", "a music studio scene with guitars and album-inspired artwork for a dining room, living room, bedroom, or studio"),
    "personalized-music-album-cover-welcome-mat-6b66b8a24c": ("Turntable Personalized Music Rug", "turntable personalized music rug", "turntable music rug, personalized album cover mat, music entryway decor", "a turntable and album-cover design with personalized music details for an entryway, living room, office, or music space"),
    "personalized-music-album-cover-rug-5088343293-5088343293": ("Personalized Song Waveform Rug", "personalized song waveform rug", "photo waveform music rug, custom song rug, personalized album cover decor", "a personalized song photo and waveform design for a living room, bedroom, entryway, or music-themed space"),
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
        raise RuntimeError("B021 scope does not match the R097 copy map")
    for row in seo_rows:
        title, primary, secondary, motif = COPY[row["Handle"]]
        meta = f"Create a {primary} with {motif}."
        description = f"Create a {primary} with {motif}. The design adds a distinctive music-themed focal point to a bedroom, living room, entryway, studio, dining room, or creative space."
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

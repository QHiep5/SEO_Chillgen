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
BATCH = "B030"
REVISION = "R106"
RUN = BASE / f"seo_runs/{SHOP}/{RUN_ID}"
OUT_DIR = BASE / f"resutls/{SHOP}/{RUN_ID}/revisions/{REVISION}"
OUT = OUT_DIR / f"SEO_Product_Optimization_revision_{REVISION}.xlsx"

COPY = {
    "custom-running-horse-welcome-mat-2e0754fd8e-2e0754fd8e": ("Charcoal Running Horse Rug", "charcoal running horse rug", "black and white horse rug, galloping horse area rug, equestrian room carpet", "a charcoal running horse design for a bedroom, living room, entryway, office, or equestrian space"),
    "custom-running-horse-welcome-mat-f114a916e7-f114a916e7": ("Running Horses Area Rug", "running horses area rug", "horse herd rug, western horse carpet, equestrian area rug", "a group of running horses for a bedroom, living room, entryway, office, or equestrian space"),
    "custom-running-horse-welcome-mat-a4c32dbbfa-a4c32dbbfa": ("Dark Horse Eye Rug", "dark horse eye rug", "horse eye area rug, equestrian art rug, horse lover floor mat", "a close-up dark horse eye design for a bedroom, living room, office, or equestrian space"),
    "custom-running-horse-area-rug-2f0ce836fd-2f0ce836fd": ("Black White Horse Rug", "black white horse rug", "two horse area rug, equestrian bedroom rug, horse lover carpet", "a black and white horse design for a bedroom, living room, entryway, office, or equestrian space"),
    "western-running-horse-area-rug-02dee68d0b": ("Mint Horse Area Rug", "mint horse area rug", "white horse rug, floral horse bedroom rug, equestrian decor carpet", "a mint-toned horse design for a bedroom, living room, office, or equestrian space"),
    "custom-running-horse-area-rug-31f0e158c5": ("Floral Horse Herd Rug", "floral horse herd rug", "horse herd area rug, equestrian floral rug, horse lover bedroom carpet", "a horse herd design with floral artwork for a bedroom, living room, office, or equestrian space"),
    "personalized-classroom-doormat-teacher-name-notebook-pattern-42da091fe1-42da091fe1": ("Ocean Classroom Welcome Mat", "ocean classroom welcome mat", "personalized teacher doormat, classroom welcome rug, under the sea teacher mat", "an ocean classroom design for a school doorway, classroom entry, teacher office, or learning space"),
    "personalized-classroom-welcome-doormat-notebook-rug-8e1bdf83d4": ("Teacher Notebook Welcome Mat", "teacher notebook welcome mat", "personalized teacher doormat, classroom notebook rug, back to school welcome mat", "a teacher notebook design for a school doorway, classroom entry, teacher office, or learning space"),
    "personalized-classroom-welcome-doormat-teacher-students-7f5a91b5a9": ("Speech Pathologist Welcome Mat", "speech pathologist welcome mat", "personalized SLP doormat, speech therapy classroom mat, teacher office welcome rug", "a speech pathologist classroom design for a school doorway, therapy room, teacher office, or learning space"),
    "personalized-classroom-welcome-mat-teacher-name-ff4e426bbe": ("Teacher Name Classroom Mat", "teacher name classroom mat", "personalized classroom welcome mat, teacher doormat, school entrance rug", "a personalized teacher name design for a school doorway, classroom entry, teacher office, or learning space"),
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
        raise RuntimeError("B030 scope does not match the R106 copy map")
    for row in seo_rows:
        title, primary, secondary, motif = COPY[row["Handle"]]
        article = "an" if primary.lower().startswith(("a ", "e ", "i ", "o ", "u ")) else "a"
        meta = f"Create {article} {primary} with {motif}."
        description = f"Create {article} {primary} with {motif}. The design adds a distinctive focal point to a bedroom, living room, entryway, classroom, office, or equestrian space."
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

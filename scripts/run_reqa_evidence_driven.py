"""Evidence-driven re-QA generator for revision workbooks.

The generator never fabricates visual/image scores.  Image ratings are read
from an optional reviewer evidence JSON; without that file image criteria are
NOT_CHECKED and the product cannot become QA_PASS.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path
from statistics import mean

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
RUN_ID = "chillgen_20260907_01"
SHOP_DOMAIN = "chillgen.com"
TZ = timezone(timedelta(hours=7))
WEIGHTS = {"P1": 15, "P2": 10, "K1": 10, "K2": 5, "K3": 5,
           "T1": 10, "T2": 5, "D1": 5, "D2": 10, "I1": 20, "E1": 5}
IMAGE_WEIGHTS = {"IM1": 40, "IM2": 30, "IM3": 20, "IM4": 10}
INTERNAL_PATTERNS = [
    "seo copy avoids", "unless confirmed during admin/export review",
    "admin/export review", "surface-performance claims", "qa note",
    "re-run qa", "internal note",
]


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def rows(ws):
    it = ws.iter_rows(values_only=True)
    header = [str(x) if x is not None else "" for x in next(it)]
    return [dict(zip(header, row)) for row in it]


def val(row, key):
    x = row.get(key)
    return "" if x is None else str(x)


def append(ws, values):
    safe = []
    for x in values:
        if isinstance(x, str) and x[:1] in ("=", "+", "-", "@"):
            safe.append("'" + x)
        else:
            safe.append(x)
    ws.append(safe)


def header(ws, values):
    append(ws, values)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def finish(ws):
    widths = {}
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            widths[c.column] = max(widths.get(c.column, 0), min(70, len(str(c.value or "")) + 2))
    for i, width in widths.items():
        ws.column_dimensions[get_column_letter(i)].width = max(12, width)


def rating_row(rating, weight):
    if rating == "FULL":
        return weight
    if rating == "PARTIAL":
        return weight / 2
    return 0


def load_manual_image_evidence(path: Path | None):
    if not path or not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    return data.get("images", data)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--revision", required=True, help="Revision id, e.g. R001")
    ap.add_argument("--source-workbook", type=Path)
    ap.add_argument("--manual-image-evidence", type=Path,
                    help="JSON with independently reviewed image ratings")
    args = ap.parse_args()
    revision = args.revision.upper()
    if not re.fullmatch(r"R\d{3}", revision):
        raise SystemExit("--revision must look like R001")
    source = args.source_workbook or (BASE / f"resutls/{SHOP_DOMAIN}/{RUN_ID}/revisions/{revision}/SEO_Product_Optimization_revision_{revision}.xlsx")
    if not source.exists():
        raise FileNotFoundError(source)
    qa_run_id = f"QA-{datetime.now(TZ):%Y%m%d}-REV-{revision}"
    qa_batch_id = f"REV_{revision}"
    out_dir = BASE / f"resutls/{SHOP_DOMAIN}/{RUN_ID}/qa/{qa_run_id}"
    work_dir = BASE / f"seo_runs/{SHOP_DOMAIN}/{RUN_ID}/qa/{qa_run_id}"
    out_dir.mkdir(parents=True, exist_ok=True)
    work_dir.mkdir(parents=True, exist_ok=True)
    checked_at = datetime.now(TZ).replace(microsecond=0).isoformat()
    source_hash = sha256(source)

    wb = load_workbook(source, read_only=True, data_only=True)
    revlog = rows(wb["Revision_Log"])
    keys = [val(r, "handle") for r in revlog if val(r, "revision_batch_id") == revision]
    if not keys:
        raise RuntimeError(f"No Revision_Log scope found for {revision}")
    products = {val(r, "Handle"): r for r in rows(wb["SEO_Products"]) if val(r, "Handle") in keys}
    evidence = {val(r, "evidence_id"): r for r in rows(wb["Product_Evidence"])}
    images = {k: [] for k in keys}
    for r in rows(wb["Image_Audit"]):
        if val(r, "Handle") in images:
            images[val(r, "Handle")].append(r)
    if len(products) != len(keys):
        raise RuntimeError(f"Revision scope mismatch: log={len(keys)}, products={len(products)}")
    manual_images = load_manual_image_evidence(args.manual_image_evidence)

    qproducts, qcriteria, qimages, qissues = [], [], [], []
    status_counts = {"QA_PASS": 0, "QA_REVISE": 0, "QA_FAIL": 0, "QA_INCOMPLETE": 0}
    scores = []
    for handle in keys:
        p = products[handle]
        ev = evidence.get(val(p, "evidence_id"), {})
        blob = " ".join(val(p, k) for k in ["title_proposed", "meta_title_seo", "meta_description_seo", "description_proposed", "description_proposed_html"]).lower()
        refs = "; ".join([f"source_workbook_sha256:{source_hash}", f"evidence_id:{val(p, 'evidence_id')}", f"revision_log:{revision}"])
        issues = []
        critical = major = minor = 0

        def issue(sev, field, observed, reason, fix):
            nonlocal critical, major, minor
            n = sum(1 for i in issues if i[3] == sev) + 1
            iid = f"{revision}-QA-{sev[0]}{n:03d}-{handle[:20]}"
            issues.append([iid, handle, "", sev, field, observed, reason, observed, fix, refs, "Re-run after correction."])
            if sev == "CRITICAL": critical += 1
            elif sev == "MAJOR": major += 1
            else: minor += 1

        internal_hits = [pat for pat in INTERNAL_PATTERNS if pat in blob]
        if internal_hits:
            issue("MAJOR", "proposed_fields", ", ".join(internal_hits), "Internal QA/process language is customer-visible copy.", "Remove internal notes and retain only customer-facing product copy.")
        unsupported = [term for term in ["outdoor", "waterproof", "machine washable", "non-slip", "anti-slip"] if term in blob and term not in (val(ev, "verified_product_facts") + " " + val(ev, "short_source_excerpt")).lower()]
        if unsupported:
            issue("MAJOR", "proposed_fields", ", ".join(unsupported), "Potential product claim is not supported by the linked evidence record.", "Remove or verify each claim against the source/export before approval.")

        criteria = []
        def add(cid, rating, reason, evidence_ref=""):
            earned = rating_row(rating, WEIGHTS[cid])
            criteria.append((cid, rating, reason, earned, evidence_ref))
        add("P1", "FULL" if val(p, "product_type") and val(ev, "verified_product_facts") else "NOT_CHECKED", "Product identity is checked only when product type and verified facts are present.")
        add("P2", "FAIL" if unsupported else ("FULL" if val(ev, "verified_product_facts") else "NOT_CHECKED"), "Claims are checked against verified_product_facts and source excerpt.")
        km = [r for r in rows(wb["Keyword_Map"]) if val(r, "product_key") == handle]
        add("K1", "FULL" if any(val(r, "keyword_role") == "PRIMARY" and val(r, "mapping_status") for r in km) else "NOT_CHECKED", "Derived from this product's Keyword_Map rows; no keyword is assumed from title text.")
        add("K2", "FULL" if any(val(r, "representative_SERP_URLs") for r in km) else "NOT_CHECKED", "SERP fit is credited only when representative SERP references exist.")
        add("K3", "PARTIAL" if val(p, "keyword_demand_evidence") else "NOT_CHECKED", "Demand evidence level is disclosed; SERP-only evidence is not treated as volume proof.")
        add("T1", "FULL" if 45 <= len(val(p, "meta_title_seo")) <= 70 else "PARTIAL" if val(p, "meta_title_seo") else "NOT_CHECKED", f"meta_title_seo length={len(val(p, 'meta_title_seo'))}; length alone does not prove semantic quality.")
        add("T2", "FULL" if val(p, "title_proposed") else "NOT_CHECKED", "Product title is checked for presence; semantic correctness still depends on independent product review.")
        md_len = len(val(p, "meta_description_seo"))
        add("D1", "FULL" if 80 <= md_len <= 320 else "PARTIAL" if val(p, "meta_description_seo") else "NOT_CHECKED", f"meta_description_seo length={md_len}.")
        add("D2", "FAIL" if internal_hits else ("FULL" if val(p, "description_proposed_html") else "NOT_CHECKED"), "HTML description must be present and free of internal process notes.")
        add("E1", "FULL" if val(p, "evidence_id") and val(ev, "fact_to_source_map") and val(p, "revision") else "NOT_CHECKED", "Traceability requires evidence id, fact-to-source map and revision.")

        for img in sorted(images[handle], key=lambda x: int(val(x, "image_number") or 0)):
            ikey = f"{handle}__img_{int(val(img, 'image_number') or 0):02d}"
            m = manual_images.get(ikey, {}) if isinstance(manual_images, dict) else {}
            ratings = {cid: m.get(cid, "NOT_CHECKED") for cid in IMAGE_WEIGHTS}
            for cid, rating in ratings.items():
                if rating not in {"FULL", "PARTIAL", "FAIL", "NOT_CHECKED"}:
                    raise RuntimeError(f"Invalid {cid} rating for {ikey}: {rating}")
            img_points = sum(rating_row(ratings[c], IMAGE_WEIGHTS[c]) for c in IMAGE_WEIGHTS)
            img_assessed = sum(IMAGE_WEIGHTS[c] for c in IMAGE_WEIGHTS if ratings[c] != "NOT_CHECKED")
            img_final = img_points if img_assessed == 100 else ""
            qimages.append([handle, ikey, val(img, "image_url"), val(img, "image_url_export") or val(img, "image_url"), val(img, "media_id"), val(img, "variant"), val(img, "image_location"), "INDEPENDENT_REVIEW_JSON" if m else "NO_INDEPENDENT_IMAGE_REVIEW", checked_at, val(img, "observed_visual_details"), val(img, "observed_visual_details"), val(img, "alt_action"), val(img, "alt_proposed"), ratings["IM1"], ratings["IM2"], ratings["IM3"], ratings["IM4"], img_points, img_assessed, img_final, img_points, img_points + (100-img_assessed), "", refs])
        image_complete = bool(images[handle]) and all(all(manual_images.get(f"{handle}__img_{int(val(i, 'image_number') or 0):02d}", {}).get(c) in {"FULL", "PARTIAL", "FAIL"} for c in IMAGE_WEIGHTS) for i in images[handle]) if isinstance(manual_images, dict) else False
        image_rating = "FULL" if image_complete and all(all(manual_images[f"{handle}__img_{int(val(i, 'image_number') or 0):02d}"][c] == "FULL" for c in IMAGE_WEIGHTS) for i in images[handle]) else "NOT_CHECKED"
        add("I1", image_rating, "Image score is derived only from independent per-image ratings; workbook metadata alone is not visual QA.")
        if not image_complete:
            iid = f"{revision}-QA-LIMG-{handle[:20]}"
            issues.append([iid, handle, "", "LIMITATION", "images", "", "No complete independent image-review JSON was supplied.", "Image criteria remain NOT_CHECKED.", "Review every in-scope image and provide IM1–IM4 ratings.", refs, "Re-run with --manual-image-evidence."])

        for cid, rating, reason, earned, eref in criteria:
            qcriteria.append([handle, cid, WEIGHTS[cid], "CHECKED" if rating != "NOT_CHECKED" else "NOT_CHECKED", rating, earned, WEIGHTS[cid] if rating != "NOT_CHECKED" else 0, reason, refs, eref])
        # Image criterion was appended last; ensure it is represented once.
        cid, rating, reason, earned, eref = criteria[-1]
        assessed = sum(WEIGHTS[c] for c, r, *_ in criteria if r != "NOT_CHECKED")
        verified = sum(e for c, r, _, e, _ in criteria) + (earned if cid == "I1" else 0)
        # criteria includes I1 as the last item, so do not double-add it.
        verified = sum(e for _, _, _, e, _ in criteria)
        upper = verified + (100 - assessed)
        final = verified if assessed == 100 and image_complete and not critical and not major else ""
        # Rubric precedence: CRITICAL -> FAIL; insufficient evidence -> INCOMPLETE;
        # only a fully assessed product can be REVISE/PASS by score.
        status = ("QA_FAIL" if critical or any(i[3] == "CRITICAL" for i in issues)
                  else "QA_INCOMPLETE" if final == ""
                  else "QA_REVISE" if major or final < 85
                  else "QA_PASS")
        status_counts[status] += 1
        if final != "": scores.append(final)
        qproducts.append([handle, val(p, "url"), val(p, "revision"), verified, assessed, verified, upper, final, status, "SERP_ONLY" if val(p, "keyword_demand_evidence") else "UNVERIFIABLE", len(images[handle]), len(images[handle]) if image_complete else 0, "TRUE" if image_complete else "UNKNOWN", "100%" if image_complete else "UNKNOWN", critical, major, minor, "; ".join(i[0] for i in issues), refs])
        qissues.extend(issues)

    out_xlsx = out_dir / f"SEO_QA_{qa_batch_id}.xlsx"
    out_md = out_dir / f"SEO_QA_{qa_batch_id}.md"
    wbout = Workbook(); ws = wbout.active; ws.title = "QA_Summary"
    header(ws, ["metric", "value", "definition"])
    batch_status = ("QA_PASS" if status_counts["QA_PASS"] == len(keys)
                    else "QA_FAIL" if status_counts["QA_FAIL"]
                    else "QA_INCOMPLETE" if status_counts["QA_INCOMPLETE"]
                    else "QA_REVISE")
    summary = [("rubric_version", "prompt_qa.md v1.0"), ("qa_run_id", qa_run_id), ("source_workbook", str(source)), ("source_workbook_sha256", source_hash), ("scope_product_count", len(keys)), ("scope_image_count", len(qimages)), ("batch_status", batch_status), ("batch_final_score", round(mean(scores), 1) if len(scores) == len(keys) else ""), ("QA_PASS", status_counts["QA_PASS"]), ("QA_REVISE", status_counts["QA_REVISE"]), ("QA_FAIL", status_counts["QA_FAIL"]), ("QA_INCOMPLETE", status_counts["QA_INCOMPLETE"]), ("approval_status", "NOT_APPROVED_NOT_DEPLOYED")]
    for k, v in summary: append(ws, [k, v, "Evidence-driven re-QA; blank final score means insufficient evidence."])
    ws = wbout.create_sheet("QA_Products"); header(ws, ["product_key","url","revision","verified_points","assessed_weight","score_lower_bound","score_upper_bound","final_score","qa_status","keyword_evidence_level","images_expected","images_checked","image_inventory_complete","image_coverage","critical_count","major_count","minor_count","issue_refs","evidence_refs"])
    for r in qproducts: append(ws, r)
    ws = wbout.create_sheet("QA_Criteria"); header(ws, ["product_key","criterion_id","weight","assessment","rating","earned_points","assessed_weight","reason","evidence_refs","issue_refs"])
    for r in qcriteria: append(ws, r)
    ws = wbout.create_sheet("QA_Images"); header(ws, ["product_key","qa_image_key","image_url_source","image_url_workbook","media_id","variant","image_location","check_method","checked_at","qa_observation","submitted_observation","alt_action","alt_effective","IM1","IM2","IM3","IM4","image_verified_points","image_assessed_weight","image_final_score","image_score_lower_bound","image_score_upper_bound","issue_refs","evidence_refs"])
    for r in qimages: append(ws, r)
    ws = wbout.create_sheet("QA_Issues"); header(ws, ["issue_id","product_key","qa_image_key","severity","field","submitted_value","source_observation","reason","recommended_fix","supporting_evidence","recheck_condition"])
    for r in qissues: append(ws, r)
    for s in wbout.worksheets: finish(s)
    wbout.save(out_xlsx)
    manifest = {"rubric_version":"prompt_qa.md v1.0", "qa_run_id":qa_run_id, "qa_batch_id":qa_batch_id, "source_workbook":str(source), "source_workbook_sha256":source_hash, "batch_product_keys":keys, "status_counts":status_counts, "batch_status":batch_status, "manual_image_evidence":str(args.manual_image_evidence) if args.manual_image_evidence else None, "approval_status":"NOT_APPROVED_NOT_DEPLOYED"}
    (work_dir / "manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    out_md.write_text(f"# Evidence-driven Re-QA {revision}\n\n- Batch status: `{batch_status}`\n- Products: {len(keys)}\n- Image rows: {len(qimages)}\n- Status counts: `{status_counts}`\n\nA blank `final_score` means the rubric did not have complete evidence. Image metadata or prior agent observations are not treated as independent visual QA.\n", encoding="utf-8")
    print(json.dumps({"qa_xlsx":str(out_xlsx),"qa_md":str(out_md),"batch_status":batch_status,"status_counts":status_counts}, indent=2))


if __name__ == "__main__":
    main()

import hashlib
import json
import os
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[6]
RUN_ID = "chillgen_20260907_01"
SHOP_DOMAIN = "chillgen.com"
SCRIPT_DIR = Path(__file__).resolve().parent
QA_BATCH_ID = os.environ.get("QA_BATCH_ID", sys.argv[1] if len(sys.argv) > 1 else SCRIPT_DIR.name)
QA_RUN_ID = f"QA-20260907-AGENT-{QA_BATCH_ID}"

BASE = ROOT / "SEO_Chillgen"
RUN_DIR = BASE / "seo_runs" / SHOP_DOMAIN / RUN_ID
RESULT_DIR = BASE / "resutls" / SHOP_DOMAIN / RUN_ID
QA_WORK_DIR = RUN_DIR / "qa" / QA_RUN_ID
QA_RESULT_DIR = RESULT_DIR / "qa" / QA_RUN_ID
INPUT_DIR = RUN_DIR / "qa_inputs" / QA_BATCH_ID
WORKBOOK = RESULT_DIR / "SEO_Product_Optimization.xlsx"
EXTRACT = INPUT_DIR / f"{QA_BATCH_ID.lower()}_workbook_extract.json"
LIVE_CHECKS = INPUT_DIR / f"{QA_BATCH_ID.lower()}_live_checks.json"
MANIFEST = INPUT_DIR / f"qa_input_{QA_BATCH_ID}.json"

CRITERIA = [
    ("P1", "Đúng sản phẩm và thiết kế", 15),
    ("P2", "Đúng thuộc tính và lời hứa", 10),
    ("K1", "Long-tail và nhu cầu phù hợp", 10),
    ("K2", "Intent và loại trang theo SERP", 5),
    ("K3", "Bằng chứng nhu cầu và tính trung thực", 5),
    ("T1", "SEO title", 10),
    ("T2", "Product Title/H1", 5),
    ("D1", "Meta description", 5),
    ("D2", "Product description", 10),
    ("I1", "Ảnh và alt", 20),
    ("E1", "Bằng chứng và nhất quán dữ liệu", 5),
]

IMAGE_CRITERIA = {"IM1": 40, "IM2": 30, "IM3": 20, "IM4": 10}
RATING_VALUE = {"FULL": 1.0, "PARTIAL": 0.5, "FAIL": 0.0, "NOT_CHECKED": None}


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def text(value):
    if value is None:
        return ""
    return str(value)


def handle_from_canonical(product_key: str) -> str:
    return product_key.split("::", 1)[-1]


def write_sheet(ws, headers, rows):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if isinstance(cell.value, str) and cell.value.startswith("http"):
                cell.hyperlink = cell.value
                cell.style = "Hyperlink"
    for idx, header in enumerate(headers, 1):
        max_len = len(header)
        for cell in ws[get_column_letter(idx)]:
            max_len = max(max_len, min(len(text(cell.value)), 90))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 55)


def rating_points(assessment, weight):
    value = RATING_VALUE.get(text(assessment).upper())
    if value is None:
        return 0.0, 0.0
    return weight * value, weight


def image_score(assessments):
    verified = 0.0
    assessed = 0.0
    for cid, weight in IMAGE_CRITERIA.items():
        earned, checked_weight = rating_points(assessments.get(cid, "NOT_CHECKED"), weight)
        verified += earned
        assessed += checked_weight
    return {
        "image_verified_points": round(verified, 4),
        "image_assessed_weight": round(assessed, 4),
        "image_final_score": round(verified, 4) if assessed == 100.0 else None,
        "image_score_lower_bound": round(verified, 4),
        "image_score_upper_bound": round(verified + (100.0 - assessed), 4),
    }


def product_status(final_score, critical_count, major_count):
    if critical_count > 0:
        return "QA_FAIL"
    if final_score is None:
        return "QA_INCOMPLETE"
    if final_score < 70:
        return "QA_FAIL"
    if final_score < 85 or major_count > 0:
        return "QA_REVISE"
    return "QA_PASS"


INTERNAL_TERMS = [
    "evidence note",
    "evidence_id",
    "serp observed",
    "buyer_search_research",
    "keyword_map",
    "qa not run",
    "public json",
    "source:",
    "validation_source",
    "unsupported claim",
    "needs_review",
    "mapping_status",
]

PROTECTED_CLAIMS = {
    "non-slip": ["non-slip", "non slip", "nonslip", "anti-slip", "anti slip"],
    "washable": ["washable", "machine washable", "machine-washable"],
    "outdoor": ["outdoor", "outdoors", "outside use"],
    "handmade": ["handmade", "hand made", "handcrafted", "hand-crafted"],
    "eco-friendly": ["eco-friendly", "eco friendly", "sustainable"],
    "premium": ["premium"],
    "wool": ["wool", "woolen", "woollen"],
    "waterproof": ["waterproof", "water-resistant", "water resistant"],
    "free shipping": ["free shipping", "free delivery"],
}


def lower(value):
    return text(value).lower()


def words(value):
    return [w for w in __import__("re").findall(r"[a-z0-9]+", lower(value)) if len(w) > 2]


def has_internal_language(value):
    haystack = lower(value)
    return [term for term in INTERNAL_TERMS if term in haystack]


def claim_hits(value):
    haystack = lower(value)
    found = []
    for claim, phrases in PROTECTED_CLAIMS.items():
        if any(p in haystack for p in phrases):
            found.append(claim)
    return found


def claims_supported(claims, evidence, imgs):
    support_text = " ".join(
        [
            lower(evidence.get("verified_product_facts")),
            lower(evidence.get("fact_to_source_map")),
            " ".join(lower(i.get("observed_visual_details")) for i in imgs),
        ]
    )
    unsupported = []
    for claim in claims:
        phrases = PROTECTED_CLAIMS[claim]
        if not any(p in support_text for p in phrases):
            unsupported.append(claim)
    return unsupported


def quality_title(title, primary):
    if not title:
        return "FAIL", "Missing title."
    internal = has_internal_language(title)
    if internal:
        return "FAIL", "Contains internal language: " + ", ".join(internal)
    title_len = len(title)
    primary_tokens = set(words(primary))
    title_tokens = set(words(title))
    has_keyword_overlap = not primary_tokens or bool(primary_tokens & title_tokens)
    if 35 <= title_len <= 70 and has_keyword_overlap:
        return "FULL", f"Readable title length {title_len}; overlaps primary keyword terms."
    if 25 <= title_len <= 80:
        return "PARTIAL", f"Readable title length {title_len}, but keyword coverage or length is borderline."
    return "FAIL", f"Title length {title_len} is outside practical SEO/title range."


def quality_meta(meta, product_text):
    if not meta:
        return "FAIL", "Missing meta description."
    internal = has_internal_language(meta)
    if internal:
        return "FAIL", "Contains internal language: " + ", ".join(internal)
    meta_len = len(meta)
    overlap = set(words(meta)) & set(words(product_text))
    if 120 <= meta_len <= 170 and len(overlap) >= 3:
        return "FULL", f"Meta length {meta_len}; shares product-specific terms."
    if 90 <= meta_len <= 190 and len(overlap) >= 2:
        return "PARTIAL", f"Meta length {meta_len}; product specificity is acceptable but borderline."
    return "FAIL", f"Meta length/specificity is weak: length={meta_len}, overlap_terms={len(overlap)}."


def quality_description(html, product_text):
    if not html:
        return "FAIL", "Missing proposed description HTML."
    internal = has_internal_language(html)
    if internal:
        return "FAIL", "Contains internal language: " + ", ".join(internal)
    plain = __import__("re").sub(r"<[^>]+>", " ", html)
    token_count = len(words(plain))
    overlap = set(words(plain)) & set(words(product_text))
    has_structure = "<p" in lower(html) and ("<li" in lower(html) or "<h3" in lower(html))
    if token_count >= 35 and len(overlap) >= 5 and has_structure:
        return "FULL", f"Structured buyer-facing HTML with {token_count} content words and product-specific overlap."
    if token_count >= 25 and len(overlap) >= 3:
        return "PARTIAL", f"Description is present but structure/specificity is borderline: words={token_count}, overlap={len(overlap)}."
    return "FAIL", f"Description is too thin or generic: words={token_count}, overlap={len(overlap)}."


def keyword_quality(primary, kws):
    if not primary:
        return "FAIL", "Missing primary keyword."
    token_count = len(words(primary))
    primary_rows = [
        k for k in kws
        if lower(k.get("keyword")) == lower(primary) and lower(k.get("keyword_role")) == "primary"
    ]
    if token_count >= 3 and primary_rows:
        return "FULL", "Primary keyword is long-tail and has a matching PRIMARY row in Keyword_Map."
    if token_count >= 2:
        return "PARTIAL", "Primary keyword is usable but matching Keyword_Map PRIMARY evidence is incomplete."
    return "FAIL", "Primary keyword is too broad for product-level targeting."


def intent_quality(kws):
    primary_rows = [k for k in kws if lower(k.get("keyword_role")) == "primary"]
    if not primary_rows:
        return "FAIL", "No PRIMARY keyword row found."
    row = primary_rows[0]
    page_type = lower(row.get("target_page_type"))
    intent = lower(row.get("intent"))
    if "product" in page_type and any(x in intent for x in ["commercial", "purchase", "transaction"]):
        return "FULL", "PRIMARY keyword maps to product page intent."
    if "product" in page_type:
        return "PARTIAL", "PRIMARY keyword maps to product page, but intent evidence is thin."
    return "FAIL", f"PRIMARY keyword maps to `{row.get('target_page_type')}` rather than product intent."


def demand_quality(product, kws, buyers):
    corpus = " ".join(
        [lower(product.get("keyword_demand_evidence"))]
        + [lower(k.get("demand_evidence")) + " " + lower(k.get("validation_source")) for k in kws]
        + [lower(b) for b in buyers]
    )
    if any(x in corpus for x in ["search console", "internal search", "review", "volume", "semrush", "ahrefs"]):
        return "FULL", "Direct/volume/customer demand evidence is present."
    if "serp" in corpus or "category" in corpus:
        return "PARTIAL", "Demand evidence is SERP/category-level only."
    return "FAIL", "No demand evidence found."


def evidence_quality(product, evidence, kws, buyers, imgs):
    required = [
        product.get("evidence_id"),
        evidence.get("evidence_id"),
        kws,
        buyers,
        imgs,
    ]
    if all(bool(x) for x in required) and lower(product.get("identity_status")) == "verified":
        return "FULL", "Workbook evidence, keyword, buyer, image, and verified identity are present."
    if all(bool(x) for x in required):
        return "PARTIAL", "Evidence is cross-linked, but identity/admin before-state is not fully verified."
    return "FAIL", "Missing cross-sheet evidence, keyword, buyer, or image data."


def assess_image_row(img, product, evidence):
    observed = text(img.get("observed_visual_details"))
    alt = text(img.get("alt_proposed"))
    image_url = text(img.get("image_url_export") or img.get("image_url"))
    product_terms = set(words(product.get("primary_keyword"))) | set(words(product.get("title_proposed")))
    observed_terms = set(words(observed))
    alt_terms = set(words(alt))
    assessments = {}
    reasons = []

    if image_url and observed:
        assessments["IM1"] = "PARTIAL"
        reasons.append("Image URL and workbook observation exist; actual visual match still depends on reviewer/contact-sheet confirmation.")
    else:
        assessments["IM1"] = "FAIL"
        reasons.append("Missing image URL or observed visual details.")

    if observed and not has_internal_language(observed):
        assessments["IM2"] = "PARTIAL"
        reasons.append("Observation is present and not internal-language, but not independently re-seen by this script.")
    else:
        assessments["IM2"] = "FAIL"
        reasons.append("Observation missing or contains internal language.")

    if alt and 8 <= len(alt) <= 125 and (alt_terms & (product_terms | observed_terms)):
        assessments["IM3"] = "FULL"
        reasons.append("Alt is present, practical length, and overlaps product/observation terms.")
    elif alt:
        assessments["IM3"] = "PARTIAL"
        reasons.append("Alt exists but product/observation specificity is borderline.")
    else:
        assessments["IM3"] = "FAIL"
        reasons.append("Missing alt.")

    repeated = len(alt_terms) - len(set(alt_terms))
    if alt and not has_internal_language(alt) and repeated <= 1 and len(words(alt)) <= 14:
        assessments["IM4"] = "FULL"
        reasons.append("Alt is concise and not keyword-stuffed.")
    elif alt:
        assessments["IM4"] = "PARTIAL"
        reasons.append("Alt may be long, repetitive, or contain weak wording.")
    else:
        assessments["IM4"] = "FAIL"
        reasons.append("Missing alt.")

    return assessments, " ".join(reasons)


def assess_product(product, live, evidence, kws, buyers, imgs):
    product_text = " ".join(
        text(product.get(c))
        for c in [
            "title_proposed",
            "meta_title_seo",
            "meta_description_seo",
            "description_proposed",
            "description_proposed_html",
            "keyword_selection_reason",
        ]
    )
    issues = []

    live_ok = bool(live.get("ok")) and int(live.get("status_code") or 0) == 200
    ev_images = None
    try:
        ev_images = int(float(evidence.get("gallery_image_count") or 0))
    except Exception:
        ev_images = None
    if live_ok and evidence and imgs and (ev_images in (None, 0, len(imgs))):
        p1 = ("FULL", 1.0, "Live PDP is reachable and workbook/evidence/image rows align for this product.")
    elif live_ok and evidence and imgs:
        p1 = ("PARTIAL", 0.5, f"Live PDP and evidence exist, but image count differs: evidence={ev_images}, Image_Audit={len(imgs)}.")
    else:
        p1 = ("FAIL", 0.0, "Live PDP/evidence/image scope is incomplete.")
        issues.append(("MAJOR", "P1_product_identity_scope", p1[2], "Re-check PDP, Product_Evidence, and Image_Audit links."))

    unsupported = claims_supported(claim_hits(product_text), evidence, imgs)
    if unsupported:
        p2 = ("FAIL", 0.0, "Unsupported protected claims found: " + ", ".join(unsupported))
        issues.append(("MAJOR", "unsupported_claims", p2[2], "Remove unsupported claims or add source evidence."))
    elif claim_hits(product_text):
        p2 = ("FULL", 1.0, "Protected claims in proposed copy are supported by evidence/image observations.")
    else:
        p2 = ("PARTIAL", 0.5, "No protected claim conflict found, but claims are limited to workbook evidence.")

    k1_assessment, k1_reason = keyword_quality(product.get("primary_keyword"), kws)
    k2_assessment, k2_reason = intent_quality(kws)
    k3_assessment, k3_reason = demand_quality(product, kws, buyers)
    t1_assessment, t1_reason = quality_title(product.get("meta_title_seo"), product.get("primary_keyword"))
    t2_assessment, t2_reason = quality_title(product.get("title_proposed"), product.get("primary_keyword"))
    d1_assessment, d1_reason = quality_meta(product.get("meta_description_seo"), product_text)
    d2_assessment, d2_reason = quality_description(product.get("description_proposed_html"), product_text)
    e1_assessment, e1_reason = evidence_quality(product, evidence, kws, buyers, imgs)

    for cid, assessment, reason in [
        ("K1", k1_assessment, k1_reason),
        ("K2", k2_assessment, k2_reason),
        ("K3", k3_assessment, k3_reason),
        ("T1", t1_assessment, t1_reason),
        ("T2", t2_assessment, t2_reason),
        ("D1", d1_assessment, d1_reason),
        ("D2", d2_assessment, d2_reason),
        ("E1", e1_assessment, e1_reason),
    ]:
        if assessment == "FAIL":
            issues.append(("MAJOR", cid, reason, "Revise this field and re-QA."))

    criteria = [
        ("P1", p1[0], p1[1], p1[2]),
        ("P2", p2[0], p2[1], p2[2]),
        ("K1", k1_assessment, RATING_VALUE[k1_assessment], k1_reason),
        ("K2", k2_assessment, RATING_VALUE[k2_assessment], k2_reason),
        ("K3", k3_assessment, RATING_VALUE[k3_assessment], k3_reason),
        ("T1", t1_assessment, RATING_VALUE[t1_assessment], t1_reason),
        ("T2", t2_assessment, RATING_VALUE[t2_assessment], t2_reason),
        ("D1", d1_assessment, RATING_VALUE[d1_assessment], d1_reason),
        ("D2", d2_assessment, RATING_VALUE[d2_assessment], d2_reason),
        ("E1", e1_assessment, RATING_VALUE[e1_assessment], e1_reason),
    ]
    return criteria, issues


def main():
    QA_WORK_DIR.mkdir(parents=True, exist_ok=True)
    QA_RESULT_DIR.mkdir(parents=True, exist_ok=True)

    extract = json.loads(EXTRACT.read_text(encoding="utf-8"))
    live_checks = json.loads(LIVE_CHECKS.read_text(encoding="utf-8"))
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    generated_at = datetime.now(timezone.utc).isoformat()
    wb_hash = sha256(WORKBOOK)

    products = {p["Handle"]: p for p in extract["products"]}
    live_by_handle = {r["handle"]: r for r in live_checks}
    evidence_by_handle = {
        row.get("product_url", "").rstrip("/").split("/")[-1]: row for row in extract["evidence"]
    }
    keywords_by_handle = {}
    buyer_by_handle = {}
    images_by_handle = {}
    auto_by_pk = {}

    for row in extract["keywords"]:
        keywords_by_handle.setdefault(row.get("product_key"), []).append(row)
    for row in extract["buyer"]:
        buyer_by_handle.setdefault(row.get("product_key"), []).append(row)
    for row in extract["images"]:
        images_by_handle.setdefault(row.get("Handle"), []).append(row)
    for issue in extract["auto_issues"]:
        auto_by_pk.setdefault(issue.get("product_key"), []).append(issue)

    issues = []
    image_rows = []
    criteria_rows = []
    product_rows = []
    issue_seq = 1

    for product_key in manifest["canonical_product_keys"]:
        handle = handle_from_canonical(product_key)
        product = products[handle]
        live = live_by_handle.get(handle, {})
        evidence = evidence_by_handle.get(handle, {})
        kws = keywords_by_handle.get(handle, [])
        buyers = buyer_by_handle.get(handle, [])
        imgs = sorted(images_by_handle.get(handle, []), key=lambda r: int(float(r.get("image_number") or 0)))
        issue_refs = []

        limitation_id = f"LIM-{issue_seq:04d}"
        issue_seq += 1
        issues.append(
            {
                "issue_id": limitation_id,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "LIMITATION",
                "field": "admin_identity_and_admin_seo_fields",
                "submitted_value": text(product.get("identity_status")),
                "source_observation": f"Live storefront URL/H1/canonical returned HTTP 200 and matched the handle, but {QA_BATCH_ID} source identity remains PUBLIC_JSON_MATCHED_TO_HANDLE and Shopify admin/export SEO values were not provided.",
                "reason": "Content QA can proceed from storefront/workbook evidence; deployment still needs admin/export identity and before-state verification.",
                "recommended_fix": "Before approval/deploy, export current Shopify product/admin SEO/media data and confirm product ID, Handle, SEO title, SEO description, Product Title, Body HTML, and media rows.",
                "supporting_evidence": f"{live.get('url', product.get('url'))}; {evidence.get('evidence_id', '')}",
                "recheck_condition": "Admin/export before-state matches the workbook product and intended fields.",
            }
        )
        issue_refs.append(limitation_id)

        limitation_id = f"LIM-{issue_seq:04d}"
        issue_seq += 1
        issues.append(
            {
                "issue_id": limitation_id,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "LIMITATION",
                "field": "keyword_demand_evidence",
                "submitted_value": text(product.get("primary_keyword")),
                "source_observation": "Keyword_Map records SERP/category pages and no paid volume or direct Chillgen customer query data; Buyer_Search_Research is SUPPORTED_CATEGORY_LANGUAGE_NO_VOLUME.",
                "reason": "The long-tail choice is product-appropriate, but demand proof is category/SERP-level rather than direct customer or volume evidence.",
                "recommended_fix": "If a stricter approval standard is required, add Search Console/internal search/review data or a fresh SERP note for the primary keyword.",
                "supporting_evidence": "; ".join([text(k.get("validation_source")) for k in kws if k.get("keyword_role") == "PRIMARY"]),
                "recheck_condition": "Direct or stronger query evidence is attached, or reviewer accepts the stated limitation.",
            }
        )
        issue_refs.append(limitation_id)

        limitation_id = f"LIM-{issue_seq:04d}"
        issue_seq += 1
        issues.append(
            {
                "issue_id": limitation_id,
                "product_key": product_key,
                "qa_image_key": "",
                "severity": "LIMITATION",
                "field": "independent_manual_qa_not_recorded",
                "submitted_value": "",
                "source_observation": "This corrected agent report has workbook/live-check metadata, but no reliable per-product semantic QA decisions or per-image visual QA decisions were recorded.",
                "reason": "The previous generator assigned FULL image/product criteria and fixed 95.0 scores. Corrected reports must not infer QA_PASS without explicit independent assessments.",
                "recommended_fix": "Run manual QA in the console or capture structured reviewer assessments for P1/P2/K1/K2/T1/T2/D1/D2 and every image IM1-IM4, then regenerate.",
                "supporting_evidence": f"{live.get('url', product.get('url'))}; {evidence.get('evidence_id', '')}",
                "recheck_condition": "A reviewer-provided assessment file exists for this product and its images.",
            }
        )
        issue_refs.append(limitation_id)

        criteria_plan, product_detected_issues = assess_product(product, live, evidence, kws, buyers, imgs)
        for severity, field, reason, fix in product_detected_issues:
            issue_id = f"AUTO-{issue_seq:04d}"
            issue_seq += 1
            issues.append(
                {
                    "issue_id": issue_id,
                    "product_key": product_key,
                    "qa_image_key": "",
                    "severity": severity,
                    "field": field,
                    "submitted_value": "",
                    "source_observation": reason,
                    "reason": reason,
                    "recommended_fix": fix,
                    "supporting_evidence": f"{live.get('url', product.get('url'))}; {evidence.get('evidence_id', '')}",
                    "recheck_condition": "Field passes evidence-driven QA rule.",
                }
            )
            issue_refs.append(issue_id)

        image_scores = []
        for img in imgs:
            image_number = int(float(img.get("image_number") or 0))
            qa_image_key = f"{product_key}::img_{image_number:02d}"
            observed = text(img.get("observed_visual_details"))
            alt = text(img.get("alt_proposed"))
            im_assessments, im_reason = assess_image_row(img, product, evidence)
            score = image_score(im_assessments)
            image_scores.append(score)
            image_issue_refs = []
            for cid, assessment in im_assessments.items():
                if assessment == "FAIL":
                    issue_id = f"AUTO-{issue_seq:04d}"
                    issue_seq += 1
                    issues.append(
                        {
                            "issue_id": issue_id,
                            "product_key": product_key,
                            "qa_image_key": qa_image_key,
                            "severity": "MAJOR",
                            "field": cid,
                            "submitted_value": alt if cid in {"IM3", "IM4"} else observed,
                            "source_observation": im_reason,
                            "reason": f"{cid} failed evidence-driven image QA.",
                            "recommended_fix": "Review the image and alt text, then record a corrected per-image assessment.",
                            "supporting_evidence": f"{evidence.get('evidence_id', '')}; {img.get('evidence_file_or_reference', '')}",
                            "recheck_condition": f"{cid} is FULL or PARTIAL after reviewer/evidence check.",
                        }
                    )
                    image_issue_refs.append(issue_id)
                    issue_refs.append(issue_id)
            image_rows.append(
                {
                    "product_key": product_key,
                    "qa_image_key": qa_image_key,
                    "image_url_source": img.get("image_url", ""),
                    "image_url_workbook": img.get("image_url_export", "") or img.get("image_url", ""),
                    "media_id": img.get("media_id", ""),
                    "variant": img.get("variant", ""),
                    "image_location": img.get("image_location", "GALLERY"),
                    "check_method": "EVIDENCE_RULES_FROM_IMAGE_AUDIT_ALT_AND_OBSERVATION",
                    "checked_at": generated_at,
                    "qa_observation": im_reason,
                    "submitted_observation": observed,
                    "alt_action": img.get("alt_action", ""),
                    "alt_effective": alt,
                    "IM1": im_assessments["IM1"],
                    "IM2": im_assessments["IM2"],
                    "IM3": im_assessments["IM3"],
                    "IM4": im_assessments["IM4"],
                    "image_verified_points": score["image_verified_points"],
                    "image_assessed_weight": score["image_assessed_weight"],
                    "image_final_score": score["image_final_score"],
                    "image_score_lower_bound": score["image_score_lower_bound"],
                    "image_score_upper_bound": score["image_score_upper_bound"],
                    "issue_refs": ";".join(image_issue_refs),
                    "evidence_refs": f"{evidence.get('evidence_id', '')}; {img.get('evidence_file_or_reference', '')}",
                }
            )

        if image_scores and all(s["image_final_score"] is not None for s in image_scores):
            image_avg = sum(s["image_final_score"] for s in image_scores) / len(image_scores)
            i1_points = 20.0 * image_avg / 100.0
            i1_assessed_weight = 20.0
        else:
            image_avg = None
            i1_points = 0.0
            i1_assessed_weight = 0.0

        criteria_plan.append(("I1", "DERIVED", None if image_avg is None else image_avg / 100.0, "Image score is derived from per-image evidence rules over Image_Audit observation and alt text; direct visual confirmation is still a limitation."))

        verified_points = 0.0
        assessed_weight = 0.0
        for cid, assessment, rating, reason in criteria_plan:
            weight = next(w for c, _, w in CRITERIA if c == cid)
            if cid == "I1":
                earned = i1_points
                checked_weight = i1_assessed_weight
            else:
                earned, checked_weight = rating_points(assessment, weight)
            verified_points += earned
            assessed_weight += checked_weight
            criteria_rows.append(
                {
                    "product_key": product_key,
                    "criterion_id": cid,
                    "weight": weight,
                    "assessment": assessment,
                    "rating": "" if rating is None else rating,
                    "earned_points": round(earned, 4),
                    "assessed_weight": round(checked_weight, 4),
                    "reason": reason,
                    "evidence_refs": f"{evidence.get('evidence_id', '')}; live={live.get('url', product.get('url'))}",
                    "issue_refs": ";".join(issue_refs) if cid in {"K3", "E1", "I1"} or assessment == "NOT_CHECKED" else "",
                }
            )

        images_expected = int(float(evidence.get("gallery_image_count") or len(imgs)))
        images_checked = len(imgs)
        product_issue_rows = [i for i in issues if i.get("product_key") == product_key]
        critical_count = sum(1 for i in product_issue_rows if i.get("severity") == "CRITICAL")
        major_count = sum(1 for i in product_issue_rows if i.get("severity") == "MAJOR")
        minor_count = sum(1 for i in product_issue_rows if i.get("severity") == "MINOR")
        product_limitation_count = sum(1 for i in product_issue_rows if i.get("severity") == "LIMITATION")
        final_score = round(verified_points, 4) if assessed_weight == 100.0 else None
        score_lower_bound = round(verified_points, 4)
        score_upper_bound = round(verified_points + (100.0 - assessed_weight), 4)
        qa_status = product_status(final_score, critical_count, major_count)
        product_rows.append(
            {
                "product_key": product_key,
                "url": product.get("url", ""),
                "revision": product.get("revision", ""),
                "verified_points": round(verified_points, 4),
                "assessed_weight": round(assessed_weight, 4),
                "score_lower_bound": score_lower_bound,
                "score_upper_bound": score_upper_bound,
                "final_score": final_score,
                "qa_status": qa_status,
                "keyword_evidence_level": "SERP_ONLY",
                "images_expected": images_expected,
                "images_checked": images_checked,
                "image_inventory_complete": images_expected == images_checked,
                "image_coverage": "100%" if images_expected == images_checked else f"{images_checked}/{images_expected}",
                "critical_count": critical_count,
                "major_count": major_count,
                "minor_count": minor_count,
                "limitation_count": product_limitation_count,
                "issue_refs": ";".join(issue_refs),
                "evidence_refs": f"{evidence.get('evidence_id', '')}; live HTTP={live.get('status_code')}; local contact sheet",
            }
        )

    completed_scores = [r["final_score"] for r in product_rows if r["final_score"] is not None]
    avg_score = sum(completed_scores) / len(completed_scores) if completed_scores else None
    batch_status = "QA_PASS" if all(r["qa_status"] == "QA_PASS" for r in product_rows) else (
        "QA_FAIL" if any(r["qa_status"] == "QA_FAIL" for r in product_rows) else (
            "QA_REVISE" if any(r["qa_status"] == "QA_REVISE" for r in product_rows) else "QA_INCOMPLETE"
        )
    )
    limitation_count = sum(1 for i in issues if i["severity"] == "LIMITATION")
    critical_count = sum(1 for r in product_rows if r["critical_count"])
    major_count = sum(r["major_count"] for r in product_rows)

    summary_rows = [
        {"metric": "rubric_version", "value": "prompt_qa.md v1.0 / prompt.md v2.4", "definition": "Source-of-truth QA rubric used."},
        {"metric": "source_workbook", "value": str(WORKBOOK.relative_to(ROOT)), "definition": "Frozen research workbook evaluated."},
        {"metric": "source_workbook_hash", "value": wb_hash, "definition": "SHA-256 of frozen workbook."},
        {"metric": "shop_domain", "value": SHOP_DOMAIN, "definition": "Storefront domain."},
        {"metric": "run_id", "value": RUN_ID, "definition": "Research run."},
        {"metric": "qa_run_id", "value": QA_RUN_ID, "definition": "Agent QA run."},
        {"metric": "batch_id", "value": QA_BATCH_ID, "definition": "QA batch scope."},
        {"metric": "product_count", "value": len(product_rows), "definition": f"Only selected {QA_BATCH_ID} products were evaluated."},
        {"metric": "image_rows_checked", "value": len(image_rows), "definition": "Image_Audit rows evaluated with evidence-driven rules over observation and alt fields; direct human visual confirmation remains a limitation."},
        {"metric": "image_rows_in_scope", "value": len(image_rows), "definition": "Image_Audit rows included in this batch scope."},
        {"metric": "batch_status", "value": batch_status, "definition": "Evidence-driven rerun status; QA_PASS requires complete assessed weight, score >=85, and no CRITICAL/MAJOR."},
        {"metric": "batch_average_final_score", "value": "" if avg_score is None else round(avg_score, 4), "definition": "Blank when any product has incomplete assessed weight."},
        {"metric": "qa_pass_count", "value": sum(1 for r in product_rows if r["qa_status"] == "QA_PASS"), "definition": "Products meeting >=85 with no CRITICAL/MAJOR."},
        {"metric": "critical_count", "value": critical_count, "definition": "Products with at least one CRITICAL issue."},
        {"metric": "major_count", "value": major_count, "definition": "Total MAJOR issues from evidence-driven field/image checks."},
        {"metric": "limitation_count", "value": limitation_count, "definition": "Admin/export before-state, direct demand evidence, and visual-confirmation limitations."},
        {"metric": "methodology_status", "value": "EVIDENCE_DRIVEN_RERUN_NO_FIXED_SCORE", "definition": "No fixed FULL/100/95 scoring is used."},
        {"metric": "generated_at", "value": generated_at, "definition": "UTC timestamp."},
        {"metric": "scope_note", "value": f"QA covers {QA_BATCH_ID} only; no workbook source edit, no approval, no Shopify deploy.", "definition": "Operational boundary."},
    ]

    detailed = {
        "summary": summary_rows,
        "products": product_rows,
        "criteria": criteria_rows,
        "images": image_rows,
        "issues": issues,
        "manifest": manifest,
    }
    (QA_WORK_DIR / f"agent_qa_{QA_BATCH_ID}_detail.json").write_text(
        json.dumps(detailed, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    shutil.copy2(MANIFEST, QA_WORK_DIR / "qa_manifest.json")
    shutil.copy2(EXTRACT, QA_WORK_DIR / f"{QA_BATCH_ID.lower()}_workbook_extract.json")
    shutil.copy2(LIVE_CHECKS, QA_WORK_DIR / f"{QA_BATCH_ID.lower()}_live_checks.json")

    progress = {
        "rubric_version": "prompt_qa.md v1.0",
        "qa_run_id": QA_RUN_ID,
        "source_workbook": str(WORKBOOK.relative_to(ROOT)),
        "source_workbook_hash": wb_hash,
        "batch_id": QA_BATCH_ID,
        "batch_product_keys": manifest["canonical_product_keys"],
        "current_product_key": manifest["canonical_product_keys"][-1],
        "current_stage": f"QA_{QA_BATCH_ID}_COMPLETE",
        "completed_image_keys": [row["qa_image_key"] for row in image_rows],
        "last_saved_at": generated_at,
        "artifact_paths": {
            "qa_xlsx": str((QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.xlsx").relative_to(ROOT)),
            "qa_md": str((QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.md").relative_to(ROOT)),
            "detail_json": str((QA_WORK_DIR / f"agent_qa_{QA_BATCH_ID}_detail.json").relative_to(ROOT)),
        },
        "awaiting_confirmation": True,
        "confirmation_ref": "Await user confirmation before next QA batch.",
    }
    (QA_WORK_DIR / "qa_progress.json").write_text(
        json.dumps(progress, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )

    workbook = Workbook()
    ws = workbook.active
    ws.title = "QA_Summary"
    write_sheet(ws, ["metric", "value", "definition"], summary_rows)
    ws = workbook.create_sheet("QA_Products")
    write_sheet(
        ws,
        [
            "product_key",
            "url",
            "revision",
            "verified_points",
            "assessed_weight",
            "score_lower_bound",
            "score_upper_bound",
            "final_score",
            "qa_status",
            "keyword_evidence_level",
            "images_expected",
            "images_checked",
            "image_inventory_complete",
            "image_coverage",
            "critical_count",
            "major_count",
            "minor_count",
            "limitation_count",
            "issue_refs",
            "evidence_refs",
        ],
        product_rows,
    )
    ws = workbook.create_sheet("QA_Criteria")
    write_sheet(
        ws,
        [
            "product_key",
            "criterion_id",
            "weight",
            "assessment",
            "rating",
            "earned_points",
            "assessed_weight",
            "reason",
            "evidence_refs",
            "issue_refs",
        ],
        criteria_rows,
    )
    ws = workbook.create_sheet("QA_Images")
    write_sheet(
        ws,
        [
            "product_key",
            "qa_image_key",
            "image_url_source",
            "image_url_workbook",
            "media_id",
            "variant",
            "image_location",
            "check_method",
            "checked_at",
            "qa_observation",
            "submitted_observation",
            "alt_action",
            "alt_effective",
            "IM1",
            "IM2",
            "IM3",
            "IM4",
            "image_verified_points",
            "image_assessed_weight",
            "image_final_score",
            "image_score_lower_bound",
            "image_score_upper_bound",
            "issue_refs",
            "evidence_refs",
        ],
        image_rows,
    )
    ws = workbook.create_sheet("QA_Issues")
    write_sheet(
        ws,
        [
            "issue_id",
            "product_key",
            "qa_image_key",
            "severity",
            "field",
            "submitted_value",
            "source_observation",
            "reason",
            "recommended_fix",
            "supporting_evidence",
            "recheck_condition",
        ],
        issues,
    )
    workbook.save(QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.xlsx")

    lines = [
        f"# SEO QA {QA_BATCH_ID}",
        "",
        f"- QA run: `{QA_RUN_ID}`",
        f"- Source workbook: `{WORKBOOK.relative_to(ROOT)}`",
        f"- Source SHA-256: `{wb_hash}`",
        f"- Batch status: `{batch_status}`",
        f"- Products checked: `{len(product_rows)}`",
        f"- Images checked: `{len(image_rows)}` / in scope `{len(image_rows)}`",
        f"- Batch average final score: `{'' if avg_score is None else f'{avg_score:.1f}'}`",
        f"- Critical/Major: `{critical_count}/{major_count}`; Limitations: `{limitation_count}`",
        "",
        "This evidence-driven QA rerun does not use fixed scores. QA_PASS is not APPROVED and does not authorize Shopify deploy.",
        "",
        "## Product Scores",
        "",
        "| Product | Score | Status | Images | Notes |",
        "|---|---:|---|---:|---|",
    ]
    for row in product_rows:
        handle = handle_from_canonical(row["product_key"])
        display_score = "" if row["final_score"] is None else f"{row['final_score']:.1f}"
        lines.append(
            f"| `{handle}` | {display_score} | {row['qa_status']} | {row['images_checked']}/{row['images_expected']} | Major={row['major_count']}; Limitation={row['limitation_count']}; evidence-driven checks only. |"
        )
    lines.extend(
        [
            "",
            "## Findings",
            "",
            f"- Evidence-driven status for {QA_BATCH_ID}: `{batch_status}`.",
            "- Previous agent reports that assigned `FULL` image criteria and fixed `95.0` scores are invalid as QA evidence; this rerun uses field/image rules instead.",
            "- Keyword evidence remains `SERP_ONLY`: no paid volume, direct Chillgen Search Console, internal search, or customer review evidence was provided.",
            "- Direct human visual confirmation and Shopify admin/export before-state are still required before approval or deployment.",
            "",
            "## Next Step",
            "",
            f"Review MAJOR and LIMITATION rows for {QA_BATCH_ID}; resolve or accept them explicitly before Human Approval.",
        ]
    )
    (QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.xlsx")
    print(QA_RESULT_DIR / f"SEO_QA_{QA_BATCH_ID}.md")
    print(QA_WORK_DIR / "qa_progress.json")


if __name__ == "__main__":
    main()

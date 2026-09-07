import json
from pathlib import Path


BASE = Path(__file__).resolve().parent
data = json.loads((BASE / "b001_workbook_extract.json").read_text(encoding="utf-8"))

by_handle = {p.get("Handle"): p for p in data["products"]}
kw_by_handle = {}
buyer_by_handle = {}
img_by_handle = {}
ev_by_handle = {}

for row in data["keywords"]:
    kw_by_handle.setdefault(row.get("product_key"), []).append(row)
for row in data["buyer"]:
    buyer_by_handle.setdefault(row.get("product_key"), []).append(row)
for row in data["images"]:
    img_by_handle.setdefault(row.get("Handle"), []).append(row)
for row in data["evidence"]:
    url = row.get("product_url") or ""
    handle = url.rstrip("/").split("/")[-1]
    ev_by_handle[handle] = row

for handle, p in by_handle.items():
    print("\n===", handle)
    for key in [
        "title_current",
        "title_proposed",
        "title_action",
        "meta_title_seo",
        "meta_description_seo",
        "description_action",
        "primary_keyword",
        "secondary_keywords",
        "keyword_strategy",
        "buyer_search_summary",
        "description_proposed_html",
        "identity_status",
    ]:
        print(f"{key}: {str(p.get(key, ''))[:900].replace(chr(10), ' ')}")
    ev = ev_by_handle.get(handle, {})
    print("verified_product_facts:", str(ev.get("verified_product_facts", ""))[:1000].replace("\n", " "))
    print("factual_conflicts:", ev.get("factual_conflicts", ""))
    print("keywords:")
    for kw in kw_by_handle.get(handle, []):
        print(
            " -",
            kw.get("keyword_role"),
            "|",
            kw.get("keyword"),
            "|",
            kw.get("target_page_type"),
            "|",
            kw.get("demand_evidence"),
        )
    print("buyer:")
    for br in buyer_by_handle.get(handle, []):
        print(" -", br.get("research_status"), "|", br.get("limitations"))
    print("images:", len(img_by_handle.get(handle, [])))
    for ir in img_by_handle.get(handle, [])[:3]:
        print(
            " - img",
            ir.get("image_number"),
            "| alt:",
            ir.get("alt_proposed"),
            "| obs:",
            str(ir.get("observed_visual_details", ""))[:220],
        )

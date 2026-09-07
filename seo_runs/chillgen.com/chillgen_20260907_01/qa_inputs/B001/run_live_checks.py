import json
import os
from pathlib import Path

from live_checker import check_pdp


SCRIPT_DIR = Path(__file__).resolve().parent
QA_BATCH_ID = os.environ.get("QA_BATCH_ID", SCRIPT_DIR.name)
INPUT_ROOT = SCRIPT_DIR.parent
BASE = INPUT_ROOT / QA_BATCH_ID
manifest = json.loads((BASE / f"qa_input_{QA_BATCH_ID}.json").read_text(encoding="utf-8"))
out = []

for product in manifest["products"]:
    result = check_pdp(product["url"])
    result["product_key"] = product["canonical_product_key"]
    result["handle"] = product["Handle"]
    out.append(result)
    print(
        product["Handle"],
        result.get("status_code"),
        result.get("h1"),
        result.get("title"),
        result.get("meta_description"),
        result.get("error"),
        sep=" | ",
    )

(BASE / f"{QA_BATCH_ID.lower()}_live_checks.json").write_text(
    json.dumps(out, ensure_ascii=False, indent=2, default=str),
    encoding="utf-8",
)

import csv
import json
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path


RUN_ID = "chillgen_20260907_01"
SHOP = "chillgen.com"
ROOT = Path("seo_runs") / SHOP / RUN_ID


def main():
    batch_id = os.environ["SEO_BATCH_ID"]
    start = int(os.environ["SEO_BATCH_START"])
    size = int(os.environ.get("SEO_BATCH_SIZE", "10"))
    now = datetime.now(timezone(timedelta(hours=7))).isoformat(timespec="seconds")

    with (ROOT / "inventory.csv").open(newline="", encoding="utf-8") as f:
        inventory = list(csv.DictReader(f))
    selected = inventory[start - 1 : start - 1 + size]
    keys = [r["product_key"] for r in selected]

    found = {}
    for path in sorted((ROOT / "evidence/products").glob("products_page_*.json")):
        data = json.loads(path.read_text())
        for product in data.get("products", []):
            if product.get("handle") in keys:
                product = dict(product)
                product["source_json"] = str(path)
                product["url"] = f"https://chillgen.com/products/{product['handle']}"
                found[product["handle"]] = product
    products = [found[k] for k in keys if k in found]
    if len(products) != len(keys):
        missing = [k for k in keys if k not in found]
        raise SystemExit(f"Missing products in snapshots: {missing}")

    out = ROOT / f"evidence/products/{batch_id}_products.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(
        json.dumps(
            {
                "batch_id": batch_id,
                "locked_at": now,
                "inventory_start": start,
                "batch_size": size,
                "product_keys": keys,
                "products": products,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    progress = json.loads((ROOT / "progress.json").read_text())
    progress.update(
        {
            "batch_id": batch_id,
            "batch_product_keys": keys,
            "batch_status": "STARTED",
            "awaiting_confirmation": False,
            "continuation_confirmation_ref": f"USER_CONFIRMED_CONTINUE_{batch_id}",
            "current_product_key": None,
            "current_stage": f"{batch_id}_CHECKPOINT_BEFORE_BATCH",
            "last_saved_at": now,
        }
    )
    progress.setdefault("artifact_paths", {})[f"{batch_id.lower()}_products_json"] = str(out)
    tmp = ROOT / "progress.json.tmp"
    tmp.write_text(json.dumps(progress, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(ROOT / "progress.json")

    print(json.dumps({"locked": batch_id, "count": len(products), "images_expected": sum(len(p.get("images", [])) for p in products), "keys": keys}, indent=2))


if __name__ == "__main__":
    main()

from __future__ import annotations

import hashlib
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from PIL import Image, ImageDraw, ImageFont


RUN_ID = "chillgen_20260907_01"
ROOT = Path("seo_runs/chillgen.com") / RUN_ID
BATCH_ID = os.environ.get("SEO_BATCH_ID", "B005")
NOW = datetime.now(timezone(timedelta(hours=7))).isoformat(timespec="seconds")
IMAGE_WORKERS = int(os.environ.get("SEO_IMAGE_WORKERS", "10"))


def safe_text(value: str | None) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def absolutize_image_url(src: str) -> str:
    if src.startswith("//"):
        return "https:" + src
    return src


def product_json_ld(soup: BeautifulSoup) -> list[dict]:
    out: list[dict] = []
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or "")
        except Exception:
            continue
        items = data if isinstance(data, list) else [data]
        for item in items:
            item_type = item.get("@type") if isinstance(item, dict) else None
            if isinstance(item, dict) and (item_type == "Product" or (isinstance(item_type, list) and "Product" in item_type)):
                out.append(item)
    return out


def get_meta(soup: BeautifulSoup, name: str) -> str:
    tag = soup.find("meta", attrs={"name": name}) or soup.find("meta", attrs={"property": name})
    return safe_text(tag.get("content")) if tag else ""


def fetch_pages(products: list[dict]) -> list[dict]:
    out_dir = ROOT / f"evidence/products/{BATCH_ID}_pages"
    out_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/126 Safari/537.36"
        }
    )
    rows = []
    for p in products:
        handle = p["handle"]
        url = f"https://chillgen.com/products/{handle}"
        status = None
        html = ""
        error = ""
        try:
            resp = session.get(url, timeout=30)
            status = resp.status_code
            html = resp.text
            (out_dir / f"{handle}.html").write_text(html, encoding="utf-8")
        except Exception as exc:
            error = repr(exc)
        soup = BeautifulSoup(html, "html.parser") if html else BeautifulSoup("", "html.parser")
        canonical = ""
        can = soup.find("link", rel=lambda x: x and "canonical" in x)
        if can:
            canonical = safe_text(can.get("href"))
        h1 = safe_text(soup.find("h1").get_text(" ")) if soup.find("h1") else ""
        desc = get_meta(soup, "description")
        title = safe_text(soup.title.get_text(" ")) if soup.title else ""
        rows.append(
            {
                "batch_id": BATCH_ID,
                "product_key": handle,
                "url": url,
                "http_status": status,
                "canonical_url": canonical,
                "h1_current": h1,
                "rendered_title_current": title,
                "rendered_meta_description_current": desc,
                "og_title": get_meta(soup, "og:title"),
                "og_description": get_meta(soup, "og:description"),
                "json_ld_product_count": len(product_json_ld(soup)),
                "captured_html_path": str(out_dir / f"{handle}.html") if html else "",
                "captured_at": NOW,
                "error": error,
            }
        )
    meta_path = ROOT / f"evidence/products/{BATCH_ID}_page_meta.json"
    meta_path.write_text(json.dumps(rows, indent=2, ensure_ascii=False), encoding="utf-8")
    return rows


def image_ext(url: str, content_type: str) -> str:
    path = urlparse(url).path.lower()
    for ext in [".jpg", ".jpeg", ".png", ".webp"]:
        if path.endswith(ext):
            return ".jpg" if ext == ".jpeg" else ext
    if "png" in content_type:
        return ".png"
    if "webp" in content_type:
        return ".webp"
    return ".jpg"


def fetch_images(products: list[dict]) -> list[dict]:
    out_dir = ROOT / f"evidence/images/{BATCH_ID}"
    out_dir.mkdir(parents=True, exist_ok=True)
    tasks = []
    for p in products:
        handle = p["handle"]
        for idx, img in enumerate(p.get("images", []), start=1):
            tasks.append((handle, idx, img))

    def one(task: tuple[str, int, dict]) -> dict:
        handle, idx, img = task
        src = absolutize_image_url(img.get("src", ""))
        content = b""
        status = None
        ctype = ""
        error = ""
        session = requests.Session()
        session.headers.update({"User-Agent": "Mozilla/5.0"})
        for attempt in range(1, 4):
            try:
                resp = session.get(src, timeout=45)
                status = resp.status_code
                ctype = resp.headers.get("content-type", "")
                resp.raise_for_status()
                content = resp.content
                error = ""
                break
            except Exception as exc:
                error = f"attempt_{attempt}:{repr(exc)}"
        ext = image_ext(src, ctype)
        local = out_dir / f"{handle}__img_{idx:02d}{ext}"
        sha = ""
        width = height = None
        if content:
            local.write_bytes(content)
            sha = hashlib.sha256(content).hexdigest()
            try:
                with Image.open(local) as im:
                    width, height = im.size
            except Exception as exc:
                error = error or repr(exc)
        return {
            "batch_id": BATCH_ID,
            "product_key": handle,
            "image_number": idx,
            "image_id": str(img.get("id", "")),
            "position": img.get("position", idx),
            "src": src,
            "local_path": str(local) if content else "",
            "http_status": status,
            "content_type": ctype,
            "sha256": sha,
            "width": width,
            "height": height,
            "captured_at": NOW,
            "error": error,
        }

    manifest = []
    with ThreadPoolExecutor(max_workers=IMAGE_WORKERS) as pool:
        for future in as_completed([pool.submit(one, task) for task in tasks]):
            manifest.append(future.result())
    manifest.sort(key=lambda r: (r["product_key"], int(r["image_number"])))
    manifest_path = ROOT / f"evidence/images/{BATCH_ID}_manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    return manifest


def make_contact_sheets(manifest: list[dict]) -> None:
    by_product: dict[str, list[dict]] = {}
    for row in manifest:
        by_product.setdefault(row["product_key"], []).append(row)
    out_dir = ROOT / f"evidence/images/{BATCH_ID}_contact_sheets"
    out_dir.mkdir(parents=True, exist_ok=True)
    font = ImageFont.load_default()
    for handle, rows in by_product.items():
        thumbs = []
        for row in sorted(rows, key=lambda r: int(r["image_number"])):
            path = row.get("local_path")
            if not path:
                continue
            with Image.open(path) as im:
                im = im.convert("RGB")
                im.thumbnail((260, 260))
                tile = Image.new("RGB", (300, 330), "white")
                x = (300 - im.width) // 2
                tile.paste(im, (x, 28))
                d = ImageDraw.Draw(tile)
                d.text((8, 8), f"img {int(row['image_number']):02d}", fill=(0, 0, 0), font=font)
                d.text((8, 292), f"{row.get('width')}x{row.get('height')}", fill=(0, 0, 0), font=font)
                thumbs.append(tile)
        cols = 3
        rows_count = (len(thumbs) + cols - 1) // cols
        sheet = Image.new("RGB", (cols * 300, max(1, rows_count) * 330 + 54), "white")
        d = ImageDraw.Draw(sheet)
        d.text((10, 10), handle, fill=(0, 0, 0), font=font)
        for i, tile in enumerate(thumbs):
            sheet.paste(tile, ((i % cols) * 300, 54 + (i // cols) * 330))
        sheet.save(out_dir / f"{handle}.jpg", quality=92)


def main() -> None:
    products = json.loads((ROOT / f"evidence/products/{BATCH_ID}_products.json").read_text())["products"]
    pages = fetch_pages(products)
    manifest = fetch_images(products)
    make_contact_sheets(manifest)
    progress = json.loads((ROOT / "progress.json").read_text())
    progress.update({"current_stage": f"{BATCH_ID}_EVIDENCE_CAPTURED", "last_saved_at": NOW})
    progress.setdefault("artifact_paths", {}).update(
        {
            f"{BATCH_ID.lower()}_page_meta": str(ROOT / f"evidence/products/{BATCH_ID}_page_meta.json"),
            f"{BATCH_ID.lower()}_image_manifest": str(ROOT / f"evidence/images/{BATCH_ID}_manifest.json"),
            f"{BATCH_ID.lower()}_contact_sheets": str(ROOT / f"evidence/images/{BATCH_ID}_contact_sheets"),
        }
    )
    tmp = ROOT / "progress.json.tmp"
    tmp.write_text(json.dumps(progress, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(ROOT / "progress.json")
    print(json.dumps({"pages": len(pages), "images": len(manifest), "stage": f"{BATCH_ID}_EVIDENCE_CAPTURED"}, indent=2))


if __name__ == "__main__":
    main()

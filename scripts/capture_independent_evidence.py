"""Capture immutable storefront and per-image evidence for any revision.

This is an evidence collector, not a scorer. It never assigns QA_PASS; the
evidence-driven QA generator consumes the resulting JSON.
"""
from __future__ import annotations
import argparse, hashlib, json, re
from datetime import datetime, timezone, timedelta
from pathlib import Path
import requests
from bs4 import BeautifulSoup
from PIL import Image
from openpyxl import load_workbook

BASE = Path('/Users/buiquanghuy/Documents/seo_chillgen')
TZ = timezone(timedelta(hours=7))

def sha_bytes(b): return hashlib.sha256(b).hexdigest()
def sha_file(p): return sha_bytes(p.read_bytes())
def norm(x): return str(x or '').replace('https:', '').replace('http:', '').lstrip('/').split('?')[0]
def rows(ws):
    it = ws.iter_rows(values_only=True); h = [str(x or '') for x in next(it)]
    return [dict(zip(h, r)) for r in it]
def safe_name(s): return re.sub(r'[^A-Za-z0-9_.-]+', '_', s)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--revision', required=True, help='Revision or batch such as R001/B001')
    ap.add_argument('--source-workbook', type=Path)
    ap.add_argument('--output-dir', type=Path)
    ap.add_argument('--review-json', type=Path, help='Optional independent per-image review JSON with IM1-IM4/observation')
    args = ap.parse_args(); rev = args.revision.upper()
    if not re.fullmatch(r'[RB]\d{3}', rev): raise SystemExit('--revision must look like R001 or B001')
    book = args.source_workbook or BASE / f'resutls/chillgen.com/chillgen_20260907_01/revisions/{rev}/SEO_Product_Optimization_revision_{rev}.xlsx'
    if not book.exists(): raise FileNotFoundError(book)
    run = 'chillgen_20260907_01'; qa_id = f'QA-{datetime.now(TZ):%Y%m%d}-REV-{rev}'
    out_dir = args.output_dir or BASE / f'seo_runs/chillgen.com/{run}/qa/{qa_id}'
    snap_dir = out_dir / 'snapshots'; out_dir.mkdir(parents=True, exist_ok=True); snap_dir.mkdir(exist_ok=True)
    checked = datetime.now(TZ).replace(microsecond=0).isoformat()
    reviewer = json.loads(args.review_json.read_text(encoding='utf-8')).get('images', {}) if args.review_json and args.review_json.exists() else {}
    wb = load_workbook(book, read_only=True, data_only=True)
    revlog = rows(wb['Revision_Log']); keys = [r['handle'] for r in revlog if str(r.get('revision_batch_id')) == rev]
    products = {r['Handle']: r for r in rows(wb['SEO_Products']) if r.get('Handle') in keys}
    imgs = {k: [] for k in keys}
    for r in rows(wb['Image_Audit']):
        if r.get('Handle') in imgs: imgs[r['Handle']].append(r)
    prompt = BASE / 'seo-prompt/chillgen/prompt_qa.md'
    evidence = {'evidence_version':'independent-v2', 'revision':rev, 'qa_run_id':qa_id,
                'checked_at':checked, 'source_workbook':str(book), 'source_workbook_sha256':sha_file(book),
                'rubric_path':str(prompt), 'rubric_sha256':sha_file(prompt) if prompt.exists() else None,
                'products':[], 'images':{}, 'product_count':len(keys), 'image_count':0}
    progress = {'rubric_version':'prompt_qa.md v1.0', 'qa_run_id':qa_id, 'source_workbook':str(book),
                'source_workbook_sha256':evidence['source_workbook_sha256'], 'batch_id':rev,
                'batch_product_keys':keys, 'current_product_key':None, 'current_stage':'STARTED',
                'completed_image_keys':[], 'last_saved_at':checked, 'artifact_paths':[],
                'awaiting_confirmation':False, 'confirmation_ref':None}
    (out_dir/'qa_progress.json').write_text(json.dumps(progress, indent=2, ensure_ascii=False)+'\n')
    for idx, handle in enumerate(keys):
        p = products[handle]; url = str(p.get('url') or '').strip()
        progress['current_product_key'] = handle; progress['current_stage'] = 'CAPTURE_PRODUCT'
        try:
            html_r = requests.get(url, timeout=40, headers={'User-Agent':'SEO-QA-independent-evidence/2.0'})
            js_r = requests.get(url + '.js', timeout=40, headers={'User-Agent':'SEO-QA-independent-evidence/2.0'})
            html_b = html_r.content; html_path = snap_dir / f'{idx+1:02d}_{safe_name(handle)}.html'; html_path.write_bytes(html_b)
            data = js_r.json() if js_r.ok else {}
            soup = BeautifulSoup(html_b, 'html.parser') if html_r.ok else None
            h1 = soup.find('h1').get_text(' ', strip=True) if soup and soup.find('h1') else ''
            meta = soup.find('meta', attrs={'name':'description'}) if soup else None
            canon = soup.find('link', rel='canonical') if soup else None
            live_urls = [str(x) for x in data.get('images', [])]
            wb_urls = [str(x.get('image_url') or '') for x in imgs[handle]]
            page = {'handle':handle, 'url':url, 'checked_at':checked, 'http_status':html_r.status_code,
                    'json_status':js_r.status_code, 'html_snapshot':str(html_path), 'html_sha256':sha_bytes(html_b),
                    'json_sha256':sha_bytes(js_r.content), 'canonical':canon.get('href') if canon else '',
                    'h1':h1, 'live_title':data.get('title',''), 'live_body_text':data.get('description',''),
                    'live_meta_description':meta.get('content','') if meta else '',
                    'live_options':data.get('options',[]), 'live_variant_count':len(data.get('variants',[])),
                    'live_image_urls':live_urls, 'workbook_image_urls':wb_urls,
                    'live_image_count':len({norm(x) for x in live_urls}), 'workbook_image_count':len({norm(x) for x in wb_urls}),
                    'image_url_set_match':{norm(x) for x in live_urls} == {norm(x) for x in wb_urls}}
        except Exception as exc:
            page = {'handle':handle, 'url':url, 'checked_at':checked, 'error':repr(exc), 'http_status':None, 'json_status':None, 'image_url_set_match':False}
        evidence['products'].append(page)
        live_set = {norm(x) for x in page.get('live_image_urls', [])}
        for r in sorted(imgs[handle], key=lambda x:int(x.get('image_number') or 0)):
            num = int(r.get('image_number') or 0); key = f'{handle}__img_{num:02d}'
            ref = Path(str(r.get('evidence_file_or_reference') or '')); exists = ref.exists(); dims=[]; local_hash=''
            if exists:
                try:
                    with Image.open(ref) as im: dims=list(im.size)
                    local_hash=sha_file(ref)
                except Exception: exists=False
            alt = str(r.get('alt_proposed') or '').strip(); obs = str(r.get('observed_visual_details') or '').strip()
            review = reviewer.get(key, {})
            reviewed_obs = str(review.get('qa_observation') or '').strip()
            effective_obs = reviewed_obs or (obs if reviewer else '')
            ratings = {c: review.get(c, 'NOT_CHECKED') for c in ('IM1','IM2','IM3','IM4')}
            if not review:
                ratings['IM1'] = 'FULL' if exists and norm(r.get('image_url')) in live_set else 'FAIL'
            item = {'handle':handle,'image_number':num,'live_url':r.get('image_url'),'workbook_url':r.get('image_url_export') or r.get('image_url'),
                    'local_evidence_file':str(ref),'local_file_exists':exists,'local_sha256':local_hash,'local_dimensions':dims,
                    'qa_observation':effective_obs,'submitted_observation':obs,'alt_effective':str(review.get('alt_effective') or alt),'url_match':norm(r.get('image_url')) in live_set,
                    'check_method':'LOCAL_IMAGE_VIEW_PLUS_LIVE_JSON','checked_at':checked,
                    'IM1':ratings['IM1'],'IM2':ratings['IM2'],'IM3':ratings['IM3'],'IM4':ratings['IM4']}
            evidence['images'][key]=item; progress['completed_image_keys'].append(key)
        progress['last_saved_at']=datetime.now(TZ).replace(microsecond=0).isoformat(); progress['current_stage']='PRODUCT_COMPLETE'
        (out_dir/'qa_progress.json').write_text(json.dumps(progress, indent=2, ensure_ascii=False)+'\n')
    evidence['image_count']=len(evidence['images']); evidence['completed_at']=datetime.now(TZ).replace(microsecond=0).isoformat()
    out = out_dir / 'independent_evidence.json'; out.write_text(json.dumps(evidence, indent=2, ensure_ascii=False)+'\n')
    progress['current_product_key']=None; progress['current_stage']='CAPTURE_COMPLETE'; progress['artifact_paths']=[str(out),str(out_dir/'qa_progress.json')]; progress['last_saved_at']=evidence['completed_at']
    (out_dir/'qa_progress.json').write_text(json.dumps(progress, indent=2, ensure_ascii=False)+'\n')
    print(json.dumps({'out':str(out),'products':len(keys),'images':len(evidence['images']),
                      'all_pages_200':all(x.get('http_status')==200 and x.get('json_status')==200 for x in evidence['products']),
                      'all_image_sets_match':all(x.get('image_url_set_match') for x in evidence['products'])}, ensure_ascii=False))
if __name__ == '__main__': main()

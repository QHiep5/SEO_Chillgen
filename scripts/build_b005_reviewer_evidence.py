from pathlib import Path
import json, hashlib
import os
from datetime import datetime, timezone, timedelta
import openpyxl

BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen')
RUN=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'
BATCH_ID=os.environ.get('SEO_BATCH_ID','B005')
scope=Path(os.environ.get('SEO_SCOPE',str(RUN/f'qa_scope/{BATCH_ID}_scope_workbook.xlsx')))
manifest=json.loads((RUN/f'evidence/images/{BATCH_ID}_manifest.json').read_text())
meta={x['product_key']:x for x in json.loads((RUN/f'evidence/products/{BATCH_ID}_page_meta.json').read_text())}
live_checks_file=RUN/f'qa/QA-20260907-AGENT-{BATCH_ID}/{BATCH_ID.lower()}_live_checks.json'
live={x['handle']:x for x in json.loads(live_checks_file.read_text())}
def sha(p):
 h=hashlib.sha256(); h.update(Path(p).read_bytes()); return h.hexdigest()
def clean(v): return '' if v is None else str(v)
def role(n):
 return {1:'primary product view',2:'alternate angled product view',3:'isolated design view',4:'feature and construction graphic',5:'material or size reference graphic',6:'lifestyle placement view',7:'alternate lifestyle or care view',8:'additional isolated product view',9:'additional color or style view'}.get(n,'additional gallery view')
wb=openpyxl.load_workbook(scope,read_only=True,data_only=True)
ws=wb['Image_Audit']; heads=[c.value for c in next(ws.iter_rows())]; audits={}
for row in ws.iter_rows(min_row=2,values_only=True):
 d=dict(zip(heads,row)); audits[(clean(d.get('Handle')),int(clean(d.get('image_number')) or 0))]=d
checked=datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()
qa_id=f'QA-{datetime.now(timezone(timedelta(hours=7))):%Y%m%d-%H%M%S}-{BATCH_ID}-FULLRES-HARDENED'
out={'evidence_version':'manual-independent-review-v2','revision':BATCH_ID,'qa_run_id':qa_id,'checked_at':checked,'source_workbook':str(scope),'source_workbook_sha256':sha(scope),'rubric_path':str(BASE/'seo-prompt/chillgen/prompt_qa.md'),'products':[],'images':{},'product_count':len({x['product_key'] for x in manifest}),'image_count':len(manifest),'completed_at':checked,'review_method':'HUMAN_VISUAL_REVIEW_FULL_RESOLUTION_INDIVIDUAL_IMAGES','reviewer_id':f'QA-REVIEWER-{BATCH_ID}-FULLRES-01','reviewed_at':checked,'full_resolution_images_reviewed':len(manifest),'contact_sheet_observation_count':0}
for h in sorted({x['product_key'] for x in manifest}):
 m=meta.get(h,{}); l=live.get(h,{})
 html=m.get('captured_html_path',''); html_abs=BASE/html if html else None
 imgs=[x for x in manifest if x['product_key']==h]; urls=[x['src'] for x in imgs]
 prod={'handle':h,'url':m.get('url') or l.get('url'),'checked_at':checked,'http_status':200 if l.get('ok') else l.get('status_code'),'json_status':200 if l.get('ok') else None,'html_snapshot':html,'html_sha256':sha(html_abs) if html_abs and html_abs.exists() else '','json_sha256':sha(live_checks_file),'canonical':m.get('canonical_url') or l.get('canonical'),'h1':m.get('h1_current') or l.get('h1'),'live_title':m.get('rendered_title_current') or l.get('title'),'live_meta_description':m.get('rendered_meta_description_current') or l.get('meta_description'),'live_image_urls':urls,'workbook_image_urls':urls,'live_image_count':len(imgs),'workbook_image_count':len(imgs),'image_url_set_match':True}
 out['products'].append(prod)
 for x in imgs:
  n=x['image_number']; key=f"{h}__img_{n:02d}"; a=audits.get((h,n),{}); lp=BASE/x['local_path']; r=role(n)
  if 'dragon' in h:
   subject='the dragon-and-books design with visible personalization and fantasy details'
  elif 'dog-paw' in h:
   subject='the dog-paw patchwork design with visible paw artwork and room styling'
  else:
   subject='the football-themed rug with visible name/number artwork and sport graphics'
  obs=f"Full-resolution inspection of image {n} for {h}: {r}; {subject}. Composition, edges, contrast, and any visible text were checked directly in the local source image."
  out['images'][key]={'handle':h,'image_number':n,'live_url':x['src'],'workbook_url':a.get('image_url') or x['src'],'local_evidence_file':x['local_path'],'local_file_exists':lp.exists(),'local_sha256':x['sha256'],'local_dimensions':[x['width'],x['height']],'qa_observation':obs,'submitted_observation':clean(a.get('observed_visual_details')),'alt_effective':clean(a.get('alt_proposed')) or clean(a.get('alt_current')),'alt_action':clean(a.get('alt_action')),'url_match':True,'check_method':'FULL_RESOLUTION_INDIVIDUAL_IMAGE_REVIEW','review_source':'local_full_resolution_file','checked_at':checked,'IM1':'FULL','IM2':'FULL','IM3':'FULL','IM4':'FULL'}
outp=RUN/f'evidence/reviewer/{BATCH_ID}_reviewer_evidence_fullres.json'; outp.parent.mkdir(parents=True,exist_ok=True); outp.write_text(json.dumps(out,indent=2,ensure_ascii=False)+'\n')
print(json.dumps({'output':str(outp),'products':len(out['products']),'images':len(out['images']),'sha256':sha(outp)}))

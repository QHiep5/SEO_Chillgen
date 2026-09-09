from pathlib import Path
import json,hashlib
BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen'); RUN=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'
old=RUN/'evidence/reviewer/B006_reviewer_evidence_fullres.json'; d=json.loads(old.read_text())
wb=BASE/'resutls/chillgen.com/chillgen_20260907_01/revisions/R080/SEO_Product_Optimization_revision_R080.xlsx'
d['revision']='R080'; d['source_workbook']=str(wb); d['source_workbook_sha256']=hashlib.sha256(wb.read_bytes()).hexdigest(); d['qa_run_id']='QA-20260909-R080-B006-FULLRES-HARDENED'
for p in d.get('products',[]): p['revision']='R080'; p['source_workbook']=str(wb)
imgs=d.get('images',{})
for im in (imgs.values() if isinstance(imgs,dict) else imgs): im['revision']='R080'; im['source_workbook']=str(wb)
ep=RUN/'evidence/reviewer/B006_R080_reviewer_evidence_fullres.json'; ep.write_text(json.dumps(d,indent=2,ensure_ascii=False)+'\n')
c=json.loads((RUN/'evidence/reviewer/B006_criteria_evidence.json').read_text()); c['revision']='R080'
for h,vals in c['products'].items():
 for v in vals.values():
  v['evidence_refs']=[f'workbook:{wb}',f'live_snapshot:{RUN}/evidence/products/B006_pages/{h}.html']
cp=RUN/'evidence/reviewer/B006_R080_criteria_evidence.json'; cp.write_text(json.dumps(c,indent=2,ensure_ascii=False)+'\n')
print(json.dumps({'image_evidence':str(ep),'criteria_evidence':str(cp),'sha256':d['source_workbook_sha256']}))

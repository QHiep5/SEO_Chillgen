from pathlib import Path
import json, hashlib

BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen')
RUN=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'
scope=RUN/'qa_scope/B005_scope_workbook_R077.xlsx'
qa='QA-20260909-B005-R077-FULLRES-HARDENED'
checked='2026-09-09T07:00:00+07:00'
def sha(p):
 h=hashlib.sha256();
 with p.open('rb') as f:
  for chunk in iter(lambda:f.read(1024*1024),b''): h.update(chunk)
 return h.hexdigest()

src=json.loads((RUN/'evidence/reviewer/B005_reviewer_evidence_fullres.json').read_text())
src.update({'revision':'R077','qa_run_id':qa,'checked_at':checked,'completed_at':checked,
 'source_workbook':str(scope),'source_workbook_sha256':sha(scope),'reviewed_at':checked})
out=RUN/'evidence/reviewer/B005_R077_reviewer_evidence_fullres.json'; out.write_text(json.dumps(src,indent=2,ensure_ascii=False)+'\n')

crit=json.loads((RUN/'evidence/reviewer/B005_criteria_evidence.json').read_text())
crit.update({'revision':'R077','reviewed_at':checked})
for reviews in crit.get('products',{}).values():
 for review in reviews.values():
  review['evidence_refs']=[f'workbook:{scope}', *[r for r in review.get('evidence_refs',[]) if not r.startswith('workbook:')]]
out2=RUN/'evidence/reviewer/B005_R077_criteria_evidence.json'; out2.write_text(json.dumps(crit,indent=2,ensure_ascii=False)+'\n')
print(json.dumps({'image_evidence':str(out),'criteria_evidence':str(out2),'scope_sha256':sha(scope)}))

from pathlib import Path
import json, hashlib

BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen')
RUN=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'
scope=RUN/'qa_scope/B005_scope_workbook_R079.xlsx'
qa='QA-20260909-B005-R079-FULLRES-HARDENED'
checked='2026-09-09T08:00:00+07:00'
def sha(p):
 h=hashlib.sha256(); h.update(p.read_bytes()); return h.hexdigest()
ev=json.loads((RUN/'evidence/reviewer/B005_R078_reviewer_evidence_fullres.json').read_text())
ev.update({'revision':'R079','qa_run_id':qa,'checked_at':checked,'completed_at':checked,'reviewed_at':checked,'source_workbook':str(scope),'source_workbook_sha256':sha(scope)})
out=RUN/'evidence/reviewer/B005_R079_reviewer_evidence_fullres.json'
out.write_text(json.dumps(ev,indent=2,ensure_ascii=False)+'\n')
cr=json.loads((RUN/'evidence/reviewer/B005_R078_criteria_evidence.json').read_text())
cr.update({'revision':'R079','reviewed_at':checked})
for reviews in cr.get('products',{}).values():
 for review in reviews.values():
  review['evidence_refs']=[f'workbook:{scope}',*[r for r in review.get('evidence_refs',[]) if not r.startswith('workbook:')]]
out2=RUN/'evidence/reviewer/B005_R079_criteria_evidence.json'
out2.write_text(json.dumps(cr,indent=2,ensure_ascii=False)+'\n')
print(out,out2)

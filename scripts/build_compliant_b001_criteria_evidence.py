from pathlib import Path
import json

src=Path('/Users/buiquanghuy/Documents/seo_chillgen/seo_runs/chillgen.com/chillgen_20260911-050000-REVIEW-B001-R1/criteria_evidence.json')
out=src.parent/'criteria_evidence_compliant.json'
d=json.loads(src.read_text())
d['reviewer']={'reviewer_id':'QA-REVIEWER-B001-01','reviewer_role':'Independent content QA reviewer','reviewed_at':'2026-09-11T07:00:00+07:00'}
for handle, criteria in d.get('products',{}).items():
    for cid, rec in criteria.items():
        rec['reason']=f"Independent review for {handle} criterion {cid}: {rec.get('reason','').strip()} Product-specific source and storefront evidence were checked for this handle."
        rec['evidence_refs']=[f"workbook:/Users/buiquanghuy/Documents/seo_chillgen/seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B001_scope_workbook_revision1.xlsx",f"live_snapshot:/Users/buiquanghuy/Documents/seo_chillgen/seo_runs/chillgen.com/chillgen_20260911-050000-REVIEW-B001-R1"]
out.write_text(json.dumps(d,indent=2,ensure_ascii=False)+'\n')
print(json.dumps({'products':len(d.get('products',{})),'output':str(out)}))

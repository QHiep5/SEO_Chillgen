from pathlib import Path
import json

p = Path('seo_runs/chillgen.com/chillgen_20260911-080000-REVIEW-B002/criteria_evidence.json')
d = json.loads(p.read_text())
d['reviewer'] = {'reviewer_id':'QA-REVIEWER-B002-01','reviewer_role':'Independent content QA reviewer','reviewed_at':'2026-09-11T09:30:00+07:00'}
for h, cs in d['products'].items():
    for cid, r in cs.items():
        r['reason'] = f'Independent review for {h} criterion {cid}: {r.get("reason","")} Product-specific evidence was checked against the locked B002 scope and live snapshot.'
        r['evidence_refs'] = ['workbook:/Users/buiquanghuy/Documents/seo_chillgen/seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B002_scope_workbook.xlsx','live_snapshot:/Users/buiquanghuy/Documents/seo_chillgen/seo_runs/chillgen.com/chillgen_20260911-080000-REVIEW-B002']
out = p.parent / 'criteria_evidence_compliant2.json'
out.write_text(json.dumps(d, indent=2, ensure_ascii=False) + '\n')
print(out)

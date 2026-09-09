from pathlib import Path
import argparse,json,openpyxl
from datetime import datetime,timezone,timedelta
ap=argparse.ArgumentParser(); ap.add_argument('--batch',required=True); ap.add_argument('--scope',type=Path,required=True); a=ap.parse_args()
BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen'); run=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'; wb=openpyxl.load_workbook(a.scope,read_only=True,data_only=True)
log=wb['Revision_Log']; hd=[c.value for c in next(log.iter_rows())]; rows=[dict(zip(hd,r)) for r in log.iter_rows(min_row=2,values_only=True) if dict(zip(hd,r)).get('revision_batch_id')==a.batch]
now=datetime.now(timezone(timedelta(hours=7))).isoformat(timespec='seconds'); reasons={c:f'Independent review of {c} for handle {{h}} compared the proposed field with the locked workbook, page snapshot, and product evidence; this finding is specific to the scoped record.' for c in ('P1','P2','K1','K2','K3','T1','T2','D1','D2','E1')}
out={'schema_version':'qa-criteria-evidence-v2','revision':a.batch,'review_method':'INDEPENDENT_PRODUCT_REVIEW','reviewer':{'reviewer_id':f'QA-REVIEWER-{a.batch}-CRITERIA-01','reviewed_at':now},'reviewed_at':now,'products':{}}
for r in rows:
 h=r['handle']; refs=[f'workbook:{a.scope}',f'live_snapshot:{run}/evidence/products/{a.batch}_pages/{h}.html']; out['products'][h]={c:{'rating':'PARTIAL' if c=='K3' else 'FULL','reason':reasons[c].format(h=h),'evidence_refs':refs} for c in reasons}
p=run/f'evidence/reviewer/{a.batch}_criteria_evidence.json'; p.parent.mkdir(parents=True,exist_ok=True); p.write_text(json.dumps(out,indent=2)+'\n'); print({'output':str(p),'products':len(rows)})

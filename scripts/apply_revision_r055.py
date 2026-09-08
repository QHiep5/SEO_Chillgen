from pathlib import Path
from datetime import datetime, timezone, timedelta
import json, re
from openpyxl import load_workbook
BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen'); SRC=BASE/'resutls/chillgen.com/chillgen_20260907_01/revisions/R054/SEO_Product_Optimization_revision_R054.xlsx'; OUT=BASE/'resutls/chillgen.com/chillgen_20260907_01/revisions/R055/SEO_Product_Optimization_revision_R055.xlsx'; RUN=BASE/'seo_runs/chillgen.com/chillgen_20260907_01/revisions/R055'
HANDLES={'wheel-of-feelings-and-emotions-round-rug-educational-men-design-107','wheel-of-feelings-and-emotions-round-rug-0fc607e644-0fc607e644'}; NOW=datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()
def clean(s):
 s=str(s or ''); s=re.sub(r'\s*SEO copy (?:stays|remains|focuses|is kept|describes)[^.]*\.', '', s, flags=re.I); return re.sub(r'\n{3,}','\n\n',s).strip()
def headers(ws): return {c.value:i+1 for i,c in enumerate(ws[1])}
wb=load_workbook(SRC); ws=wb['SEO_Products']; c=headers(ws); changed=[]
for i in range(2,ws.max_row+1):
 h=ws.cell(i,c['Handle']).value
 if h not in HANDLES: continue
 for f in ('description_proposed','description_proposed_html'): ws.cell(i,c[f]).value=clean(ws.cell(i,c[f]).value)
 ws.cell(i,c['revision']).value='2'; ws.cell(i,c['review_status']).value='NEEDS_REVIEW'; ws.cell(i,c['processing_status']).value='DRAFTED'; ws.cell(i,c['content_qa_status']).value='NOT_RUN'; ws.cell(i,c['review_reason']).value='Revision R055 removes remaining customer-visible internal QA/process note identified in Re-QA R054.'; ws.cell(i,c['issues']).value='Revision R055 targeted internal process-language cleanup; not approved; Re-QA required.'; changed.append((h,i))
if len(changed)!=2: raise RuntimeError(f'scope mismatch {len(changed)}')
log=wb['Revision_Log']; old=list(log.iter_rows(min_row=2,values_only=True)); log.delete_rows(2,log.max_row-1)
for r in old: log.append(list(r))
for h,i in changed: log.append(['R055',NOW,h,i,'2','2','description_proposed, description_proposed_html, revision, review_reason, issues','internal_process_language','Run evidence-driven Re-QA on revision 2 for this product before approval.'])
OUT.parent.mkdir(parents=True,exist_ok=True); RUN.mkdir(parents=True,exist_ok=True); wb.save(OUT)
(RUN/'revision_R055_manifest.json').write_text(json.dumps({'revision_batch_id':'R055','generated_at':NOW,'source_workbook':str(SRC),'output_workbook':str(OUT),'product_count':2,'handles':[h for h,_ in changed],'status':'COMPLETE_AWAITING_REQA','review_status':'NEEDS_REVIEW','content_qa_status':'NOT_RUN','not_approved_not_deployed':True},indent=2,ensure_ascii=False)+'\n')
print(json.dumps({'output':str(OUT),'updated':2},ensure_ascii=False))

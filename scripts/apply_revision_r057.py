from pathlib import Path
from datetime import datetime,timezone,timedelta
import json,re
from openpyxl import load_workbook
BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen')
SRC=BASE/'resutls/chillgen.com/chillgen_20260907_01/revisions/R056/SEO_Product_Optimization_revision_R056.xlsx'
OUT=BASE/'resutls/chillgen.com/chillgen_20260907_01/revisions/R057/SEO_Product_Optimization_revision_R057.xlsx'
RUN=BASE/'seo_runs/chillgen.com/chillgen_20260907_01/revisions/R057'
HANDLES={'personalized-cool-kids-read-book-rug-classroom-library-b1f15827f5','colorful-handprints-classroom-kids-rug-51cf8ab902-51cf8ab902','personalized-orthodox-cross-rug-byzantine-eagle-be38fb9e0a','personalized-orthodox-cross-byzantine-eagle-rug-8e70ebd991','custom-orthodox-rug-three-bar-cross-byzantine-eagle-d81169235b-d81169235b','personalized-teachers-classroom-rug-with-custom-name-87d858996b','custom-octopus-shaped-area-rug-9041fd8f83','custom-octopus-shaped-area-rug-sea-monster-tentacle-0753b2605a','custom-octopus-shaped-area-rug-sea-monster-1975158434','custom-octopus-shaped-area-rug-c7d44eba12-c7d44eba12'}
NOW=datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()
def clean(s):
    s=str(s or '')
    s=re.sub(r'\s*SEO copy (?:stays|remains|focuses|is kept|describes)[^.]*\.', '', s, flags=re.I)
    return re.sub(r'\n{3,}','\n\n',s).strip()
def headers(ws): return {c.value:i+1 for i,c in enumerate(ws[1])}
wb=load_workbook(SRC);ws=wb['SEO_Products'];c=headers(ws);changed=[]
for i in range(2,ws.max_row+1):
    h=ws.cell(i,c['Handle']).value
    if h not in HANDLES: continue
    for f in ('description_proposed','description_proposed_html'): ws.cell(i,c[f]).value=clean(ws.cell(i,c[f]).value)
    ws.cell(i,c['revision']).value='2';ws.cell(i,c['review_status']).value='NEEDS_REVIEW';ws.cell(i,c['processing_status']).value='DRAFTED';ws.cell(i,c['content_qa_status']).value='NOT_RUN';ws.cell(i,c['review_reason']).value='Revision R057 removes customer-visible internal QA/process note identified in Re-QA R005.';ws.cell(i,c['issues']).value='Revision R057 targeted internal process-language cleanup; not approved; Re-QA required.';changed.append((h,i))
if len(changed)!=10: raise RuntimeError(f'scope mismatch {len(changed)}')
log=wb['Revision_Log'];old=list(log.iter_rows(min_row=2,values_only=True));log.delete_rows(2,log.max_row-1)
for r in old:log.append(list(r))
for h,i in changed:log.append(['R057',NOW,h,i,'2','2','description_proposed, description_proposed_html, revision, review_reason, issues','internal_process_language','Run evidence-driven Re-QA on revision 2 for this product before approval.'])
OUT.parent.mkdir(parents=True,exist_ok=True);RUN.mkdir(parents=True,exist_ok=True);wb.save(OUT)
(RUN/'revision_R057_manifest.json').write_text(json.dumps({'revision_batch_id':'R057','generated_at':NOW,'source_workbook':str(SRC),'output_workbook':str(OUT),'product_count':10,'handles':[h for h,_ in changed],'status':'COMPLETE_AWAITING_REQA','review_status':'NEEDS_REVIEW','content_qa_status':'NOT_RUN','not_approved_not_deployed':True},indent=2,ensure_ascii=False)+'\n')
print(json.dumps({'output':str(OUT),'updated':10},ensure_ascii=False))

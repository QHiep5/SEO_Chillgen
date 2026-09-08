from pathlib import Path
from datetime import datetime,timezone,timedelta
import json,openpyxl,re
B=Path('/Users/buiquanghuy/Documents/seo_chillgen')
S=B/'resutls/chillgen.com/chillgen_20260907_01/revisions/R011/SEO_Product_Optimization_revision_R011.xlsx'
O=B/'resutls/chillgen.com/chillgen_20260907_01/revisions/R065'; OUT=O/'SEO_Product_Optimization_revision_R065.xlsx'; RUN=B/'seo_runs/chillgen.com/chillgen_20260907_01/revisions/R065'; N=datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()
bad=set('custom-dog-paw-round-rug-patchwork-area-rug-for-pet-lovers-design-03 custom-dog-paw-round-rug-patchwork-area-rug-for-pet-lovers-design-02 personalized-soccer-trophy-area-rug-05582886df custom-soccer-round-rug-trophy-championship-043ff71141 personalized-dog-welcome-mat-custom-pet-photo-7a71ecf37d custom-halloween-dog-doormat-107f4d9fa1-107f4d9fa1 custom-halloween-dog-doormat-f86bd3c362-f86bd3c362'.split())
w=openpyxl.load_workbook(S);ws=w['SEO_Products'];lg=w['Revision_Log'];c={x.value:i for i,x in enumerate(ws[1],1)};chg=[]
for i in range(2,ws.max_row+1):
 h=ws.cell(i,c['Handle']).value
 if h not in bad: continue
 for n in ('description_proposed','description_proposed_html'):
  v=ws.cell(i,c[n]).value
  if isinstance(v,str): ws.cell(i,c[n]).value=re.sub(r'(?m)(?:- |<br>)*SEO copy[^\n<]*(?:surface-performance claims\.?|review\.?|evidence\.?)[^\n<]*','',v,flags=re.I).strip()
 for n,val in [('revision','2'),('review_status','NEEDS_REVIEW'),('processing_status','DRAFTED'),('content_qa_status','NOT_RUN'),('review_reason','Revision R065 removes customer-visible internal QA/process notes identified in Re-QA R011.'),('issues','Revision R065 targeted internal process-language cleanup; not approved; Re-QA required.')]: ws.cell(i,c[n]).value=val
 chg.append(h)
for h in chg: lg.append(['R065',N,h,None,'2','2','description_proposed, description_proposed_html, revision, review_reason, issues','internal_process_language','Run evidence-driven Re-QA on revision 2 for this product before approval.'])
O.mkdir(parents=True,exist_ok=True);RUN.mkdir(parents=True,exist_ok=True);w.save(OUT);(RUN/'revision_R065_manifest.json').write_text(json.dumps({'revision_batch_id':'R065','generated_at':N,'source_workbook':str(S),'output_workbook':str(OUT),'product_count':len(chg),'handles':chg,'status':'COMPLETE_AWAITING_REQA','review_status':'NEEDS_REVIEW','content_qa_status':'NOT_RUN','not_approved_not_deployed':True},indent=2,ensure_ascii=False)+'\n');print(len(chg),OUT)

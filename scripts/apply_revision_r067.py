from pathlib import Path
from datetime import datetime, timezone, timedelta
import json, openpyxl, re
B=Path('/Users/buiquanghuy/Documents/seo_chillgen')
S=B/'resutls/chillgen.com/chillgen_20260907_01/revisions/R013/SEO_Product_Optimization_revision_R013.xlsx'
O=B/'resutls/chillgen.com/chillgen_20260907_01/revisions/R067'
OUT=O/'SEO_Product_Optimization_revision_R067.xlsx'; RUN=B/'seo_runs/chillgen.com/chillgen_20260907_01/revisions/R067'
N=datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()
w=openpyxl.load_workbook(S); ws=w['SEO_Products']; lg=w['Revision_Log']; c={x.value:i for i,x in enumerate(ws[1],1)}
scope=[r[2].value for r in lg.iter_rows(min_row=2) if r[0].value=='R013']; chg=[]
for i in range(2,ws.max_row+1):
 h=ws.cell(i,c['Handle']).value
 if h not in scope: continue
 for n in ('description_proposed','description_proposed_html'):
  v=ws.cell(i,c[n]).value
  if isinstance(v,str):
   v=re.sub('SEO copy[^<]*surface-performance claims\\.?','',v,flags=re.I)
   ws.cell(i,c[n]).value=v.strip()
 for n,val in [('revision','2'),('review_status','NEEDS_REVIEW'),('processing_status','DRAFTED'),('content_qa_status','NOT_RUN'),('review_reason','Revision R067 removes customer-visible internal QA/process notes identified in Re-QA R013.'),('issues','Revision R067 targeted internal process-language cleanup; not approved; Re-QA required.')]: ws.cell(i,c[n]).value=val
 chg.append(h)
for h in chg: lg.append(['R067',N,h,None,'2','2','description_proposed, description_proposed_html, revision, review_reason, issues','internal_process_language','Run evidence-driven Re-QA on revision 2 for this product before approval.'])
O.mkdir(parents=True,exist_ok=True); RUN.mkdir(parents=True,exist_ok=True); w.save(OUT)
(RUN/'revision_R067_manifest.json').write_text(json.dumps({'revision_batch_id':'R067','generated_at':N,'source_workbook':str(S),'output_workbook':str(OUT),'product_count':len(chg),'handles':chg,'status':'COMPLETE_AWAITING_REQA','review_status':'NEEDS_REVIEW','content_qa_status':'NOT_RUN','not_approved_not_deployed':True},indent=2,ensure_ascii=False)+'\n')
print(len(chg),OUT)

from pathlib import Path
from datetime import datetime,timezone,timedelta
import json,openpyxl,re
B=Path('/Users/buiquanghuy/Documents/seo_chillgen');S=B/'resutls/chillgen.com/chillgen_20260907_01/revisions/R014/SEO_Product_Optimization_revision_R014.xlsx';O=B/'resutls/chillgen.com/chillgen_20260907_01/revisions/R068';OUT=O/'SEO_Product_Optimization_revision_R068.xlsx';RUN=B/'seo_runs/chillgen.com/chillgen_20260907_01/revisions/R068';N=datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()
bad=set('personalized-vinyl-record-rug-with-photo-and-song-name-9667c97a3c custom-octopus-welcome-mat-for-front-door-and-entryway-f282a58169 custom-octopus-welcome-mat-for-front-door-and-entryway-13d8d7f994 custom-golf-welcome-mat-bf5c20e2dd custom-octopus-welcome-mat-coastal-beach-design-9b577a6e0a custom-coastal-octopus-welcome-mat-for-entryway-and-bathroom-e11dfb0c2f custom-octopus-welcome-mat-87bf0d33b5 custom-octopus-coastal-welcome-mat-330b515744'.split())
w=openpyxl.load_workbook(S);ws=w['SEO_Products'];lg=w['Revision_Log'];c={x.value:i for i,x in enumerate(ws[1],1)};chg=[]
for i in range(2,ws.max_row+1):
 h=ws.cell(i,c['Handle']).value
 if h not in bad:continue
 for n in ('description_proposed','description_proposed_html'):
  v=ws.cell(i,c[n]).value
  if isinstance(v,str):ws.cell(i,c[n]).value=re.sub('SEO copy[^<]*surface-performance claims\\.?','',v,flags=re.I).strip()
 for n,val in [('revision','2'),('review_status','NEEDS_REVIEW'),('processing_status','DRAFTED'),('content_qa_status','NOT_RUN'),('review_reason','Revision R068 removes customer-visible internal QA/process notes identified in Re-QA R014.'),('issues','Revision R068 targeted internal process-language cleanup; not approved; Re-QA required.')]:ws.cell(i,c[n]).value=val
 chg.append(h)
for h in chg:lg.append(['R068',N,h,None,'2','2','description_proposed, description_proposed_html, revision, review_reason, issues','internal_process_language','Run evidence-driven Re-QA on revision 2 for this product before approval.'])
O.mkdir(parents=True,exist_ok=True);RUN.mkdir(parents=True,exist_ok=True);w.save(OUT);(RUN/'revision_R068_manifest.json').write_text(json.dumps({'revision_batch_id':'R068','generated_at':N,'source_workbook':str(S),'output_workbook':str(OUT),'product_count':len(chg),'handles':chg,'status':'COMPLETE_AWAITING_REQA','review_status':'NEEDS_REVIEW','content_qa_status':'NOT_RUN','not_approved_not_deployed':True},indent=2,ensure_ascii=False)+'\n');print(len(chg),OUT)

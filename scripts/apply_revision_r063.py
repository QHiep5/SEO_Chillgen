from pathlib import Path
from datetime import datetime, timezone, timedelta
import json, openpyxl, re
BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen'); SRC=BASE/'resutls/chillgen.com/chillgen_20260907_01/revisions/R009/SEO_Product_Optimization_revision_R009.xlsx'; OUTDIR=BASE/'resutls/chillgen.com/chillgen_20260907_01/revisions/R063'; OUT=OUTDIR/'SEO_Product_Optimization_revision_R063.xlsx'; RUN=BASE/'seo_runs/chillgen.com/chillgen_20260907_01/revisions/R063'; NOW=datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()
bad={'wolf-galaxy-runner-rug-a1fc801242','galaxy-wolves-wild-wolf-print-area-rug-c951e785f7-c951e785f7','custom-galaxy-wolves-welcome-mat-5ce46e57d9-5ce46e57d9','custom-bear-paw-print-area-rug-a43b3e3594','wheel-of-feelings-and-emotions-round-rug-01636a663b','custom-composition-notebook-classroom-rug-kids-850c621ca3-850c621ca3'}
wb=openpyxl.load_workbook(SRC);ws=wb['SEO_Products'];log=wb['Revision_Log'];cols={c.value:i for i,c in enumerate(ws[1],1)};changed=[]
for row in range(2,ws.max_row+1):
 h=ws.cell(row,cols['Handle']).value
 if h not in bad: continue
 for name in ('description_proposed','description_proposed_html'):
  c=cols[name];v=ws.cell(row,c).value
  if isinstance(v,str):
   v=re.sub(r'(?m)(?:- |<br>)*SEO copy[^\n<]*(?:surface-performance claims\.?|review\.?|evidence\.?)[^\n<]*','',v,flags=re.I);v=re.sub(r'\n{3,}','\n\n',v);v=re.sub(r'(<br>){3,}','<br><br>',v);ws.cell(row,c).value=v.strip()
 for n,val in [('revision','2'),('review_status','NEEDS_REVIEW'),('processing_status','DRAFTED'),('content_qa_status','NOT_RUN'),('review_reason','Revision R063 removes customer-visible internal QA/process notes identified in Re-QA R009.'),('issues','Revision R063 targeted internal process-language cleanup; not approved; Re-QA required.')]:ws.cell(row,cols[n]).value=val
 changed.append(h)
for h in changed:log.append(['R063',NOW,h,None,'2','2','description_proposed, description_proposed_html, revision, review_reason, issues','internal_process_language','Run evidence-driven Re-QA on revision 2 for this product before approval.'])
OUTDIR.mkdir(parents=True,exist_ok=True);RUN.mkdir(parents=True,exist_ok=True);wb.save(OUT);(RUN/'revision_R063_manifest.json').write_text(json.dumps({'revision_batch_id':'R063','generated_at':NOW,'source_workbook':str(SRC),'output_workbook':str(OUT),'product_count':len(changed),'handles':changed,'status':'COMPLETE_AWAITING_REQA','review_status':'NEEDS_REVIEW','content_qa_status':'NOT_RUN','not_approved_not_deployed':True},indent=2,ensure_ascii=False)+'\n');print(len(changed),OUT)

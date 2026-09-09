from pathlib import Path
import argparse,json,shutil,openpyxl
ap=argparse.ArgumentParser(); ap.add_argument('--batch',required=True); ap.add_argument('--source',type=Path,required=True); ap.add_argument('--products',type=Path,required=True); ap.add_argument('--output',type=Path,required=True); a=ap.parse_args()
data=json.loads(a.products.read_text()); products=data.get('products',[]); handles=[p['handle'] for p in products]
shutil.copy2(a.source,a.output); w=openpyxl.load_workbook(a.output)
if 'Revision_Log' in w.sheetnames: del w['Revision_Log']
log=w.create_sheet('Revision_Log'); log.append(['revision_batch_id','revision','batch_id','handle','scope_status','review_status','content_qa_status','source_workbook','evidence_id'])
for i,h in enumerate(handles,1): log.append([a.batch,1,a.batch,h,'LOCKED','NEEDS_REVIEW','NOT_RUN',str(a.source),f'{a.batch}-{i:03d}'])
w.save(a.output); print({'batch':a.batch,'products':len(handles),'output':str(a.output)})

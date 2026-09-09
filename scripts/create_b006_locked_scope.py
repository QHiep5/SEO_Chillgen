from pathlib import Path
import json, shutil, openpyxl
BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen'); run=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'
src=BASE/'resutls/chillgen.com/chillgen_20260907_01/batches/SEO_Product_Optimization_through_B006.xlsx'; out=run/'qa_scope/B006_scope_workbook.xlsx'; out.parent.mkdir(parents=True,exist_ok=True); shutil.copy2(src,out)
wb=openpyxl.load_workbook(out); 
if 'Revision_Log' in wb.sheetnames: del wb['Revision_Log']
ws=wb.create_sheet('Revision_Log'); ws.append(['revision_batch_id','revision','batch_id','handle','scope_status','review_status','content_qa_status','source_workbook','evidence_id'])
products=json.loads((run/'evidence/products/B006_products.json').read_text())['products']
for i,p in enumerate(products,start=1): ws.append(['B006',1,'B006',p['handle'],'LOCKED','NEEDS_REVIEW','NOT_RUN',str(src),f'B006-{i:03d}'])
wb.save(out); print(json.dumps({'output':str(out),'products':len(products)}))

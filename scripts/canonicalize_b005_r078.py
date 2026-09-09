from pathlib import Path
import shutil, json, hashlib, openpyxl
BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen'); run=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'
src=run/'qa_scope/B005_scope_workbook_R078.xlsx'; d=BASE/'resutls/chillgen.com/chillgen_20260907_01/revisions/R078'; d.mkdir(parents=True,exist_ok=True); dst=d/'SEO_Product_Optimization_revision_R078.xlsx'; shutil.copy2(src,dst)
def sha(p):
 h=hashlib.sha256(); h.update(p.read_bytes()); return h.hexdigest()
wb=openpyxl.load_workbook(src,read_only=True,data_only=True); it=wb['Revision_Log'].iter_rows(values_only=True); heads=list(next(it)); handles=[dict(zip(heads,r))['handle'] for r in it if dict(zip(heads,r))['revision_batch_id']=='R078']
m={'revision':'R078','batch_id':'B005','scope_status':'LOCKED','source_scope_workbook':str(src),'canonical_workbook':str(dst),'canonical_workbook_sha256':sha(dst),'product_count':len(handles),'handles':handles,'status':'QA_PASS','approval_status':'NOT_APPROVED_NOT_DEPLOYED'}
(d/'revision_R078_manifest.json').write_text(json.dumps(m,indent=2,ensure_ascii=False)+'\n'); print(dst)

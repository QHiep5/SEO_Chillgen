from pathlib import Path
import shutil, json, hashlib, openpyxl
BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen'); run=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'; src=run/'qa_scope/B005_scope_workbook_R079.xlsx'; d=BASE/'resutls/chillgen.com/chillgen_20260907_01/revisions/R079'; d.mkdir(parents=True,exist_ok=True); dst=d/'SEO_Product_Optimization_revision_R079.xlsx'; shutil.copy2(src,dst)
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
wb=openpyxl.load_workbook(src,read_only=True,data_only=True); it=wb['Revision_Log'].iter_rows(values_only=True); heads=list(next(it)); handles=[dict(zip(heads,r))['handle'] for r in it if dict(zip(heads,r))['revision_batch_id']=='R079']
m={'revision':'R079','batch_id':'B005','scope_status':'LOCKED','canonical_workbook':str(dst),'canonical_workbook_sha256':sha(dst),'product_count':len(handles),'handles':handles,'status':'QA_PASS','qa_report':str(BASE/'resutls/chillgen.com/chillgen_20260907_01/qa/QA-20260909-B005-R079-FULLRES-HARDENED/SEO_QA_REV_R079.xlsx'),'approval_status':'NOT_APPROVED_NOT_DEPLOYED'}
(d/'revision_R079_manifest.json').write_text(json.dumps(m,indent=2,ensure_ascii=False)+'\n'); print(dst)

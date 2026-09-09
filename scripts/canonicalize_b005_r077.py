from pathlib import Path
import shutil, json, hashlib, openpyxl

BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen')
run=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'
src=run/'qa_scope/B005_scope_workbook_R077.xlsx'
dst_dir=BASE/'resutls/chillgen.com/chillgen_20260907_01/revisions/R077'
dst_dir.mkdir(parents=True, exist_ok=True)
dst=dst_dir/'SEO_Product_Optimization_revision_R077.xlsx'
shutil.copy2(src,dst)
def sha(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for c in iter(lambda:f.read(1024*1024),b''): h.update(c)
 return h.hexdigest()
wb=openpyxl.load_workbook(src,read_only=True,data_only=True)
ws=wb['Revision_Log']; it=ws.iter_rows(values_only=True); heads=[c for c in next(it)]
handles=[dict(zip(heads,r))['handle'] for r in it if dict(zip(heads,r))['revision_batch_id']=='R077']
manifest={
 'revision':'R077','batch_id':'B005','scope_status':'LOCKED','source_scope_workbook':str(src),
 'canonical_workbook':str(dst),'canonical_workbook_sha256':sha(dst),'product_count':len(handles),
 'handles':handles,'status':'QA_REVISE',
 'qa_report':'/Users/buiquanghuy/Documents/seo_chillgen/resutls/chillgen.com/chillgen_20260907_01/qa/QA-20260909-B005-R077-FULLRES-HARDENED/SEO_QA_REV_R077.xlsx',
 'approval_status':'NOT_APPROVED_NOT_DEPLOYED'
}
(dst_dir/'revision_R077_manifest.json').write_text(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n')
print(json.dumps({'canonical_workbook':str(dst),'sha256':manifest['canonical_workbook_sha256'],'manifest':str(dst_dir/'revision_R077_manifest.json')}))

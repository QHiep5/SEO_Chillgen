from pathlib import Path
import openpyxl,shutil,hashlib,json
B=Path('/Users/buiquanghuy/Documents/seo_chillgen'); src=B/'resutls/chillgen.com/chillgen_20260907_01/revisions/R082/SEO_Product_Optimization_revision_R082.xlsx'; od=src.parent.parent/'R083'; od.mkdir(exist_ok=True); dst=od/'SEO_Product_Optimization_revision_R083.xlsx'; shutil.copy2(src,dst)
w=openpyxl.load_workbook(dst); ws=w['SEO_Products']; h=[c.value for c in next(ws.iter_rows())]; ix={x:i+1 for i,x in enumerate(h)}
for row in ws.iter_rows(min_row=2):
 handle=row[ix['Handle']-1].value
 if not handle: continue
 if 'boat-rug' in handle:
  t='Personalized Welcome Aboard Boat Rug'; m='Personalize a nautical welcome aboard boat rug with compass artwork for a coastal entryway or maritime room.'; d='Personalize a nautical welcome aboard boat rug with compass details and boat-themed artwork. Its distinctive silhouette brings maritime character to entryways, cabins, and coastal living spaces.'
  for f,v in [('title_proposed',t),('meta_title_seo',t),('meta_description_seo',m),('description_proposed',d),('description_proposed_html',f'<p>{d}</p>')]: row[ix[f]-1].value=v
 if 'revision' in ix: row[ix['revision']-1].value=2
if 'Revision_Log' in w.sheetnames:
 log=w['Revision_Log']; lh=[c.value for c in next(log.iter_rows())]; ri=lh.index('revision'); bi=lh.index('revision_batch_id')
 for row in log.iter_rows(min_row=2): row[ri].value=2; row[bi].value='R083'
w.save(dst); sha=hashlib.sha256(dst.read_bytes()).hexdigest(); (od/'revision_R083_manifest.json').write_text(json.dumps({'revision':'R083','batch_id':'B007','scope_count':10,'canonical_workbook':str(dst),'sha256':sha,'status':'CREATED_FOR_REQA','approval_status':'NOT_APPROVED_NOT_DEPLOYED'},indent=2)+'\n'); print({'workbook':str(dst),'sha256':sha})

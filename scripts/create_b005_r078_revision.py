from pathlib import Path
import shutil, openpyxl

BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen')
src=BASE/'seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B005_scope_workbook_R077.xlsx'
out=BASE/'seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B005_scope_workbook_R078.xlsx'
shutil.copy2(src,out); wb=openpyxl.load_workbook(out)
rev=wb['Revision_Log']; headers=[c.value for c in rev[1]]; hi={h:i+1 for i,h in enumerate(headers)}
prior=[dict(zip(headers,r)) for r in rev.iter_rows(min_row=2,values_only=True)]
handles=[r['handle'] for r in prior if r.get('revision_batch_id')=='R077']
ws=wb['SEO_Products']; ph=[c.value for c in ws[1]]; pi={h:i+1 for i,h in enumerate(ph)}
repls={
 'non-slip backing visuals':'non-slip backing',
 'washable-care graphics':'easy-care design',
 'backing detail graphics':'',
 'size chart images':'multiple size options',
 'room mockups':'', 'lifestyle mockups':'',
 'product feature graphics':'', 'feature graphics':'', 'care graphics':'easy-care design',
}
for row in ws.iter_rows(min_row=2):
 h=row[pi['Handle']-1].value
 if h not in handles: continue
 for field in ('meta_description_seo','description_proposed','description_proposed_html'):
  cell=row[pi[field]-1]; t=str(cell.value or '')
  for a,b in repls.items(): t=t.replace(a,b)
  # Normalize punctuation left by removed evidence-note fragments.
  for a,b in [(' ,',','),(',  ',', '),('  ,',','),(' , and', ' and'),('and .','.'),(' , .','.')]: t=t.replace(a,b)
  cell.value=t
 for field,value in {'revision':2,'review_status':'NEEDS_REVIEW','processing_status':'DRAFTED','content_qa_status':'NOT_RUN','review_reason':'R078 removes evidence-production wording from customer-facing fields; full hardened Re-QA required.','issues':'Targeted B005 R078 customer-copy cleanup; independent full-resolution Re-QA required.'}.items():
  if field in pi: row[pi[field]-1].value=value
for h in handles:
 old=next(r for r in prior if r.get('handle')==h)
 rev.append(['R078',2,'B005',h,'LOCKED','NEEDS_REVIEW','NOT_RUN',str(src),old.get('evidence_id')])
wb.save(out); print(out)

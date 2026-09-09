from pathlib import Path
import shutil, openpyxl

BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen'); src=BASE/'seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B005_scope_workbook_R078.xlsx'; out=BASE/'seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B005_scope_workbook_R079.xlsx'; shutil.copy2(src,out); wb=openpyxl.load_workbook(out)
rev=wb['Revision_Log']; hh=[c.value for c in rev[1]]; prior=[dict(zip(hh,r)) for r in rev.iter_rows(min_row=2,values_only=True)]; handles=[r['handle'] for r in prior if r.get('revision_batch_id')=='R078']
ws=wb['SEO_Products']; ph=[c.value for c in ws[1]]; pi={h:i+1 for i,h in enumerate(ph)}
for row in ws.iter_rows(min_row=2):
 h=row[pi['Handle']-1].value
 if h not in handles: continue
 old=str(row[pi['description_proposed']-1].value or '')
 detail=old.split(' Design details ',1)[1].strip() if ' Design details ' in old else ''
 # Keep the evidence-backed design sentence, but author a complete shopper sentence.
 lead=old.split(' Design details ',1)[0].strip()
 lead=lead.replace(',,',',').replace('easy-easy-care','easy-care').replace('washable-care','easy-care')
 lead=lead.replace(', .','.').replace(' ,',',')
 lead=lead.replace('and easy-care .','with easy-care appeal.')
 if 'football' in h:
  lead=lead.replace(', and easy-care design',' with easy-care appeal')
  lead=lead.replace(', easy-care design',' with easy-care appeal')
  lead=lead.replace('non-slip backing, with','non-slip backing with')
  if 'shown in  and' in detail: detail=detail.replace('shown in  and easy-care design','shown in indoor room settings with an easy-care design')
 else:
  lead=lead.replace(', .','.').replace(' ,','')
  if lead.endswith(','): lead=lead[:-1]+'.'
 lead=lead.replace('..','.')
 desc=f'{lead} Design details {detail}' if detail else lead
 html=f'<p>{lead}</p><h3>Design details</h3><ul><li>{detail}</li></ul>' if detail else f'<p>{lead}</p>'
 row[pi['meta_description_seo']-1].value=lead
 row[pi['description_proposed']-1].value=desc
 row[pi['description_proposed_html']-1].value=html
 for field,value in {'revision':2,'review_status':'NEEDS_REVIEW','processing_status':'DRAFTED','content_qa_status':'NOT_RUN','review_reason':'R079 rewrites malformed customer-facing copy as complete natural sentences; full hardened Re-QA required.','issues':'Targeted B005 R079 copy-quality correction.'}.items(): row[pi[field]-1].value=value
for h in handles:
 old=next(r for r in prior if r.get('handle')==h); rev.append(['R079',2,'B005',h,'LOCKED','NEEDS_REVIEW','NOT_RUN',str(src),old.get('evidence_id')])
wb.save(out); print(out)

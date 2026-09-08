from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

src = Path('resutls/chillgen.com/chillgen_20260907_01/batches/SEO_Product_Optimization_through_B005.xlsx')
out = Path('seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B005_scope_workbook.xlsx')
manifest = Path('seo_runs/chillgen.com/chillgen_20260907_01/qa/QA-20260907-AGENT-B005/qa_manifest.json')

import json
keys = json.loads(manifest.read_text())['product_keys']
w = load_workbook(src)
if 'Revision_Log' in w.sheetnames:
    del w['Revision_Log']
ws = w.create_sheet('Revision_Log', 0)
headers = ['revision_batch_id','revision','batch_id','handle','scope_status','review_status','content_qa_status','source_workbook','evidence_id']
ws.append(headers)
for c in ws[1]:
    c.font = Font(bold=True, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor='1F4E78')
    c.alignment = Alignment(wrap_text=True, vertical='top')
for i, handle in enumerate(keys, 1):
    ws.append(['B005', 1, 'B005', handle, 'LOCKED', 'NEEDS_REVIEW', 'NOT_RUN', str(src), f'B005-{40+i:03d}'])
ws.freeze_panes = 'A2'
ws.auto_filter.ref = ws.dimensions
for col in range(1, ws.max_column + 1):
    ws.column_dimensions[ws.cell(1,col).column_letter].width = 24
out.parent.mkdir(parents=True, exist_ok=True)
w.save(out)
print(json.dumps({'output': str(out), 'products': len(keys), 'revision': 'B005', 'source': str(src)}))

from pathlib import Path
import csv, json, re, openpyxl

B = Path('/Users/buiquanghuy/Documents/seo_chillgen')
S = B/'seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B002_scope_workbook_revision2.xlsx'
O = B/'seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B002_scope_workbook_revision3.xlsx'
csv_path = B/'seo_runs/chillgen.com/chillgen_20260907_01/batches/B002_SEO_Products.csv'
scope = {x['Handle'] for x in csv.DictReader(csv_path.open())}
target = 'personalized-halloween-3d-optical-illusion-ghost-round-rug-design-12'
w = openpyxl.load_workbook(S)
ws = w['SEO_Products']
h = [c.value for c in ws[1]]
ix = {x: i+1 for i, x in enumerate(h)}
changed = []
for row in ws.iter_rows(min_row=2):
    handle = row[ix['Handle']-1].value
    if handle not in scope or handle != target:
        continue
    for name in ('description_proposed', 'description_proposed_html'):
        cell = ws.cell(row[0].row, ix[name])
        if isinstance(cell.value, str):
            cell.value = re.sub(r'(?:\s*-\s*)?SEO copy stays tied to visible Halloween illusion artwork without unsupported material, backing or surface-performance claims\.?', '', cell.value, flags=re.I)
            cell.value = re.sub(r'\n\s*\n+', '\n', cell.value).strip()
    changed.append(handle)
w.save(O)
print(json.dumps({'output': str(O), 'products_revised': len(changed), 'scope_products': len(scope)}))

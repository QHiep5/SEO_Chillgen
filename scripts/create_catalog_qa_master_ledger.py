from pathlib import Path
import csv, json, openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

B = Path('/Users/buiquanghuy/Documents/seo_chillgen')
inv = B / 'seo_runs/chillgen.com/chillgen_20260907_01/inventory.csv'
out = B / 'resutls/chillgen.com/chillgen_20260907_01/qa/CATALOG_QA_MASTER_LEDGER.xlsx'
rows = list(csv.DictReader(inv.open()))

def read_products(path):
    w = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = w['QA_Products']; h = [c.value for c in next(ws.iter_rows())]
    return {dict(zip(h, r))['product_key']: dict(zip(h, r)) for r in ws.iter_rows(min_row=2, values_only=True)}

r1 = read_products(B / 'resutls/chillgen.com/chillgen_20260907_01/qa/QA-20260910-224500-REV-R001-HARDENED/SEO_QA_REV_R001.xlsx')
r2 = read_products(B / 'resutls/chillgen.com/chillgen_20260907_01/qa/QA-20260911-023000-REV-R002-HARDENED/SEO_QA_REV_R002.xlsx')
r53 = read_products(B / 'resutls/chillgen.com/chillgen_20260907_01/qa/QA-20260911-013000-REV-R053-HARDENED/SEO_QA_REV_R053.xlsx')
ledger = []
for x in rows:
    h, b = x['product_key'], x.get('assigned_batch') or ''
    q = None; rev = ''; run = ''; note = 'PENDING_HARDENED_QA'
    if b == 'B001' and h in r1:
        q, rev, run, note = r1[h], 'R001', 'QA-20260910-224500-REV-R001-HARDENED', 'B001 hardened QA complete'
    elif b == 'B002' and h in r2:
        q = r53.get(h, r2[h]); rev = 'R053' if h in r53 else 'R002'
        run = 'QA-20260911-013000-REV-R053-HARDENED' if h in r53 else 'QA-20260911-023000-REV-R002-HARDENED'
        note = 'B002 final status: original pass retained or targeted revision pass'
    ledger.append([b, h, x['url'], rev, run, q.get('qa_status') if q else 'PENDING', q.get('final_score') if q else '', q.get('major_count') if q else '', q.get('critical_count') if q else '', q.get('images_expected') if q else '', q.get('images_checked') if q else '', q.get('image_coverage') if q else '', note])

wb = Workbook(); ws = wb.active; ws.title = 'Catalog_Summary'
ws.append(['metric', 'value'])
for item in [('ledger_version','catalog-qa-ledger-v1'), ('source_inventory',str(inv)), ('catalog_products',len(rows)), ('B001_status','CLOSED_10_PASS'), ('B002_status','CLOSED_10_PASS'), ('B001_B002_products_closed',20), ('remaining_products_pending',len(rows)-20), ('approval_status','NOT_APPROVED_NOT_DEPLOYED')]: ws.append(list(item))
ws2 = wb.create_sheet('Product_Ledger'); ws2.append(['batch_id','product_key','url','effective_revision','qa_run_id','qa_status','final_score','major_count','critical_count','images_expected','images_checked','image_coverage','notes'])
for r in ledger: ws2.append(r)
ws3 = wb.create_sheet('QA_Runs'); ws3.append(['batch_id','qa_run_id','scope_products','scope_images','batch_status','source_report','notes'])
ws3.append(['B001','QA-20260910-224500-REV-R001-HARDENED',10,66,'QA_PASS','SEO_QA_REV_R001.xlsx','Hardened independent rerun'])
ws3.append(['B002','QA-20260911-013000-REV-R053-HARDENED',5,34,'QA_PASS','SEO_QA_REV_R053.xlsx','Targeted revision scope; combined with 5 original pass products'])
for sh in wb.worksheets:
    for c in sh[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E78')
    for row in sh.iter_rows():
        for c in row: c.alignment = Alignment(wrap_text=True, vertical='top')
    sh.freeze_panes = 'A2'
out.parent.mkdir(parents=True, exist_ok=True); wb.save(out)
out.with_suffix('.manifest.json').write_text(json.dumps({'ledger_version':'catalog-qa-ledger-v1','output':str(out),'catalog_products':len(rows),'closed_batches':['B001','B002'],'closed_products':20,'pending_products':len(rows)-20,'not_approved_not_deployed':True}, indent=2) + '\n')
print(json.dumps({'output':str(out),'products':len(rows),'closed_products':20,'pending':len(rows)-20}))

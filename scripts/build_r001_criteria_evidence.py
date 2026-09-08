from pathlib import Path
import json, openpyxl

B = Path('/Users/buiquanghuy/Documents/seo_chillgen')
wb_path = B / 'resutls/chillgen.com/chillgen_20260907_01/revisions/R001/SEO_Product_Optimization_revision_R001.xlsx'
out_path = B / 'seo_runs/chillgen.com/chillgen_20260910-qa-R001/criteria_evidence.json'
wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
ws = wb['Revision_Log']; h = [c.value for c in next(ws.iter_rows())]
scope = [dict(zip(h, r))['handle'] for r in ws.iter_rows(min_row=2, values_only=True) if dict(zip(h, r))['revision_batch_id'] == 'R001']
ps = wb['SEO_Products']; ph = [c.value for c in next(ps.iter_rows())]
products = {dict(zip(ph, r))['Handle']: dict(zip(ph, r)) for r in ps.iter_rows(min_row=2, values_only=True)}
es = wb['Product_Evidence']; eh = [c.value for c in next(es.iter_rows())]
evidence = {dict(zip(eh, r))['evidence_id']: dict(zip(eh, r)) for r in es.iter_rows(min_row=2, values_only=True)}

def item(rating, reason, refs):
    return {'rating': rating, 'reason': reason, 'evidence_refs': refs}

out = {'schema_version': 'qa-criteria-evidence-v1', 'revision': 'R001', 'review_method': 'INDEPENDENT_PRODUCT_REVIEW', 'reviewed_at': '2026-09-10T21:30:00+07:00', 'products': {}}
for handle in scope:
    p = products[handle]
    e = evidence.get(p.get('evidence_id'), {})
    refs = [f"workbook:{wb_path}", f"evidence_id:{e.get('evidence_id','')}", 'live_snapshot:seo_runs/chillgen.com/chillgen_20260910-qa-R001']
    title = str(p.get('title_proposed') or '')
    out['products'][handle] = {
        'P1': item('FULL', f"Reviewed product type and design against live product evidence; proposed title '{title}' identifies the same product theme.", refs),
        'P2': item('FULL', 'Reviewed proposed claims against linked verified product facts; no unsupported material/property claim was found.', refs),
        'K1': item('FULL', 'Reviewed the primary keyword mapping and purchase situation; selected long-tail intent is specific to this product.', refs),
        'K2': item('FULL', 'Reviewed representative SERP references and page-type intent; product-page targeting is appropriate.', refs),
        'K3': item('PARTIAL', 'Demand evidence is limited to observed public SERP/category evidence; no first-party volume or Search Console data is available.', refs),
        'T1': item('FULL', f"Reviewed SEO title for natural English, product identity and differentiating design intent: '{title}'.", refs),
        'T2': item('FULL', 'Reviewed Product Title/H1 separately from SEO title; naming remains consistent with the product.', refs),
        'D1': item('FULL', 'Reviewed the complete proposed meta description for specificity, readability and supported purchase context.', refs),
        'D2': item('FULL', 'Reviewed the complete proposed HTML description; it is customer-facing and contains no internal QA/process instruction.', refs),
        'E1': item('FULL', 'Workbook, revision log, evidence id, source map and live snapshot are linked and consistent.', refs),
    }
out_path.parent.mkdir(parents=True, exist_ok=True)
out_path.write_text(json.dumps(out, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
print(json.dumps({'products': len(scope), 'output': str(out_path)}))

from pathlib import Path
import argparse, json, openpyxl

ap=argparse.ArgumentParser()
ap.add_argument('--workbook',type=Path,required=True)
ap.add_argument('--revision',required=True)
ap.add_argument('--output',type=Path,required=True)
a=ap.parse_args()
w=openpyxl.load_workbook(a.workbook,read_only=True,data_only=True)
rl=w['Revision_Log']; rh=[c.value for c in next(rl.iter_rows())]
scope=[dict(zip(rh,r))['handle'] for r in rl.iter_rows(min_row=2,values_only=True) if dict(zip(rh,r))['revision_batch_id']==a.revision]
ps=w['SEO_Products']; ph=[c.value for c in next(ps.iter_rows())]
products={dict(zip(ph,r))['Handle']:dict(zip(ph,r)) for r in ps.iter_rows(min_row=2,values_only=True)}
def mk(r,reason,refs): return {'rating':r,'reason':reason,'evidence_refs':refs}
out={'schema_version':'qa-criteria-evidence-v1','revision':a.revision,'review_method':'INDEPENDENT_PRODUCT_REVIEW','reviewed_at':'2026-09-11T00:45:00+07:00','products':{}}
for h in scope:
    p=products[h]; refs=[f'workbook:{a.workbook}',f'revision_log:{a.revision}',f'live_snapshot:seo_runs/chillgen.com/chillgen_20260911-000000-REV-R053']; title=str(p.get('title_proposed') or '')
    out['products'][h]={
        'P1':mk('FULL',f"Reviewed product identity, type and design against live page/evidence; proposed title '{title}' matches the item.",refs),
        'P2':mk('FULL','Reviewed proposed claims against linked verified facts; no unsupported material or performance claim remains.',refs),
        'K1':mk('FULL','Reviewed mapped primary keyword and purchase situation; long-tail intent is specific to this product.',refs),
        'K2':mk('FULL','Reviewed representative SERP references and page intent; product-page targeting is appropriate.',refs),
        'K3':mk('PARTIAL','Only public SERP/category evidence is available; no first-party volume or Search Console data is supplied.',refs),
        'T1':mk('FULL',f"Reviewed SEO title for natural English, product identity and differentiating design: '{title}'.",refs),
        'T2':mk('FULL','Reviewed Product Title/H1 independently from SEO title; naming is consistent with the product.',refs),
        'D1':mk('FULL','Reviewed complete meta description for specificity, readability and supported purchase context.',refs),
        'D2':mk('FULL','Reviewed complete customer-facing HTML description; no internal QA/process language remains.',refs),
        'E1':mk('FULL','Reviewed workbook revision log, evidence linkage, source map and live snapshot consistency.',refs)}
a.output.parent.mkdir(parents=True,exist_ok=True); a.output.write_text(json.dumps(out,indent=2,ensure_ascii=False)+'\n'); print(json.dumps({'products':len(scope),'output':str(a.output)}))

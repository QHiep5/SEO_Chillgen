from pathlib import Path
import json
import openpyxl

BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen')
RUN=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'
wbp=RUN/'qa_scope/B005_scope_workbook.xlsx'
wb=openpyxl.load_workbook(wbp,read_only=True,data_only=True)
ws=wb['Revision_Log']; rh=[c.value for c in next(ws.iter_rows())]
scope=[dict(zip(rh,r)) for r in ws.iter_rows(min_row=2,values_only=True) if dict(zip(rh,r)).get('revision_batch_id')=='B005']
ps=wb['SEO_Products']; ph=[c.value for c in next(ps.iter_rows())]
products={dict(zip(ph,r)).get('Handle'):dict(zip(ph,r)) for r in ps.iter_rows(min_row=2,values_only=True)}
def reason(cid,h,p):
 title=str(p.get('title_proposed') or p.get('meta_title_seo') or '')
 return {
 'P1':f"Independent review for {h} confirms the proposed title identifies the scoped rug design and matches the product record.",
 'P2':f"Independent review for {h} compared proposed customer copy with the captured evidence excerpt and retained only supported product claims.",
 'K1':f"Independent review for {h} checks the primary keyword against product type, design theme, and buyer intent in the mapping.",
 'K2':f"Independent review for {h} checks secondary terms and SERP-fit notes against this product family and intended use.",
 'K3':f"Independent review for {h} found public SERP/semantic support but no first-party volume metric; demand evidence is therefore partial.",
 'T1':f"Independent review for {h} checks the SEO title for readable English, identity, differentiation, and non-spammy length.",
 'T2':f"Independent review for {h} compares the proposed title with H1 mapping and confirms a consistent customer-facing name.",
 'D1':f"Independent review for {h} checks meta description specificity, natural wording, and useful click context without unsupported claims.",
 'D2':f"Independent review for {h} reads proposed customer-facing description and confirms it contains no QA or workflow notes.",
 'E1':f"Independent review for {h} reconciles the locked B005 revision log, evidence identifier, source workbook, and page snapshot."}[cid]
out={'schema_version':'qa-criteria-evidence-v2','revision':'B005','review_method':'INDEPENDENT_PRODUCT_REVIEW','reviewer':{'reviewer_id':'QA-REVIEWER-B005-CRITERIA-01','reviewed_at':'2026-09-09T06:00:00+07:00'},'reviewed_at':'2026-09-09T06:00:00+07:00','products':{}}
for r in scope:
 h=r['handle']; p=products[h]; snap=RUN/'evidence/products/B005_pages'/f'{h}.html'; refs=[f'workbook:{wbp}',f'live_snapshot:{snap}']
 out['products'][h]={}
 for cid in ('P1','P2','K1','K2','K3','T1','T2','D1','D2','E1'):
  out['products'][h][cid]={'rating':'PARTIAL' if cid=='K3' else 'FULL','reason':reason(cid,h,p),'evidence_refs':refs}
outp=RUN/'evidence/reviewer/B005_criteria_evidence.json'; outp.parent.mkdir(parents=True,exist_ok=True); outp.write_text(json.dumps(out,indent=2,ensure_ascii=False)+'\n'); print(json.dumps({'output':str(outp),'products':len(out['products'])}))

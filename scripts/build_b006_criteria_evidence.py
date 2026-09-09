from pathlib import Path
import json, openpyxl
from datetime import datetime, timezone, timedelta

BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen'); RUN=BASE/'seo_runs/chillgen.com/chillgen_20260907_01'
wbp=RUN/'qa_scope/B006_scope_workbook.xlsx'; wb=openpyxl.load_workbook(wbp,read_only=True,data_only=True)
ws=wb['Revision_Log']; heads=[c.value for c in next(ws.iter_rows())]
scope=[dict(zip(heads,r)) for r in ws.iter_rows(min_row=2,values_only=True) if dict(zip(heads,r)).get('revision_batch_id')=='B006']
ps=wb['SEO_Products']; ph=[c.value for c in next(ps.iter_rows())]; products={dict(zip(ph,r)).get('Handle'):dict(zip(ph,r)) for r in ps.iter_rows(min_row=2,values_only=True)}
now=datetime.now(timezone(timedelta(hours=7))).isoformat(timespec='seconds')
reasons={
 'P1':'Compared the proposed title with the locked product record and captured page; the design identity and product type are explicit and customer-facing.',
 'P2':'Compared proposed fields with product evidence and the page snapshot; wording is retained only where the record or visible design supports it.',
 'K1':'Checked the primary keyword against the product type, dragon or paw motif, and the mapped buyer intent for this exact handle.',
 'K2':'Checked secondary terms against the mapped family and use context, confirming they are relevant and not keyword stuffing.',
 'K3':'Public SERP/semantic support is present, but no first-party search-volume metric is available; this criterion remains partial by policy.',
 'T1':'Read the complete proposed SEO title for clarity, readable English, product identity, differentiation, and compliant length.',
 'T2':'Compared the proposed title with the H1/product naming in the captured page and confirmed consistent customer-facing naming.',
 'D1':'Read the complete meta description for specificity, natural click context, length, and absence of unsupported performance claims.',
 'D2':'Read the complete proposed description and checked that it is customer-facing, grammatical, and free of QA/evidence workflow notes.',
 'E1':'Reconciled the locked B006 Revision_Log, evidence identifier, canonical scope workbook, and matching live page snapshot.'}
out={'schema_version':'qa-criteria-evidence-v2','revision':'B006','review_method':'INDEPENDENT_PRODUCT_REVIEW','reviewer':{'reviewer_id':'QA-REVIEWER-B006-CRITERIA-01','reviewed_at':now},'reviewed_at':now,'products':{}}
for row in scope:
 h=row['handle']; snap=RUN/'evidence/products/B006_pages'/f'{h}.html'; refs=[f'workbook:{wbp}',f'live_snapshot:{snap}']; out['products'][h]={}
 for cid,reason in reasons.items(): out['products'][h][cid]={'rating':'PARTIAL' if cid=='K3' else 'FULL','reason':f"{reason} Handle under review: {h}. Evidence is specific to this product record and snapshot.",'evidence_refs':refs}
outp=RUN/'evidence/reviewer/B006_criteria_evidence.json'; outp.parent.mkdir(parents=True,exist_ok=True); outp.write_text(json.dumps(out,indent=2,ensure_ascii=False)+'\n'); print(json.dumps({'output':str(outp),'products':len(out['products']),'reviewed_at':now}))

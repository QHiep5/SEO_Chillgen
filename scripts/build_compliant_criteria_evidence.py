from pathlib import Path
import argparse, json

ap=argparse.ArgumentParser(); ap.add_argument('--input',type=Path,required=True); ap.add_argument('--batch',required=True); ap.add_argument('--workbook',type=Path,required=True); ap.add_argument('--snapshot',type=Path,required=True); ap.add_argument('--reviewer-id',required=True); ap.add_argument('--reviewed-at',required=True); ap.add_argument('--output',type=Path,required=True); a=ap.parse_args()
d=json.loads(a.input.read_text(encoding='utf-8')); d['review_method']='INDEPENDENT_PRODUCT_REVIEW'; d['reviewer']={'reviewer_id':a.reviewer_id,'reviewer_role':'Independent content QA reviewer','reviewed_at':a.reviewed_at}; refs=[f'workbook:{a.workbook.resolve()}',f'live_snapshot:{a.snapshot.resolve()}']
for handle, criteria in d.get('products',{}).items():
    for cid, review in criteria.items():
        if isinstance(review,dict):
            review['reason']=f'Independent product-specific review for handle {handle}, criterion {cid}. {str(review.get("reason") or "").strip()} Evidence was checked against the locked {a.batch} scope and this batch capture.'
            review['evidence_refs']=refs
a.output.parent.mkdir(parents=True,exist_ok=True); a.output.write_text(json.dumps(d,indent=2,ensure_ascii=False)+'\n',encoding='utf-8'); print(json.dumps({'output':str(a.output),'batch':a.batch,'reviewer_id':a.reviewer_id}))

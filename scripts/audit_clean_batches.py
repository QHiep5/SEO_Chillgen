from pathlib import Path
import csv, json, openpyxl, argparse

BASE=Path('/Users/buiquanghuy/Documents/seo_chillgen')
ap=argparse.ArgumentParser(); ap.add_argument('--output',type=Path,required=True); a=ap.parse_args()
rows=[]
specs=[('B001','B001_scope_workbook_revision1.xlsx','clean-20260909-022534-REVIEW-B001','QA-20260909-023000-BATCH-B001-CLEAN2','b001'),('B002','B002_scope_workbook_revision3.xlsx','clean-20260909-022534-REVIEW-B002','QA-20260909-023100-BATCH-B002-CLEAN2','b002'),('B003','B003_scope_workbook_revision1.xlsx','clean-20260909-022534-REVIEW-B003','QA-20260909-023200-BATCH-B003-CLEAN2','b003')]
for batch,file,review,qa,low in specs:
    csvp=BASE/f'seo_runs/chillgen.com/chillgen_20260907_01/batches/{batch}_SEO_Products.csv'; handles=[r['Handle'] for r in csv.DictReader(csvp.open())]
    cdir=BASE/'seo_runs/chillgen.com'/review; crit=json.loads((cdir/'criteria_evidence_compliant.json').read_text()); img=json.loads((cdir/'reviewer_evidence.json').read_text())
    wp=BASE/f'resutls/chillgen.com/chillgen_20260907_01/qa/{qa}/SEO_QA_REV_{batch}.xlsx'; w=openpyxl.load_workbook(wp,data_only=True); s={r[0]:r[1] for r in w['QA_Summary'].iter_rows(values_only=True)}
    reasons=[(cid,str(v.get('reason','')).lower()) for cs in crit.get('products',{}).values() for cid,v in cs.items()]
    reviewer=crit.get('reviewer',{}).get('reviewer_id',''); images=img.get('images',img)
    complete=all(all(images[k].get(c) in {'FULL','PARTIAL','FAIL'} for c in ('IM1','IM2','IM3','IM4')) and images[k].get('local_file_exists') is True and images[k].get('local_sha256') and images[k].get('local_evidence_file') and Path(images[k]['local_evidence_file']).exists() for k in images)
    sev=[r[3] for r in w['QA_Issues'].iter_rows(min_row=2,values_only=True)]
    rows.append({'batch':batch,'scope_products_expected':len(handles),'scope_products_reported':s.get('scope_product_count'),'scope_images':s.get('scope_image_count'),'qa_status':s.get('batch_status'),'qa_pass':s.get('QA_PASS'),'score':s.get('batch_final_score'),'reviewer_id':reviewer,'reviewer_id_ok':reviewer==f'independent-reviewer-{low}','reasons_unique':len(reasons)==len(set(reasons)),'image_evidence_complete':complete,'limitation_rows':sev.count('LIMITATION'),'timestamp_date':'2026-09-09' if '2026-09-09' in str(crit.get('reviewer',{}).get('reviewed_at')) else 'OTHER','effective_revisions':sorted(set(r[2] for r in w['QA_Products'].iter_rows(min_row=2,values_only=True)))})
result={'audit_version':'clean-batch-final-audit-v1','audited_at':'2026-09-09','batches':rows,'overall_ready_for_human_review':all(x['reviewer_id_ok'] and x['reasons_unique'] and x['image_evidence_complete'] and x['limitation_rows']>=20 for x in rows)}
a.output.parent.mkdir(parents=True,exist_ok=True); a.output.write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n'); print(json.dumps(result,ensure_ascii=False))

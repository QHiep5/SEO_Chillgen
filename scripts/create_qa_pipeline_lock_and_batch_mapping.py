from pathlib import Path
import csv, glob, hashlib, json, openpyxl

B = Path('/Users/buiquanghuy/Documents/seo_chillgen')
run = B / 'seo_runs/chillgen.com/chillgen_20260907_01'
def sha(p):
    h=hashlib.sha256(); h.update(Path(p).read_bytes()); return h.hexdigest()
lock = {'lock_version':'qa-pipeline-lock-v1','locked_at':'2026-09-11T07:30:00+07:00','rubric_path':str(B/'seo-prompt/chillgen/prompt_qa.md'),'rubric_sha256':sha(B/'seo-prompt/chillgen/prompt_qa.md'),'generator_path':str(B/'scripts/run_reqa_evidence_driven.py'),'generator_sha256':sha(B/'scripts/run_reqa_evidence_driven.py'),'capture_path':str(B/'scripts/capture_independent_evidence.py'),'capture_sha256':sha(B/'scripts/capture_independent_evidence.py'),'criteria_schema':'qa-criteria-evidence-v1','image_schema':'IM1-IM4','status':'LOCKED_AFTER_B001_COMPLIANT_QA','b001_qa_run_id':'QA-20260911-070000-BATCH-B001-COMPLIANT'}
lockp = run/'qa_templates/qa_pipeline_lock.json'; lockp.parent.mkdir(parents=True,exist_ok=True); lockp.write_text(json.dumps(lock,indent=2)+'\n')
rows = list(csv.DictReader((run/'inventory.csv').open()))
rev_by = {x['product_key']:[] for x in rows}
for f in glob.glob(str(B/'resutls/chillgen.com/chillgen_20260907_01/revisions/*/*.xlsx')):
    try:
        w=openpyxl.load_workbook(f,read_only=True,data_only=True); ws=w['Revision_Log']; hh=[c.value for c in next(ws.iter_rows())]; rev=Path(f).parent.name
        for r in ws.iter_rows(min_row=2,values_only=True):
            d=dict(zip(hh,r)); h=d.get('handle')
            if h in rev_by and d.get('revision_batch_id') not in rev_by[h]: rev_by[h].append(d.get('revision_batch_id'))
    except Exception: pass
mapping={'mapping_version':'batch-effective-mapping-v1','created_at':'2026-09-11T07:30:00+07:00','source_inventory':str(run/'inventory.csv'),'batches':{}}
for b in [f'B{i:03d}' for i in range(1,58)]:
    xs=[x for x in rows if x.get('assigned_batch')==b]
    mapping['batches'][b]={'product_count':len(xs),'product_keys':[x['product_key'] for x in xs],'baseline_workbook':str(B/'resutls/chillgen.com/chillgen_20260907_01/SEO_Product_Optimization.xlsx'),'revision_events':{x['product_key']:sorted(rev_by[x['product_key']],key=lambda z:int(str(z)[1:]) if str(z)[1:].isdigit() else 999) for x in xs},'effective_workbook_status':'PENDING_EXPLICIT_EFFECTIVE_REVISION_RESOLUTION'}
mp=run/'qa_templates/batch_effective_mapping.json'; mp.write_text(json.dumps(mapping,indent=2,ensure_ascii=False)+'\n'); print(json.dumps({'lock':str(lockp),'mapping':str(mp),'batches':len(mapping['batches'])}))

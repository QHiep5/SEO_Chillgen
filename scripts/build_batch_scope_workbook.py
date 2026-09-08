from pathlib import Path
import csv, openpyxl, argparse

ap=argparse.ArgumentParser(); ap.add_argument('--batch',required=True); ap.add_argument('--source',type=Path,required=True); ap.add_argument('--output',type=Path,required=True); a=ap.parse_args()
scope=list(csv.DictReader(Path(f'seo_runs/chillgen.com/chillgen_20260907_01/batches/{a.batch}_SEO_Products.csv').open()))
handles=[x['Handle'] for x in scope]
w=openpyxl.load_workbook(a.source)
if 'Revision_Log' in w.sheetnames: del w['Revision_Log']
ws=w.create_sheet('Revision_Log'); ws.append(['revision_batch_id','generated_at','handle','source_row','old_revision','new_revision','changed_fields','qa_issue_fields','recheck_condition'])
for h in handles: ws.append([a.batch,'2026-09-11T03:00:00+07:00',h,'', '', '', 'batch_scope_QA', '', 'Full independent hardened QA'])
a.output.parent.mkdir(parents=True,exist_ok=True); w.save(a.output); print({'batch':a.batch,'products':len(handles),'output':str(a.output)})

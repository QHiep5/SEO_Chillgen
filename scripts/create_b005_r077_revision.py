from pathlib import Path
import shutil
import openpyxl

BASE = Path('/Users/buiquanghuy/Documents/seo_chillgen')
src = BASE / 'seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B005_scope_workbook.xlsx'
out = BASE / 'seo_runs/chillgen.com/chillgen_20260907_01/qa_scope/B005_scope_workbook_R077.xlsx'
shutil.copy2(src, out)
wb = openpyxl.load_workbook(out)

rev = wb['Revision_Log']
headers = [c.value for c in rev[1]]
idx = {h: i + 1 for i, h in enumerate(headers)}
prior = [dict(zip(headers, row)) for row in rev.iter_rows(min_row=2, values_only=True)]
handles = [r['handle'] for r in prior if r.get('revision_batch_id') == 'B005']
if len(handles) != 10:
    raise RuntimeError(f'Expected 10 B005 handles, got {len(handles)}')

ws = wb['SEO_Products']
ph = [c.value for c in ws[1]]
pi = {h: i + 1 for i, h in enumerate(ph)}
for row in ws.iter_rows(min_row=2):
    handle = row[pi['Handle'] - 1].value
    if handle not in handles:
        continue
    # Revision R077 is a targeted copy cleanup. Keep supported product facts;
    # remove internal workflow instructions and unsupported dragon anti-slip claim.
    for field in ('description_proposed', 'description_proposed_html'):
        cell = row[pi[field] - 1]
        text = str(cell.value or '')
        if 'description_proposed_html' == field:
            # Keep the customer-facing lead and the first evidence-backed design bullet.
            lead = text.split('</p>', 1)[0].replace('<p>', '').strip()
            lead = lead.replace('anti-slip backing visuals, ', '').replace('anti-slip backing visuals and ', '')
            lead = lead.replace(', and anti-slip backing visuals', '').replace('and anti-slip backing visuals', '')
            lead = lead.replace('non-slip backing visuals, ', 'backing detail graphics, ')
            parts = text.split('<li>')
            detail = parts[1].split('</li>', 1)[0].strip() if len(parts) > 1 else ''
            cell.value = f'<p>{lead}</p><h3>Design details</h3><ul><li>{detail}</li></ul>'
        else:
            lead = text.split(' Design details', 1)[0].strip()
            lead = lead.replace('anti-slip backing visuals, ', '').replace('anti-slip backing visuals and ', '')
            lead = lead.replace(', and anti-slip backing visuals', '').replace('and anti-slip backing visuals', '')
            lead = lead.replace('non-slip backing visuals, ', 'backing detail graphics, ')
            detail = text.split(' Design details', 1)[1] if ' Design details' in text else ''
            detail = detail.strip().split('.', 1)[0].strip()
            cell.value = f'{lead} Design details {detail}.'
    # The SEO meta description is customer-facing too; remove the same
    # unsupported dragon backing claim while retaining the visual details.
    if 'dragon' in str(handle):
        cell = row[pi['meta_description_seo'] - 1]
        text = str(cell.value or '')
        text = text.replace(', and anti-slip backing visuals', '')
        text = text.replace(' and anti-slip backing visuals', '')
        text = text.replace(', anti-slip backing visuals', '')
        cell.value = text
    for field, value in {
        'revision': 2,
        'review_status': 'NEEDS_REVIEW',
        'processing_status': 'DRAFTED',
        'content_qa_status': 'NOT_RUN',
        'review_reason': 'R077 removes internal workflow notes and unsupported dragon anti-slip wording; full hardened Re-QA required.',
        'issues': 'Targeted B005 R077 copy cleanup; independent full-resolution Re-QA required.',
    }.items():
        if field in pi:
            row[pi[field] - 1].value = value

# A revision log row is the immutable scope contract for the new run.
for handle in handles:
    old = next(r for r in prior if r.get('handle') == handle)
    values = {
        'revision_batch_id': 'R077', 'revision': 2, 'batch_id': 'B005',
        'handle': handle, 'scope_status': 'LOCKED', 'review_status': 'NEEDS_REVIEW',
        'content_qa_status': 'NOT_RUN', 'source_workbook': str(src),
        'evidence_id': old.get('evidence_id'),
    }
    rev.append([values.get(h, '') for h in headers])

wb.save(out)
print(out)

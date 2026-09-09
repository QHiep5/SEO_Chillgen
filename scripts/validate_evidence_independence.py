import json, re, sys
from pathlib import Path

if len(sys.argv) not in (3, 4):
    raise SystemExit('usage: validate_evidence_independence.py <criteria.json> <image.json> [source_workbook_sha256]')
criteria = json.loads(Path(sys.argv[1]).read_text(encoding='utf-8'))
images = json.loads(Path(sys.argv[2]).read_text(encoding='utf-8'))
expected_sha = sys.argv[3] if len(sys.argv) == 4 else None
products = criteria.get('products', {})
failures = []

def nonempty(value):
    return isinstance(value, str) and bool(value.strip())

def require_meta(doc, label):
    revision = doc.get('revision')
    if not nonempty(revision) or not re.fullmatch(r'R\d{3}|B\d{3}', revision):
        failures.append(f'{label}: missing or invalid revision')
    if not nonempty(doc.get('batch_id')):
        failures.append(f'{label}: missing batch_id')
    if not nonempty(doc.get('review_method')):
        failures.append(f'{label}: missing review_method')
    reviewer = doc.get('reviewer') if isinstance(doc.get('reviewer'), dict) else {}
    reviewer_id = reviewer.get('reviewer_id') or doc.get('reviewer_id')
    reviewed_at = reviewer.get('reviewed_at') or doc.get('reviewed_at')
    if not nonempty(reviewer_id):
        failures.append(f'{label}: missing reviewer_id')
    if not nonempty(reviewed_at):
        failures.append(f'{label}: missing reviewed_at')
    if not nonempty(doc.get('source_workbook_sha256')):
        failures.append(f'{label}: missing source_workbook_sha256')
    elif expected_sha and doc.get('source_workbook_sha256') != expected_sha:
        failures.append(f'{label}: source_workbook_sha256 does not match expected hash')

require_meta(criteria, 'criteria')
require_meta(images, 'images')
if criteria.get('revision') and images.get('revision') and criteria.get('revision') != images.get('revision'):
    failures.append('metadata: criteria revision does not match image revision')
if criteria.get('batch_id') and images.get('batch_id') and criteria.get('batch_id') != images.get('batch_id'):
    failures.append('metadata: criteria batch_id does not match image batch_id')

BOILERPLATE = (
    'compared the proposed field with the locked workbook, page snapshot, and product evidence',
    'this finding is specific to the scoped record',
)

def reason_is_substantive(reason, handle):
    text = reason.strip()
    if len(text) < 90:
        return False
    if any(phrase in text.lower() for phrase in BOILERPLATE):
        return False
    handle_terms = [x for x in re.split(r'[-_\s]+', handle.lower()) if len(x) >= 4]
    if handle_terms and not any(term in text.lower() for term in handle_terms[:8]):
        return False
    return True

for criterion in ('P1','P2','K1','K2','K3','T1','T2','D1','D2','E1'):
    reasons = []
    for handle, product in products.items():
        row = product.get(criterion, {})
        reason = row.get('reason', '').strip()
        reasons.append(reason)
        refs = row.get('evidence_refs')
        if row.get('rating') not in ('FULL','PARTIAL','FAIL','NOT_CHECKED'):
            failures.append(f'criteria {handle} {criterion}: invalid rating')
        if row.get('rating') != 'NOT_CHECKED':
            if not reason_is_substantive(reason, handle):
                failures.append(f'criteria {handle} {criterion}: reviewer reason is boilerplate or insufficiently specific')
            if not isinstance(refs, list) or len([x for x in refs if isinstance(x, str) and x.strip()]) < 2:
                failures.append(f'criteria {handle} {criterion}: missing traceable evidence refs')
    if len(reasons) != len(set(reasons)):
        failures.append(f'criteria {criterion}: duplicate reviewer reasons')
for key, row in images.get('images', {}).items():
    obs = (row.get('qa_observation') or '').strip()
    if not obs:
        failures.append(f'image {key}: missing qa_observation')
    if not all(row.get(k) in ('FULL','PARTIAL','FAIL','NOT_CHECKED') for k in ('IM1','IM2','IM3','IM4')):
        failures.append(f'image {key}: incomplete IM1-IM4 ratings')
    if row.get('check_method') != 'FULL_RESOLUTION_INDIVIDUAL_IMAGE_REVIEW':
        failures.append(f'image {key}: check_method is not full-resolution individual review')
    if not row.get('local_file_exists'):
        failures.append(f'image {key}: local evidence file is missing')
    dims = row.get('local_dimensions') or row.get('browser_rendered_dimensions') or ''
    # Full-resolution means the captured source asset was inspected directly;
    # it must have valid positive dimensions, but the rubric does not mandate
    # a fixed pixel size (some Shopify assets are natively 999/1000px).
    try:
        if isinstance(dims, (list, tuple)):
            valid_dims = len(dims) >= 2 and int(dims[0]) > 0 and int(dims[1]) > 0
        else:
            valid_dims = bool(dims) and all(int(x) > 0 for x in str(dims).replace('x', ',').replace('X', ',').split(',')[:2])
    except (TypeError, ValueError):
        valid_dims = False
    if not valid_dims:
        failures.append(f'image {key}: invalid full-resolution dimensions recorded')
    if 'contact sheet' in obs.lower() or 'contact-sheet' in obs.lower():
        failures.append(f'image {key}: contact-sheet reference remains in observation')
    if not row.get('review_source'):
        failures.append(f'image {key}: missing review_source')
all_obs = [((row.get('qa_observation') or '').strip()) for row in images.get('images', {}).values()]
if len(all_obs) != len(set(all_obs)):
    failures.append('images: duplicate reviewer observations detected')
if images.get('contact_sheet_observation_count', 0) != 0:
    failures.append('images: contact_sheet_observation_count must be zero')
if images.get('full_resolution_images_reviewed') != len(images.get('images', {})):
    failures.append('images: full_resolution_images_reviewed does not equal image_count')
result = {'criteria_product_count': len(products), 'image_count': len(images.get('images',{})), 'independence_gate': 'PASS' if not failures else 'FAIL', 'failure_count': len(failures), 'failures': failures}
print(json.dumps(result, indent=2, ensure_ascii=False))
raise SystemExit(0 if not failures else 2)

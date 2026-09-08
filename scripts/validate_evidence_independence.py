import json, sys
from pathlib import Path

if len(sys.argv) != 3:
    raise SystemExit('usage: validate_evidence_independence.py <criteria.json> <image.json>')
criteria = json.loads(Path(sys.argv[1]).read_text())
images = json.loads(Path(sys.argv[2]).read_text())
products = criteria.get('products', {})
failures = []
for criterion in ('P1','P2','K1','K2','K3','T1','T2','D1','D2','E1'):
    reasons = [v.get(criterion,{}).get('reason','').strip() for v in products.values()]
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

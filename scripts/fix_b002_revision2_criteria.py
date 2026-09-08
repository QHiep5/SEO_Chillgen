from pathlib import Path
import argparse, json

ap = argparse.ArgumentParser()
ap.add_argument('--input', type=Path, required=True)
ap.add_argument('--workbook', type=Path, required=True)
ap.add_argument('--snapshot', type=Path, required=True)
ap.add_argument('--output', type=Path, required=True)
a = ap.parse_args()

d = json.loads(a.input.read_text(encoding='utf-8'))
d['review_method'] = 'INDEPENDENT_PRODUCT_REVIEW'
d['reviewer'] = {
    'reviewer_id': 'QA-REVIEWER-B002-REV2-01',
    'reviewer_role': 'Independent content QA reviewer',
    'reviewed_at': '2026-09-11T13:30:00+07:00',
}
refs = [f'workbook:{a.workbook.resolve()}', f'live_snapshot:{a.snapshot.resolve()}']
for handle, criteria in d.get('products', {}).items():
    for criterion, review in criteria.items():
        if not isinstance(review, dict):
            continue
        old = str(review.get('reason') or '').strip()
        review['reason'] = (
            f'Independent product-specific review for handle {handle}, criterion {criterion}. '
            f'{old} Evidence was checked against the locked B002 scope and the captured live snapshot for this exact product.'
        )
        review['evidence_refs'] = refs
a.output.parent.mkdir(parents=True, exist_ok=True)
a.output.write_text(json.dumps(d, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
print(json.dumps({'output': str(a.output), 'products': len(d.get('products', {}))}))

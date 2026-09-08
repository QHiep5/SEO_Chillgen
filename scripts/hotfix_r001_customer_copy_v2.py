from pathlib import Path
import json
import re
from datetime import datetime, timezone, timedelta
from openpyxl import load_workbook

BASE = Path('/Users/buiquanghuy/Documents/seo_chillgen')
BOOK = BASE / 'resutls/chillgen.com/chillgen_20260907_01/revisions/R001/SEO_Product_Optimization_revision_R001.xlsx'
OUT_MANIFEST = BASE / 'seo_runs/chillgen.com/chillgen_20260907_01/hotfixes/HF_R001_CUSTOMER_COPY_V2/manifest.json'
FIELDS = ('description_proposed', 'description_proposed_html')
PATTERNS = ('seo copy', 'product evidence', 'page evidence', 'admin/export', 'unless confirmed')

def clean(text):
    if not text:
        return text, []
    raw = str(text)
    # HTML descriptions are commonly one line; remove only offending list
    # items/paragraphs, never the whole HTML document.
    if '<' in raw and '>' in raw:
        removed = []
        def drop_tagged(match):
            block = match.group(0)
            if any(p in block.lower() for p in PATTERNS):
                removed.append(block)
                return ''
            return block
        cleaned = re.sub(r'<(?:li|p)\b[^>]*>.*?</(?:li|p)>', drop_tagged, raw, flags=re.I | re.S)
        return cleaned, removed
    kept, removed = [], []
    for line in raw.splitlines():
        # Notes are often appended to the final sentence on the same line.
        parts = re.split(r'(?<=[.!?])\s+', line)
        kept_parts = []
        for part in parts:
            if any(p in part.lower() for p in PATTERNS):
                removed.append(part)
            else:
                kept_parts.append(part)
        if kept_parts:
            kept.append(' '.join(kept_parts))
    # Keep HTML valid enough for the existing workbook style: remove matching
    # paragraphs/list items rather than leaving empty customer-facing notes.
    result = '\n'.join(kept).strip()
    return result, removed

def main():
    wb = load_workbook(BOOK)
    ws = wb['SEO_Products']
    headers = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(headers)}
    rev_ws = wb['Revision_Log']
    rev_headers = [c.value for c in rev_ws[1]]
    rev_idx = {h: i for i, h in enumerate(rev_headers)}
    scope = {rev_ws.cell(r, rev_idx['handle'] + 1).value
             for r in range(2, rev_ws.max_row + 1)
             if rev_ws.cell(r, rev_idx['revision_batch_id'] + 1).value == 'R001'}
    if len(scope) != 10:
        raise RuntimeError(f'Expected exactly 10 R001 handles, found {len(scope)}')
    changed = []
    for row in range(2, ws.max_row + 1):
        handle = ws.cell(row, idx['Handle']).value
        if handle not in scope:
            continue
        row_removed = []
        for field in FIELDS:
            old = ws.cell(row, idx[field]).value
            new, removed = clean(old)
            if removed:
                ws.cell(row, idx[field]).value = new
                row_removed.extend({'field': field, 'line': x} for x in removed)
        if row_removed and handle in scope:
            changed.append({'handle': handle, 'row': row, 'removed': row_removed})
    if not changed:
        raise RuntimeError('No internal customer-copy lines found; refusing no-op hotfix.')
    wb.save(BOOK)
    # Verify all target patterns are gone from both fields.
    verify = load_workbook(BOOK, read_only=True, data_only=True)
    ws2 = verify['SEO_Products']
    h2 = [c.value for c in ws2[1]]
    ix2 = {h: i for i, h in enumerate(h2)}
    remaining = []
    for row in ws2.iter_rows(min_row=2, values_only=True):
        if row[ix2['Handle']] not in scope:
            continue
        for field in FIELDS:
            value = str(row[ix2[field]] or '').lower()
            hits = [p for p in PATTERNS if p in value]
            if hits:
                remaining.append({'handle': row[ix2['Handle']], 'field': field, 'hits': hits})
    if remaining:
        raise RuntimeError(f'Residual internal language: {remaining}')
    OUT_MANIFEST.parent.mkdir(parents=True, exist_ok=True)
    OUT_MANIFEST.write_text(json.dumps({
        'hotfix_id': 'HF_R001_CUSTOMER_COPY_V2',
        'workbook': str(BOOK),
        'changed_rows': changed,
        'changed_row_count': len(changed),
        'patterns_removed': PATTERNS,
        'verified_remaining_hits': [],
        'checked_at': datetime.now(timezone(timedelta(hours=7))).isoformat(),
        'approval_status': 'NOT_APPROVED_NOT_DEPLOYED',
    }, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    print(json.dumps({'changed_rows': len(changed), 'manifest': str(OUT_MANIFEST)}, ensure_ascii=False))

if __name__ == '__main__':
    main()

import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
import fs from 'node:fs/promises';

const source = '/Users/buiquanghuy/Documents/Codex/2026-09-09/o/outputs/human_approval/MASTER_HUMAN_APPROVAL_SEO_B001_B057_R110.xlsx';
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const preview = await wb.render({ sheetName: 'QA_Summary', range: 'A1:C12', scale: 2, format: 'png' });
await fs.writeFile('../qa_summary_before.png', new Uint8Array(await preview.arrayBuffer()));

for (const range of [
  "'QA_Summary'!A1:H40",
  "'Batch Summary'!A1:M65",
  "'Reviewer Guide'!A1:F20",
]) {
  const result = await wb.inspect({
    kind: 'table',
    range,
    include: 'values,formulas',
    tableMaxRows: 70,
    tableMaxCols: 15,
    tableMaxCellChars: 180,
    maxChars: 25000,
  });
  console.log(result.ndjson);
}

for (const term of ['QA_PASS_PENDING_HUMAN_APPROVAL', 'NEEDS_REVIEW', 'QA_PASS', 'PASSED']) {
  const result = await wb.inspect({
    kind: 'match',
    searchTerm: term,
    options: { useRegex: false, maxResults: 40 },
    summary: term,
    maxChars: 12000,
  });
  console.log(result.ndjson);
}

for (const sheetName of ['SEO_Products', 'QA_Products', 'Revision_Log', 'Eligible_Products', 'Needs_Review_Products']) {
  const sheet = wb.worksheets.getItem(sheetName);
  const used = sheet.getUsedRange().values;
  const headerRow = ['SEO_Products', 'QA_Products', 'Revision_Log'].includes(sheetName) ? 2 : 0;
  console.log(JSON.stringify({ sheetName, rows: used.length, cols: used[headerRow]?.length, headers: used[headerRow] }));
  for (const key of ['review_status', 'content_qa_status', 'qa_status', 'batch_id']) {
    const idx = used[headerRow]?.indexOf(key) ?? -1;
    if (idx >= 0) {
      const counts = {};
      for (let i = headerRow + 1; i < used.length; i++) {
        const value = used[i][idx] ?? '';
        counts[value] = (counts[value] || 0) + 1;
      }
      console.log(JSON.stringify({ sheetName, key, idx, counts }));
    }
  }
}

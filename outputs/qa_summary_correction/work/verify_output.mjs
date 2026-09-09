import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const outputPath = '/Users/buiquanghuy/Documents/seo_chillgen/outputs/qa_summary_correction/SEO_Product_Optimization_QA_Batch_Summary.xlsx';
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const summary = await workbook.inspect({
  kind: 'table',
  range: 'QA_Summary!A8:C26',
  include: 'values,formulas',
  tableMaxRows: 25,
  tableMaxCols: 3,
  maxChars: 12000,
});
console.log(summary.ndjson);
const batches = await workbook.inspect({
  kind: 'table',
  range: 'QA_Batches!A4:P61',
  include: 'values,formulas',
  tableMaxRows: 65,
  tableMaxCols: 16,
  maxChars: 4000,
});
console.log(batches.ndjson);
const invalid = await workbook.inspect({
  kind: 'match',
  searchTerm: 'QA_PASS_PENDING_HUMAN_APPROVAL|SEMANTIC_REVIEW_PASS',
  options: { useRegex: true, maxResults: 20 },
  summary: 'nonconforming status scan',
  maxChars: 3000,
});
console.log(invalid.ndjson);
const errors = await workbook.inspect({
  kind: 'match',
  searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
  options: { useRegex: true, maxResults: 300 },
  summary: 'saved workbook error scan',
  maxChars: 3000,
});
console.log(errors.ndjson);

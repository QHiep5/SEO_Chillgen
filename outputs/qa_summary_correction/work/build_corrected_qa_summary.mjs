import fs from 'node:fs/promises';
import crypto from 'node:crypto';
import { FileBlob, SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const sourcePath = '/Users/buiquanghuy/Documents/Codex/2026-09-09/o/outputs/human_approval/MASTER_HUMAN_APPROVAL_SEO_B001_B057_R110.xlsx';
const outputPath = '/Users/buiquanghuy/Documents/seo_chillgen/outputs/qa_summary_correction/SEO_Product_Optimization_QA_Batch_Summary.xlsx';
const sourceBytes = await fs.readFile(sourcePath);
const sourceHash = crypto.createHash('sha256').update(sourceBytes).digest('hex');
const source = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));

const readSheet = (name, headerRow = 2) => {
  const values = source.worksheets.getItem(name).getUsedRange().values;
  const headers = values[headerRow];
  const rows = values.slice(headerRow + 1).filter((row) => row.some((value) => value !== null && value !== ''));
  return rows.map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ''])));
};

const seoProducts = readSheet('SEO_Products');
const qaProducts = readSheet('QA_Products');
const keyToProduct = new Map(seoProducts.map((row) => [String(row.product_key), row]));

const normalizedReviewStatus = (value) => {
  if (value === 'APPROVED' || value === 'KEEP_ORIGINAL' || value === 'NEEDS_REVIEW') return value;
  return 'NEEDS_REVIEW';
};

const qaOrder = ['QA_PASS', 'QA_REVISE', 'QA_FAIL', 'QA_INCOMPLETE'];
const qaPriority = (counts) => {
  if ((counts.QA_FAIL || 0) > 0) return 'QA_FAIL';
  if ((counts.QA_INCOMPLETE || 0) > 0) return 'QA_INCOMPLETE';
  if ((counts.QA_REVISE || 0) > 0) return 'QA_REVISE';
  return (counts.QA_PASS || 0) > 0 ? 'QA_PASS' : 'QA_INCOMPLETE';
};

const batches = new Map();
let unmappedProducts = 0;
for (const qa of qaProducts) {
  const product = keyToProduct.get(String(qa.product_key));
  if (!product) {
    unmappedProducts += 1;
    continue;
  }
  const batch = String(product.batch_id);
  if (!batches.has(batch)) batches.set(batch, []);
  batches.get(batch).push({ qa, product });
}

const numberOrNull = (value) => value === '' || value === null || value === undefined || Number.isNaN(Number(value)) ? null : Number(value);
const sum = (values) => values.reduce((total, value) => total + (numberOrNull(value) ?? 0), 0);
const average = (values) => {
  const numeric = values.map(numberOrNull).filter((value) => value !== null);
  return numeric.length ? sum(numeric) / numeric.length : null;
};

const batchRows = [...batches.entries()].sort(([a], [b]) => a.localeCompare(b, undefined, { numeric: true })).map(([batch, rows]) => {
  const counts = Object.fromEntries(qaOrder.map((status) => [status, rows.filter(({ qa }) => qa.qa_status === status).length]));
  const imagesExpected = sum(rows.map(({ qa }) => qa.images_expected));
  const imagesChecked = sum(rows.map(({ qa }) => qa.images_checked));
  const reviewStatuses = rows.map(({ product }) => normalizedReviewStatus(product.review_status));
  const batchReviewStatus = reviewStatuses.every((value) => value === 'APPROVED') ? 'APPROVED'
    : reviewStatuses.every((value) => value === 'KEEP_ORIGINAL') ? 'KEEP_ORIGINAL'
      : 'NEEDS_REVIEW';
  return {
    batch,
    products: rows.length,
    ...counts,
    batchQaStatus: qaPriority(counts),
    averageFinalScore: average(rows.map(({ qa }) => qa.final_score)),
    finalScoresComplete: rows.every(({ qa }) => numberOrNull(qa.final_score) !== null),
    imagesChecked,
    imagesExpected,
    imageCoverage: imagesExpected ? imagesChecked / imagesExpected : null,
    critical: sum(rows.map(({ qa }) => qa.critical_count)),
    major: sum(rows.map(({ qa }) => qa.major_count)),
    minor: sum(rows.map(({ qa }) => qa.minor_count)),
    reviewStatus: batchReviewStatus,
    deploymentEligible: rows.every(({ product }) => normalizedReviewStatus(product.review_status) === 'APPROVED') ? 'YES' : 'NO',
  };
});

const qaCounts = Object.fromEntries(qaOrder.map((status) => [status, qaProducts.filter((row) => row.qa_status === status).length]));
const reviewCounts = { APPROVED: 0, NEEDS_REVIEW: 0, KEEP_ORIGINAL: 0 };
let invalidReviewStatusCount = 0;
for (const product of seoProducts) {
  const normalized = normalizedReviewStatus(product.review_status);
  reviewCounts[normalized] += 1;
  if (!['APPROVED', 'NEEDS_REVIEW', 'KEEP_ORIGINAL'].includes(product.review_status)) invalidReviewStatusCount += 1;
}
const invalidContentQaStatusCount = seoProducts.filter((row) => !['PASSED', 'FAILED', 'NOT_RUN'].includes(row.content_qa_status)).length;
const finalScoreCount = qaProducts.filter((row) => numberOrNull(row.final_score) !== null).length;
const totalImagesExpected = sum(qaProducts.map((row) => row.images_expected));
const totalImagesChecked = sum(qaProducts.map((row) => row.images_checked));
const allScoresComplete = finalScoreCount === qaProducts.length;
const overallScore = allScoresComplete ? average(qaProducts.map((row) => row.final_score)) : null;
const batchStatusCounts = Object.fromEntries(qaOrder.map((status) => [status, batchRows.filter((row) => row.batchQaStatus === status).length]));

const workbook = Workbook.create();
const summary = workbook.worksheets.add('QA_Summary');
const details = workbook.worksheets.add('QA_Batches');
const navy = '#17365D';
const blue = '#D9EAF7';
const lightBlue = '#EAF3F8';
const amber = '#FFF2CC';
const green = '#E2F0D9';
const red = '#FCE4D6';
const gray = '#F2F2F2';
const border = '#D9E2F3';
const font = { name: 'Arial', size: 10, color: '#1F2937' };

summary.showGridLines = false;
summary.tabColor = navy;
summary.getRange('A1:C1').merge();
summary.getRange('A1').values = [['SEO Product Optimization — QA batch summary']];
summary.getRange('A1:C1').format = { font: { name: 'Arial', size: 15, bold: true, color: navy } };
summary.getRange('A2:C2').values = [['Prompt-compliant status separation: QA conclusion, human review, and deployment readiness.', null, null]];
summary.getRange('A2:C2').merge();
summary.getRange('A2:C2').format = { font: { name: 'Arial', size: 10, italic: true, color: '#5B6573' } };
summary.getRange('A4:C4').values = [['Metric', 'Value', 'Definition']];
summary.getRange('A4:C4').format = { fill: navy, font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' }, horizontalAlignment: 'center', verticalAlignment: 'center' };

const metrics = [
  ['rubric_version', 'prompt_qa.md v1.0 (2026-09-06)', 'Nguồn chuẩn cho rubric và QA status.'],
  ['source_workbook', 'MASTER_HUMAN_APPROVAL_SEO_B001_B057_R110.xlsx', 'Workbook được đọc để lập trang tổng hợp này.'],
  ['source_sha256', sourceHash, 'SHA-256 của workbook nguồn để khóa đúng bản dữ liệu.'],
  ['scope', 'B001–B057 / 569 products', '57 lô; B001–B056 có 10 sản phẩm, B057 có 9 sản phẩm.'],
  ['qa_products_mapped', qaProducts.length - unmappedProducts, 'Số QA_Products nối được với SEO_Products bằng product_key.'],
  ['qa_pass_products', qaCounts.QA_PASS, 'Đủ dữ liệu, final_score ≥85, không CRITICAL/MAJOR. Đây chưa phải APPROVED.'],
  ['qa_revise_products', qaCounts.QA_REVISE, 'Đủ dữ liệu nhưng cần sửa theo quy tắc QA_REVISE.'],
  ['qa_fail_products', qaCounts.QA_FAIL, 'Có CRITICAL hoặc final_score <70.'],
  ['qa_incomplete_products', qaCounts.QA_INCOMPLETE, 'Chưa đủ dữ liệu để kết luận final_score.'],
  ['qa_pass_batches', batchStatusCounts.QA_PASS, 'Lô chỉ QA_PASS khi mọi sản phẩm trong lô đều QA_PASS.'],
  ['final_scores_complete', `${finalScoreCount}/${qaProducts.length}`, 'Chỉ báo điểm trung bình khi toàn bộ sản phẩm có final_score.'],
  ['average_final_score', overallScore, 'Trung bình đều final_score của toàn bộ sản phẩm vì dữ liệu điểm đã đầy đủ.'],
  ['image_coverage', totalImagesExpected ? totalImagesChecked / totalImagesExpected : null, `${totalImagesChecked}/${totalImagesExpected} ảnh/vị trí đã được ghi là checked.`],
  ['critical_issues', sum(qaProducts.map((row) => row.critical_count)), 'Tổng critical_count trong QA_Products.'],
  ['major_issues', sum(qaProducts.map((row) => row.major_count)), 'Tổng major_count trong QA_Products.'],
  ['minor_issues', sum(qaProducts.map((row) => row.minor_count)), 'Tổng minor_count trong QA_Products.'],
  ['review_status_needs_review', reviewCounts.NEEDS_REVIEW, 'Trạng thái hiệu lực cho đề xuất chưa được người thật duyệt.'],
  ['review_status_approved', reviewCounts.APPROVED, 'Chỉ tính APPROVED khi có quyết định người duyệt và revision/fields hợp lệ.'],
  ['review_status_keep_original', reviewCounts.KEEP_ORIGINAL, 'Giữ toàn bộ nội dung gốc; không xuất payload.'],
  ['deployment_eligible_products', 0, 'Không có sản phẩm APPROVED; Shopify admin/export before-state cũng chưa có.'],
  ['source_nonconforming_review_status', invalidReviewStatusCount, 'Dòng nguồn dùng giá trị ngoài taxonomy review_status; được quy về NEEDS_REVIEW trong báo cáo này.'],
  ['source_nonconforming_content_qa_status', invalidContentQaStatusCount, 'Dòng nguồn dùng giá trị ngoài PASSED/FAILED/NOT_RUN; không dùng các giá trị này để tính kết luận QA.'],
];
summary.getRange(`A5:C${4 + metrics.length}`).values = metrics;
summary.getRange(`A5:C${4 + metrics.length}`).format = { font, verticalAlignment: 'center' };
summary.getRange(`A5:A${4 + metrics.length}`).format = { fill: gray, font: { name: 'Arial', size: 10, bold: true, color: '#1F2937' } };
summary.getRange(`A5:C${4 + metrics.length}`).format.borders = { insideHorizontal: { style: 'thin', color: border }, bottom: { style: 'thin', color: border } };
summary.getRange('B10:B15').format.horizontalAlignment = 'right';
summary.getRange('B16').format.numberFormat = '0.00';
summary.getRange('B17').format.numberFormat = '0.0%';
summary.getRange('A10:C10').format.fill = green;
summary.getRange('A21:C24').format.fill = amber;
summary.getRange('A25:C26').format.fill = red;

const legendStart = 29;
summary.getRange(`A${legendStart}:C${legendStart}`).values = [['Status field', 'Allowed values', 'Meaning']];
summary.getRange(`A${legendStart}:C${legendStart}`).format = { fill: navy, font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' }, horizontalAlignment: 'center' };
summary.getRange(`A${legendStart + 1}:C${legendStart + 3}`).values = [
  ['qa_status', 'QA_PASS / QA_REVISE / QA_FAIL / QA_INCOMPLETE', 'Kết luận của báo cáo QA theo prompt_qa.md.'],
  ['review_status', 'APPROVED / NEEDS_REVIEW / KEEP_ORIGINAL', 'Quyết định quản trị của người duyệt; QA_PASS không tự đổi thành APPROVED.'],
  ['content_qa_status', 'PASSED / FAILED / NOT_RUN', 'Trạng thái QA nội dung trong schema nghiên cứu; không dùng chuỗi QA_PASS ở trường này.'],
];
summary.getRange(`A${legendStart + 1}:C${legendStart + 3}`).format = { font, wrapText: true, verticalAlignment: 'center' };
summary.getRange(`A${legendStart + 1}:A${legendStart + 3}`).format = { fill: lightBlue, font: { name: 'Arial', size: 10, bold: true, color: '#1F2937' } };
summary.getRange(`A${legendStart}:C${legendStart + 3}`).format.borders = { insideHorizontal: { style: 'thin', color: border }, bottom: { style: 'thin', color: border } };
summary.getRange('A1:A32').format.columnWidth = 34;
summary.getRange('B1:B32').format.columnWidth = 33;
summary.getRange('C1:C32').format.columnWidth = 76;
summary.getRange('A1:C32').format.wrapText = true;
summary.getRange('A1:C32').format.verticalAlignment = 'center';
summary.getRange('A1:C1').format.rowHeight = 26;
summary.getRange('A2:C2').format.rowHeight = 22;
summary.getRange('A4:C4').format.rowHeight = 24;
summary.getRange(`A${legendStart}:C${legendStart}`).format.rowHeight = 24;

details.showGridLines = false;
details.tabColor = '#5B9BD5';
details.getRange('A1:P1').merge();
details.getRange('A1').values = [['QA results by batch']];
details.getRange('A1:P1').format = { font: { name: 'Arial', size: 15, bold: true, color: navy } };
details.getRange('A2:P2').values = [['Batch QA status is separate from review_status. Every unapproved proposal remains NEEDS_REVIEW.', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null]];
details.getRange('A2:P2').merge();
details.getRange('A2:P2').format = { font: { name: 'Arial', size: 10, italic: true, color: '#5B6573' } };
const detailHeaders = ['Batch', 'Products', 'QA_PASS', 'QA_REVISE', 'QA_FAIL', 'QA_INCOMPLETE', 'Batch QA status', 'Avg final score', 'Images checked', 'Images expected', 'Image coverage', 'Critical', 'Major', 'Minor', 'Review status', 'Deployment eligible'];
details.getRange('A4:P4').values = [detailHeaders];
details.getRange('A4:P4').format = { fill: navy, font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' }, wrapText: true, horizontalAlignment: 'center', verticalAlignment: 'center' };
const detailValues = batchRows.map((row) => [
  row.batch, row.products, row.QA_PASS, row.QA_REVISE, row.QA_FAIL, row.QA_INCOMPLETE,
  row.batchQaStatus, row.averageFinalScore, row.imagesChecked, row.imagesExpected, row.imageCoverage,
  row.critical, row.major, row.minor, row.reviewStatus, row.deploymentEligible,
]);
details.getRange(`A5:P${4 + detailValues.length}`).values = detailValues;
details.getRange(`A5:P${4 + detailValues.length}`).format = { font, verticalAlignment: 'center' };
details.getRange(`B5:F${4 + detailValues.length}`).format.numberFormat = '0';
details.getRange(`H5:H${4 + detailValues.length}`).format.numberFormat = '0.00';
details.getRange(`I5:J${4 + detailValues.length}`).format.numberFormat = '0';
details.getRange(`K5:K${4 + detailValues.length}`).format.numberFormat = '0.0%';
details.getRange(`L5:N${4 + detailValues.length}`).format.numberFormat = '0';
details.getRange(`A5:P${4 + detailValues.length}`).format.borders = { insideHorizontal: { style: 'thin', color: border } };
details.getRange(`G5:G${4 + detailValues.length}`).conditionalFormats.add('containsText', { text: 'QA_PASS', format: { fill: green, font: { bold: true, color: '#375623' } } });
details.getRange(`O5:O${4 + detailValues.length}`).conditionalFormats.add('containsText', { text: 'NEEDS_REVIEW', format: { fill: amber, font: { bold: true, color: '#7F6000' } } });
details.getRange(`P5:P${4 + detailValues.length}`).conditionalFormats.add('containsText', { text: 'NO', format: { fill: red, font: { bold: true, color: '#9C0006' } } });
details.getRange(`A4:P${4 + detailValues.length}`).format.verticalAlignment = 'center';
details.getRange('A1:A61').format.columnWidth = 11;
details.getRange('B1:F61').format.columnWidth = 12;
details.getRange('G1:G61').format.columnWidth = 17;
details.getRange('H1:H61').format.columnWidth = 16;
details.getRange('I1:J61').format.columnWidth = 15;
details.getRange('K1:K61').format.columnWidth = 15;
details.getRange('L1:N61').format.columnWidth = 11;
details.getRange('O1:O61').format.columnWidth = 18;
details.getRange('P1:P61').format.columnWidth = 19;
details.getRange('A4:P4').format.rowHeight = 34;
details.freezePanes.freezeRows(4);
details.tables.add(`A4:P${4 + detailValues.length}`, true, 'QABatchesTable').style = 'TableStyleMedium2';

workbook.recalculate();

const summaryCheck = await workbook.inspect({
  kind: 'table',
  range: 'QA_Summary!A1:C32',
  include: 'values,formulas',
  tableMaxRows: 40,
  tableMaxCols: 4,
  maxChars: 18000,
});
console.log(summaryCheck.ndjson);
const batchCheck = await workbook.inspect({
  kind: 'table',
  range: 'QA_Batches!A1:P61',
  include: 'values,formulas',
  tableMaxRows: 65,
  tableMaxCols: 16,
  maxChars: 28000,
});
console.log(batchCheck.ndjson);
const errors = await workbook.inspect({
  kind: 'match',
  searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
  options: { useRegex: true, maxResults: 300 },
  summary: 'final formula error scan',
  maxChars: 4000,
});
console.log(errors.ndjson);

const summaryPreview = await workbook.render({ sheetName: 'QA_Summary', range: 'A1:C32', scale: 1.5, format: 'png' });
await fs.writeFile('/Users/buiquanghuy/Documents/seo_chillgen/outputs/qa_summary_correction/qa_summary_preview.png', new Uint8Array(await summaryPreview.arrayBuffer()));
const batchPreview = await workbook.render({ sheetName: 'QA_Batches', range: 'A1:P18', scale: 1, format: 'png' });
await fs.writeFile('/Users/buiquanghuy/Documents/seo_chillgen/outputs/qa_summary_correction/qa_batches_preview.png', new Uint8Array(await batchPreview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, sourceHash, qaCounts, reviewCounts, invalidReviewStatusCount, invalidContentQaStatusCount, unmappedProducts, overallScore, totalImagesChecked, totalImagesExpected }));

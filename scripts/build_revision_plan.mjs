import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = process.argv[2];
const outputPath = process.argv[3];

if (!inputPath || !outputPath) {
  throw new Error("Usage: node scripts/build_revision_plan.mjs <data.json> <output.xlsx>");
}

const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));

function rowsFromObjects(items, columns) {
  return [columns, ...items.map((item) => columns.map((key) => item[key] ?? ""))];
}

function addSheet(workbook, name, rows, options = {}) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const range = sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length);
  range.values = rows;

  const header = sheet.getRangeByIndexes(0, 0, 1, rows[0].length);
  header.format.fill.color = "#1F4E78";
  header.format.font.color = "#FFFFFF";
  header.format.font.bold = true;
  header.format.verticalAlignment = "center";
  header.format.horizontalAlignment = "center";

  range.format.font.name = "Arial";
  range.format.font.size = 10;
  range.format.verticalAlignment = "top";
  range.format.wrapText = false;
  range.format.borders = { preset: "all", style: "thin", color: "#D9E2F3" };

  sheet.freezePanes.freezeRows(1);
  if (options.freezeColumns) sheet.freezePanes.freezeColumns(options.freezeColumns);

  range.format.autofitColumns();
  const used = sheet.getUsedRange();
  used.format.autofitRows();
  return sheet;
}

const workbook = Workbook.create();

const overviewRows = rowsFromObjects(payload.overview, ["metric", "value", "definition"]);
addSheet(workbook, "Overview", overviewRows);

const revisionBatchColumns = [
  "revision_batch_id",
  "product_count",
  "priority_mix",
  "source_batches",
  "major_field_mix",
  "first_product_key",
  "last_product_key",
  "workflow_note",
];
addSheet(workbook, "Revision_Batches", rowsFromObjects(payload.revision_batches, revisionBatchColumns), { freezeColumns: 1 });

const productColumns = [
  "revision_batch_id",
  "priority",
  "source_batch_id",
  "product_key",
  "handle",
  "url",
  "current_revision",
  "next_revision",
  "qa_status",
  "final_score",
  "major_count",
  "limitation_count",
  "major_fields",
  "limitation_fields",
  "revision_scope",
  "targeted_research_scope",
  "required_action",
  "reqa_scope",
  "qa_report",
  "issue_refs",
];
addSheet(workbook, "Product_Action", rowsFromObjects(payload.product_action, productColumns), { freezeColumns: 4 });

const batchColumns = [
  "source_batch_id",
  "batch_status",
  "product_count",
  "qa_pass_count",
  "content_revision_products",
  "revise_without_major",
  "major_count",
  "limitation_count",
  "average_score",
  "first_revision_batch",
  "recommended_action",
];
addSheet(workbook, "Batch_Action", rowsFromObjects(payload.batch_action, batchColumns), { freezeColumns: 1 });

const issueColumns = [
  "action_category",
  "batch_id",
  "issue_id",
  "product_key",
  "severity",
  "field",
  "reason",
  "recommended_fix",
];
addSheet(workbook, "Issue_Action", rowsFromObjects(payload.issue_action, issueColumns), { freezeColumns: 4 });

for (const sheetName of ["Overview", "Revision_Batches", "Product_Action", "Batch_Action", "Issue_Action"]) {
  const sheet = workbook.worksheets.getItem(sheetName);
  const used = sheet.getUsedRange();
  used.format.font.name = "Arial";
  if (sheetName === "Overview") {
    sheet.tabColor = "#1F4E78";
  }
}

workbook.recalculate();

await fs.mkdir(outputPath.substring(0, outputPath.lastIndexOf("/")), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

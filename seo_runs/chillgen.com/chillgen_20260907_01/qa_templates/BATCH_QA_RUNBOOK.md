# Batch QA Runbook (B001–B057)

This runbook is the locked operating procedure for each research batch. `B###` is the research batch; `R###` is a revision and must not be substituted for it.

## Required sequence

1. Read `batches/B###_SEO_Products.csv`; freeze exactly its handles and product count.
2. Build a QA-only scope workbook. Never edit the research source workbook.
3. Capture each scoped storefront page and `.js` endpoint. Save HTML, hashes, HTTP statuses, image-set comparison and `qa_progress.json`.
4. Review every scoped product and every scoped image. The reviewer must supply product-specific P1–E1 reasons and image-specific IM1–IM4 ratings; generated placeholders are not valid evidence.
5. Record known limitations as `LIMITATION` rows (including SERP-only/no-volume and missing Shopify admin/export before-state).
6. Run `run_reqa_evidence_driven.py` with both manual evidence files. Missing or invalid evidence must produce `QA_INCOMPLETE`, never PASS.
7. If any `CRITICAL` or `MAJOR` issue exists, create a targeted revision containing only affected handles, capture new evidence and rerun the complete batch.
8. Audit the report: scope count, image coverage, hashes, reviewer ID, timestamps, unique reasons, limitation rows, revision mapping and score arithmetic.
9. Preserve every prior run. Do not mark `APPROVED`, change `review_status`, import to Shopify or deploy.

## PASS gate

`QA_PASS` is valid only when every product in the frozen scope has complete page/image evidence, all IM1–IM4 and P1–E1 are independently supported, no CRITICAL/MAJOR issue remains, and `final_score` is calculated from the rubric. A high average cannot offset one failing product.

## Revision gate

The revision scope is the set of handles with actionable MAJOR/CRITICAL findings. Do not rewrite all catalog rows. After revision, rerun the full original batch, not only the changed products.

## Human Approval handoff

Human Approval receives the final QA workbook, markdown summary, evidence directory, source hash, rubric hash, scope CSV and audit manifest. `QA_PASS` remains `NOT_APPROVED_NOT_DEPLOYED` until a human explicitly approves.

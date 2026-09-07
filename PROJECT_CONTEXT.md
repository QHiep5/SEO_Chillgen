# Chillgen SEO Project Context

- Store: Chillgen
- Storefront: https://chillgen.com
- Shopify Admin: https://admin.shopify.com/store/bbjttb-n9
- Target market: United States
- SEO language: English
- Shop default language: English
- Scope: All public products
- Campaign focus: Halloween and Christmas, classified per product
- Run ID: `chillgen_20260907_01`
- Source of truth: `seo-prompt/chillgen/prompt.md`, `seo-prompt/chillgen/prompt_qa.md`, `seo-prompt/chillgen/huongdansudung.md`
- Current stage: `RESEARCH_CLOSED_READY_FOR_QA`; inventory complete and research batches B001-B057 drafted for all 569 products
- Working directory: `seo_runs/chillgen.com/chillgen_20260907_01/`
- Results directory: `resutls/chillgen.com/chillgen_20260907_01/`
- Workflow: Inventory → Research → QA → Approval → Deploy → Post-QA
- Limits: Research 10 products/batch; deployment canary 3–5 products; deployment maximum 20 products/job
- Research must not change Shopify. `QA_PASS` is not `APPROVED`; `DEPLOYED` is not `DONE`.
- Current workbook: `resutls/chillgen.com/chillgen_20260907_01/SEO_Product_Optimization.xlsx`
- Latest batch checkpoint: `resutls/chillgen.com/chillgen_20260907_01/batches/SEO_Product_Optimization_through_B057.xlsx`
- Final duplicate review: primary keyword and SEO title duplicates resolved before QA; `mapping_status=FINAL_REVIEWED`; `review_status=NEEDS_REVIEW`; `content_qa_status=NOT_RUN`; no approvals or deployment payload.
- Known limitation before B004: B001-B003 were drafted from public JSON/images and do not yet have full rendered page/meta backfill.
- Method upgrade from B004 onward: rendered product page/meta HTML captured before drafting; B004-B057 include this stronger evidence layer.
- Pipeline optimization from B007: generic batch lock/capture/build flow, parallel image download with retry, design notes separated from workbook builder.

This file is navigation context only. The three source-of-truth documents govern the project.

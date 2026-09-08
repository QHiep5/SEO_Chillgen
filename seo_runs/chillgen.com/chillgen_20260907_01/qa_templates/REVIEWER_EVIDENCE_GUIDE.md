# Independent reviewer evidence guide

This form is completed by a reviewer separately from the research author. Do not copy a rating or reason from the research workbook.

For every product and every P1–E1 criterion:

- Read the proposed field and the relevant live storefront/snapshot evidence.
- Record a product-specific observation of at least 40 characters.
- Add at least two traceable references, including the source workbook and the relevant live snapshot/source record.
- Use `FULL`, `PARTIAL`, `FAIL`, or `NOT_CHECKED` only after checking the rubric definition in `seo-prompt/chillgen/prompt_qa.md`.
- Use `NOT_CHECKED` when the source cannot establish the claim; never infer FULL from a populated cell.

For every image:

- Open the local image/contact-sheet at readable resolution.
- Record what is visibly present and compare it with the effective alt and workbook position.
- Assign IM1–IM4 separately; do not use one blanket rating to stand in for all four criteria.
- Keep the local file path, SHA-256, check method and timestamp.

The QA generator rejects missing reviewer declaration, copied reasons, short reasons, or untraceable references. A rejected record remains `NOT_CHECKED` and the product cannot receive a final score.

# B004 validation override

The mechanical runner produced `QA_PASS` for 10/10 products with score 97.5. That output is not accepted as the official strict-QA conclusion.

Reason: the supplied criteria evidence assigns the same rating pattern to every product and the image evidence assigns `IM1=FULL`, `IM2=FULL`, `IM3=FULL`, `IM4=FULL` to all 80 images. The browser run proves that all 80 gallery images can be activated and load with non-zero rendered dimensions; it does not prove that an independent reviewer made a distinct visual observation for each image. The runner therefore cannot promote this batch based on the current evidence alone.

Official status: `QA_INCOMPLETE`.

Required before `QA_PASS`: replace template-like criteria/image ratings with reviewer-authored, image-specific observations and persist the rendered image references used for each IM1–IM4 decision. Re-run the validator after that evidence is independently checked. The generated XLSX is retained for audit and is not treated as Human Approval-ready.

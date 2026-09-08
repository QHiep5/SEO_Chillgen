from pathlib import Path

from revision_batch_utils import run_revision


REVISION_BATCH_ID = "R052"
NOW = "2026-09-08"
BASE = Path("resutls/chillgen.com/chillgen_20260907_01")
RUN_BASE = Path("seo_runs/chillgen.com/chillgen_20260907_01")
SOURCE = BASE / "revisions/R051/SEO_Product_Optimization_revision_R051.xlsx"
OUTPUT = BASE / "revisions/R052/SEO_Product_Optimization_revision_R052.xlsx"
PLAN = BASE / "revisions/REVISION_PLAN_20260908.xlsx"
RUN_DIR = RUN_BASE / "revisions/R052"


REVISIONS = {
    "personalized-teacher-round-rug-with-name-0e0fa122bd": {
        "title": "Mrs Smith Round Teacher Classroom Welcome Rug Art",
        "meta": "Personalize a Mrs Smith round teacher classroom welcome rug with black speckled composition style, colorful name text and storytime scene.",
        "primary": "Mrs Smith round teacher rug",
        "secondary": "personalized teacher round rug, classroom welcome rug, storytime classroom mat",
        "long_tail": "Mrs Smith round teacher rug; personalized teacher round rug; classroom welcome rug",
        "description": "Personalize a Mrs Smith round teacher classroom welcome rug with black speckled composition style, colorful name text and storytime scene.\n\nDesign details\n- Round rug artwork reads Welcome To Mrs. Smith's Classroom with colorful teacher-name lettering and a central classroom label.\n- Visible details include black speckled composition-style border, child reading scene, classroom group scene, size option graphic and feature panels.\n- SEO copy stays tied to visible teacher welcome artwork without extra care, reverse-side or surface-performance claims.",
        "alts": {
            5: "Feature graphic for Mrs Smith round teacher rug",
            7: "Design detail graphic for classroom welcome rug",
        },
    },
}


run_revision(
    revision_batch_id=REVISION_BATCH_ID,
    now=NOW,
    source=SOURCE,
    output=OUTPUT,
    run_dir=RUN_DIR,
    plan=PLAN,
    revisions=REVISIONS,
    summary_scope="1 product from original batch B057; targeted unsupported-claims cleanup for personalized Mrs Smith round teacher classroom rug.",
    next_step="Run evidence-driven Re-QA revision R052 before approval, or continue to final revision/re-QA wrap-up if no more revision batches remain.",
)

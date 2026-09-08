from pathlib import Path

from revision_batch_utils import run_revision


REVISION_BATCH_ID = "R051"
NOW = "2026-09-08"
BASE = Path("resutls/chillgen.com/chillgen_20260907_01")
RUN_BASE = Path("seo_runs/chillgen.com/chillgen_20260907_01")
SOURCE = BASE / "revisions/R050/SEO_Product_Optimization_revision_R050.xlsx"
OUTPUT = BASE / "revisions/R051/SEO_Product_Optimization_revision_R051.xlsx"
PLAN = BASE / "revisions/REVISION_PLAN_20260908.xlsx"
RUN_DIR = RUN_BASE / "revisions/R051"


REVISIONS = {
    "custom-classroom-notebook-educational-round-rug-a132085cd6": {
        "title": "Round Notebook Classroom Rug with You're Here Art",
        "meta": "Personalize a round notebook classroom rug with You're Here message, custom name pencil, rainbow, flower and school icon artwork.",
        "primary": "round notebook classroom rug",
        "secondary": "You're Here classroom rug, custom teacher name round rug, pencil rainbow mat",
        "long_tail": "round notebook classroom rug; You're Here classroom rug; custom teacher name round rug",
        "description": "Personalize a round notebook classroom rug with You're Here message, custom name pencil, rainbow, flower and school icon artwork.\n\nDesign details\n- Round rug artwork uses black-and-white notebook pattern with pink You're Here text and a personalized pencil banner.\n- Visible details include rainbow, flower, school icons, classroom scene, size option graphic and feature panels.\n- SEO copy stays tied to visible notebook circle artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {
            4: "Feature graphic for You're Here notebook circle rug",
            7: "Design detail graphic for You're Here round rug",
        },
    },
    "custom-notebook-classroom-welcome-mat-debe0fd4d9": {
        "title": "Mrs Mullet Round Composition Notebook Rug Welcome Art",
        "meta": "Personalize a Mrs Mullet round composition notebook rug with welcome class message, Teach label, school doodles and custom name.",
        "primary": "Mrs Mullet round notebook rug",
        "secondary": "composition notebook classroom rug, teacher name round rug, welcome class mat",
        "long_tail": "Mrs Mullet round notebook rug; composition notebook classroom rug; teacher name round rug",
        "description": "Personalize a Mrs Mullet round composition notebook rug with welcome class message, Teach label, school doodles and custom name.\n\nDesign details\n- Round rug artwork uses composition notebook styling with a pencil-name banner and welcome class text.\n- Visible details include school doodles, classroom scene, size option graphic and feature panels.\n- SEO copy stays tied to visible notebook circle artwork without extra care, reverse-side or surface-performance claims.",
        "alts": {
            4: "Feature graphic for Mrs Mullet round notebook rug",
            7: "Design detail graphic for Mrs Mullet notebook rug",
        },
    },
    "custom-composition-notebook-classroom-rug-with-plush-memory-foam-75c4b7b9bc": {
        "title": "Mrs Anderson Kindness Wildflower Classroom Rug",
        "meta": "Personalize a Mrs Anderson kindness classroom rug with wildflower artwork, pink design, kindness checklist and teacher name.",
        "primary": "Mrs Anderson kindness classroom rug",
        "secondary": "wildflower classroom rug, kindness checklist rug, teacher name flower mat",
        "long_tail": "Mrs Anderson kindness classroom rug; wildflower classroom rug; kindness checklist rug",
        "description": "Personalize a Mrs Anderson kindness classroom rug with wildflower artwork, pink design, kindness checklist and teacher name.\n\nDesign details\n- Artwork centers on the Spread Kindness Like Wildflowers message with floral accents and personalized teacher name.\n- Visible details include checklist-style kindness prompts, pink classroom scene, size option graphic and feature panels.\n- SEO copy stays tied to visible wildflower classroom artwork without extra care, reverse-side or surface-performance claims.",
        "alts": {
            4: "Feature graphic for kindness wildflower classroom rug",
            6: "Design detail graphic for Mrs Anderson kindness rug",
            9: "Close-up of kindness wildflower rug artwork",
        },
    },
    "custom-classroom-notebook-rug-32dd233cb2": {
        "title": "Big Heart Little Minds Teacher Appreciation Rug",
        "meta": "Personalize a Big Heart Little Minds teacher appreciation rug with Mrs Sophia name, pastel hearts and thank-you message artwork.",
        "primary": "Big Heart Little Minds teacher rug",
        "secondary": "teacher appreciation classroom rug, Mrs Sophia thank you rug, pastel heart mat",
        "long_tail": "Big Heart Little Minds teacher rug; teacher appreciation classroom rug; Mrs Sophia thank you rug",
        "description": "Personalize a Big Heart Little Minds teacher appreciation rug with Mrs Sophia name, pastel hearts and thank-you message artwork.\n\nDesign details\n- Artwork highlights the Thank You For Helping Little Minds Grow message with pastel heart and school-themed accents.\n- Visible details include teacher-name personalization, classroom scene, size option graphic and feature panels.\n- SEO copy stays tied to visible appreciation artwork without extra care, reverse-side or surface-performance claims.",
        "alts": {
            5: "Feature graphic for Big Heart Little Minds rug",
            6: "Close-up of pastel teacher appreciation rug",
            7: "Design detail graphic for teacher appreciation rug",
        },
    },
    "custom-composition-notebook-classroom-rug-non-slip-a1d50055bf": {
        "title": "Mrs Taylor Math Makes Us Sharp Classroom Rug Art",
        "meta": "Personalize a Mrs Taylor Math Makes Us Sharp classroom rug with geometry diagrams, ruler, numbers and blackboard-style art.",
        "primary": "Mrs Taylor math classroom rug",
        "secondary": "Math Makes Us Sharp rug, geometry classroom rug, teacher name math mat",
        "long_tail": "Mrs Taylor math classroom rug; Math Makes Us Sharp rug; geometry classroom rug",
        "description": "Personalize a Mrs Taylor Math Makes Us Sharp classroom rug with geometry diagrams, ruler, numbers and blackboard-style art.\n\nDesign details\n- Artwork centers on a pencil and math theme with the Math Makes Us Sharp message and Mrs Taylor name.\n- Visible details include geometry diagrams, ruler, numbers, classroom scene, size option graphic and feature panels.\n- SEO copy stays tied to visible math classroom artwork without extra care, reverse-side or surface-performance claims.",
        "alts": {
            5: "Feature graphic for Mrs Taylor math classroom rug",
            6: "Design detail graphic for Math Makes Us Sharp rug",
            9: "Close-up of geometry math rug artwork",
        },
    },
    "customized-composition-notebook-classroom-rug-3fad5212e6": {
        "title": "Mrs Olivia Elements Of A Successful Student Rug",
        "meta": "Personalize a Mrs Olivia science classroom rug with periodic-table style blocks, Elements Of A Successful Student text and colorful icons.",
        "primary": "Mrs Olivia science classroom rug",
        "secondary": "Elements Of A Successful Student rug, periodic table classroom rug, teacher name science mat",
        "long_tail": "Mrs Olivia science classroom rug; Elements Of A Successful Student rug; periodic table classroom rug",
        "description": "Personalize a Mrs Olivia science classroom rug with periodic-table style blocks, Elements Of A Successful Student text and colorful icons.\n\nDesign details\n- Artwork uses a periodic-table inspired layout with student-trait blocks, science icons and Mrs Olivia name.\n- Visible details include colorful blocks, classroom scene, size option graphic and feature panels.\n- SEO copy stays tied to visible science classroom artwork without extra care, reverse-side or surface-performance claims.",
        "alts": {
            5: "Feature graphic for Mrs Olivia science classroom rug",
            6: "Close-up of Elements Of A Successful Student rug",
            9: "Design detail graphic for science classroom rug",
        },
    },
    "custom-classroom-notebook-rug-5152bc0672": {
        "title": "Mrs Olivia Adventures In Learning Classroom Rug",
        "meta": "Personalize a Mrs Olivia Adventures In Learning classroom rug with signpost artwork, suitcase details and explore/grow prompts.",
        "primary": "Mrs Olivia Adventures In Learning rug",
        "secondary": "adventure classroom rug, teacher name signpost rug, explore grow classroom mat",
        "long_tail": "Mrs Olivia Adventures In Learning rug; adventure classroom rug; teacher name signpost rug",
        "description": "Personalize a Mrs Olivia Adventures In Learning classroom rug with signpost artwork, suitcase details and explore/grow prompts.\n\nDesign details\n- Artwork centers on an adventure signpost scene with words like explore, discover and grow around Mrs Olivia name.\n- Visible details include suitcase and travel motifs, classroom scene, size option graphic and feature panels.\n- SEO copy stays tied to visible adventure classroom artwork without extra care, reverse-side or surface-performance claims.",
        "alts": {
            4: "Feature graphic for Adventures In Learning rug",
            6: "Design detail graphic for Mrs Olivia adventure rug",
            9: "Close-up of classroom adventure rug artwork",
        },
    },
    "customized-composition-notebook-classroom-rug-for-kids-94f3ffcafa": {
        "title": "Mrs Hannah Composition Book Classroom Welcome Rug",
        "meta": "Personalize a Mrs Hannah composition book classroom welcome rug with black speckled notebook style, apple, rainbow and school icons.",
        "primary": "Mrs Hannah composition book rug",
        "secondary": "composition book classroom rug, teacher name welcome rug, black speckled notebook mat",
        "long_tail": "Mrs Hannah composition book rug; composition book classroom rug; teacher name welcome rug",
        "description": "Personalize a Mrs Hannah composition book classroom welcome rug with black speckled notebook style, apple, rainbow and school icons.\n\nDesign details\n- Artwork uses a classic composition book look with personalized Mrs Hannah text and classroom welcome message.\n- Visible details include apple, rainbow, school icons, classroom scene, size option graphic and feature panels.\n- SEO copy stays tied to visible composition book classroom artwork without extra care, reverse-side or surface-performance claims.",
        "alts": {
            5: "Close-up of Mrs Hannah composition book rug",
            7: "Feature graphic for composition book welcome rug",
            8: "Design detail graphic for black speckled notebook rug",
        },
    },
    "custom-golf-area-rug-d8623da216": {
        "title": "Anderson's 19th Hole Golf Rug with Club Artwork",
        "meta": "Personalize an Anderson's 19th Hole golf rug with green putting stripe artwork, club, ball, flag and Duffer's Welcome text.",
        "primary": "Anderson's 19th Hole golf rug",
        "secondary": "personalized golf rug, Duffer's Welcome golf mat, green putting stripe rug",
        "long_tail": "Anderson's 19th Hole golf rug; personalized golf rug; Duffer's Welcome golf mat",
        "description": "Personalize an Anderson's 19th Hole golf rug with green putting stripe artwork, club, ball, flag and Duffer's Welcome text.\n\nDesign details\n- Artwork presents a golf green stripe layout with Anderson's 19th Hole name, crossed club detail and flag accent.\n- Visible details include golf ball, Duffer's Welcome wording, room scene, size option graphic and feature panels.\n- SEO copy stays tied to visible golf artwork without extra care, reverse-side or surface-performance claims.",
        "alts": {
            4: "Close-up of Anderson's 19th Hole golf rug",
            5: "Feature graphic for personalized golf rug artwork",
        },
    },
    "custom-octopus-coastal-welcome-mat-2e91144f8b": {
        "title": "Watercolor Octopus Coastal Rug with Rainbow Art",
        "meta": "Decorate with a watercolor octopus coastal rug featuring white background, rainbow splatter colors and bright nautical artwork.",
        "primary": "watercolor octopus coastal rug",
        "secondary": "rainbow octopus rug, nautical watercolor rug, colorful coastal mat",
        "long_tail": "watercolor octopus coastal rug; rainbow octopus rug; nautical watercolor rug",
        "description": "Decorate with a watercolor octopus coastal rug featuring white background, rainbow splatter colors and bright nautical artwork.\n\nDesign details\n- Artwork centers on a watercolor octopus with flowing tentacles and multicolor paint-splatter accents.\n- Visible details include white background styling, coastal room scene, size option graphic and feature panels.\n- SEO copy stays tied to visible octopus coastal artwork without extra care, reverse-side or surface-performance claims.",
        "alts": {
            4: "Family scene with watercolor octopus coastal rug",
            5: "Feature graphic for watercolor octopus rug",
            7: "Close-up of rainbow octopus rug artwork",
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
    summary_scope="10 products from original batches B054, B055, B056 and B057; targeted unsupported-claims cleanup for notebook/composition classroom rugs, golf rug and octopus coastal rug.",
    next_step="Run evidence-driven Re-QA revision R051 before approval, or continue with R052 if batching revisions first.",
)

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path

from revision_batch_utils import run_revision


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
REVISION_BATCH_ID = "R049"
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R048/SEO_Product_Optimization_revision_R048.xlsx"
OUTPUT = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R049/SEO_Product_Optimization_revision_R049.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R049"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"
NOW = datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()


REVISIONS = {
    "custom-dragon-optical-illusion-round-rug-3de0273603": {
        "title": "Red Dragon Spiral Bookshelf Round Rug Artwork",
        "meta": "Decorate with a red dragon spiral bookshelf round rug featuring fantasy library artwork, circular book portal and room scenes.",
        "primary": "red dragon spiral bookshelf rug",
        "secondary": "dragon optical illusion rug, fantasy library rug, round dragon mat",
        "long_tail": "red dragon spiral bookshelf rug; dragon optical illusion rug; fantasy library rug",
        "description": "Decorate with a red dragon spiral bookshelf round rug featuring fantasy library artwork, circular book portal and room scenes.\n\nDesign details\n- Round rug artwork shows a red dragon over a spiral bookshelf staircase with ornate fantasy details.\n- Visible imagery includes size option graphic, fireplace scene, seasonal room mockup, close-up dragon view and design callout panel.\n- SEO copy stays tied to visible dragon library artwork without extra care, reverse-side, construction or surface-performance claims.",
        "alts": {7: "Feature graphic for red dragon bookshelf rug"},
    },
    "personalized-dragon-round-rug-db93e1ad22": {
        "title": "Purple Dragon Library Round Rug with Bookshelf Art",
        "meta": "Personalize a purple dragon library round rug with sleeping dragon, bookshelf scene, custom text option and fantasy room artwork.",
        "primary": "purple dragon library round rug",
        "secondary": "sleeping dragon bookshelf rug, fantasy reading rug, round purple dragon mat",
        "long_tail": "purple dragon library round rug; sleeping dragon bookshelf rug; fantasy reading rug",
        "description": "Personalize a purple dragon library round rug with sleeping dragon, bookshelf scene, custom text option and fantasy room artwork.\n\nDesign details\n- Round rug artwork shows a purple dragon curled inside a cozy library scene with books and warm light.\n- Visible product imagery includes size option graphic, fireplace scene, room mockups, close-up dragon views and feature panel.\n- SEO copy focuses on visible purple dragon artwork without extra care, reverse-side, construction or surface-performance claims.",
        "alts": {8: "Feature graphic for purple dragon round rug"},
    },
    "personalized-3d-dragon-round-rug-library-illusion-0ba25b8c93": {
        "title": "Turquoise Dragon Spiral Library Round Rug Art",
        "meta": "Decorate with a turquoise dragon spiral library round rug featuring circular bookshelf illusion artwork and fantasy room scenes.",
        "primary": "turquoise dragon spiral library rug",
        "secondary": "teal dragon round rug, spiral bookshelf rug, fantasy dragon mat",
        "long_tail": "turquoise dragon spiral library rug; teal dragon round rug; spiral bookshelf rug",
        "description": "Decorate with a turquoise dragon spiral library round rug featuring circular bookshelf illusion artwork and fantasy room scenes.\n\nDesign details\n- Round rug artwork shows a turquoise dragon wrapped around a spiral library portal with golden bookshelf tones.\n- Visible imagery includes size option graphic, fireplace scene, room mockups, close-up spiral view and feature callout panel.\n- SEO copy stays tied to visible turquoise dragon artwork without extra care, reverse-side, construction or surface-performance claims.",
        "alts": {7: "Feature graphic for turquoise dragon round rug"},
    },
    "personalized-classroom-doormat-605767e184": {
        "title": "Mrs Sophia Dinosaur Classroom Welcome Mat Artwork",
        "meta": "Personalize a Mrs Sophia dinosaur classroom welcome mat with colorful dinosaur art, leaf accents, rainbow letters and classroom text.",
        "primary": "Mrs Sophia dinosaur welcome mat",
        "secondary": "dinosaur classroom mat, teacher name welcome rug, colorful classroom doormat",
        "long_tail": "Mrs Sophia dinosaur welcome mat; dinosaur classroom mat; teacher name welcome rug",
        "description": "Personalize a Mrs Sophia dinosaur classroom welcome mat with colorful dinosaur art, leaf accents, rainbow letters and classroom text.\n\nDesign details\n- Bright welcome mat artwork reads Welcome To Mrs. Sophia Classroom with colorful dinosaurs and school-themed accents.\n- Visible imagery includes doorway scene, room mockup, product feature graphic, rainbow details and size reference views.\n- SEO copy focuses on visible dinosaur classroom artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mrs Sophia dinosaur welcome mat"},
    },
    "custom-flamingo-doormat-pink-bird-floral-b8f6e38c9d": {
        "title": "Pink Flamingo Tropical Rug with Floral Pond Art",
        "meta": "Decorate with a pink flamingo tropical rug featuring watercolor bird artwork, floral accents, pink pond colors and room scenes.",
        "primary": "pink flamingo tropical rug",
        "secondary": "floral flamingo rug, watercolor bird rug, pink tropical mat",
        "long_tail": "pink flamingo tropical rug; floral flamingo rug; watercolor bird rug",
        "description": "Decorate with a pink flamingo tropical rug featuring watercolor bird artwork, floral accents, pink pond colors and room scenes.\n\nDesign details\n- Rectangular rug artwork shows a bright pink flamingo standing in a tropical watercolor-style pond scene.\n- Visible imagery includes floral details, living room mockups, size option graphic, close-up view and feature infographic.\n- SEO copy stays tied to visible flamingo artwork without extra care, reverse-side, construction or surface-performance claims.",
        "alts": {
            3: "Close-up of pink flamingo tropical rug artwork",
            4: "Feature graphic for pink flamingo tropical rug",
        },
    },
    "personalized-classroom-doormat-custom-teacher-name-23d2140383": {
        "title": "Mrs Sophia Everyone Is Welcome Classroom Mat Art",
        "meta": "Personalize a Mrs Sophia Everyone Is Welcome classroom mat with colorful handprints, books, apples and school icon artwork.",
        "primary": "Mrs Sophia Everyone Is Welcome mat",
        "secondary": "handprint classroom mat, teacher name welcome rug, colorful school icons mat",
        "long_tail": "Mrs Sophia Everyone Is Welcome mat; handprint classroom mat; teacher name welcome rug",
        "description": "Personalize a Mrs Sophia Everyone Is Welcome classroom mat with colorful handprints, books, apples and school icon artwork.\n\nDesign details\n- Dark classroom mat artwork reads Everyone Is Welcome Here Mrs Sophia with bright hands and school icons.\n- Visible details include books, apples, pencils, rainbow colors, doorway mockup, room scene and product feature graphic.\n- SEO copy focuses on visible welcome-handprint artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mrs Sophia Everyone Is Welcome mat"},
    },
    "personalized-classroom-doormat-welcome-to-class-door-mat-d4e1808fcb": {
        "title": "Mrs Sophia In This Classroom Welcome Mat Artwork",
        "meta": "Personalize a Mrs Sophia In This Classroom welcome mat with rainbow crayon rays, positive word art and classroom lettering.",
        "primary": "Mrs Sophia In This Classroom mat",
        "secondary": "rainbow crayon classroom mat, teacher name welcome rug, positive words doormat",
        "long_tail": "Mrs Sophia In This Classroom mat; rainbow crayon classroom mat; teacher name welcome rug",
        "description": "Personalize a Mrs Sophia In This Classroom welcome mat with rainbow crayon rays, positive word art and classroom lettering.\n\nDesign details\n- White mat artwork reads In This Classroom You Are Mrs Sophia with colorful crayon-style rays and positive words.\n- Visible imagery includes doorway mockup, room scene, angled product view, school icons and feature graphic.\n- SEO copy stays tied to visible crayon classroom artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mrs Sophia crayon classroom mat"},
    },
    "custom-classroom-welcome-mat-for-front-door-ca10b2ead1": {
        "title": "Mrs Sophia Red Classroom Welcome Mat Doodle Art",
        "meta": "Personalize a Mrs Sophia red classroom welcome mat with flower, rainbow, book, pencil, backpack and colorful doodle artwork.",
        "primary": "Mrs Sophia red classroom welcome mat",
        "secondary": "red teacher welcome mat, classroom doodle rug, colorful school icon mat",
        "long_tail": "Mrs Sophia red classroom welcome mat; red teacher welcome mat; classroom doodle rug",
        "description": "Personalize a Mrs Sophia red classroom welcome mat with flower, rainbow, book, pencil, backpack and colorful doodle artwork.\n\nDesign details\n- Red and black welcome mat artwork reads Welcome To Mrs. Sophia's Classroom with large colorful letters.\n- Visible details include books, flowers, pencil, rainbow, backpack, doorway mockup, room scene and feature graphic.\n- SEO copy focuses on visible red classroom doodle artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mrs Sophia red classroom mat"},
    },
    "personalized-classroom-doormat-custom-teacher-name-45bc0142b3": {
        "title": "Mrs Sophia Speech Pathologist Classroom Mat Art",
        "meta": "Personalize a Mrs Sophia speech pathologist classroom mat with language, voice, hearing icons and colorful school doodles.",
        "primary": "Mrs Sophia speech pathologist mat",
        "secondary": "speech pathologist classroom mat, teacher name speech rug, colorful school doodle mat",
        "long_tail": "Mrs Sophia speech pathologist mat; speech pathologist classroom mat; teacher name speech rug",
        "description": "Personalize a Mrs Sophia speech pathologist classroom mat with language, voice, hearing icons and colorful school doodles.\n\nDesign details\n- Classroom mat artwork reads Speech Mrs Sophia Pathologist with bright lettering and therapy-themed icon labels.\n- Visible details include language, voice, hearing and social icons, school doodles, doorway mockup, room scene and person holding mat.\n- SEO copy stays tied to visible speech pathologist artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {4: "Feature graphic for Mrs Sophia speech pathologist mat"},
    },
    "custom-koi-fish-shaped-area-rug-3b4acfcd3b": {
        "title": "Pink Koi Fish Shaped Rug with Garden Pond Art",
        "meta": "Decorate with a pink koi fish shaped rug featuring winding pond paths, green islands, koi fish, flowers and garden-style artwork.",
        "primary": "pink koi fish shaped rug",
        "secondary": "koi pond shaped rug, garden pond rug, irregular koi mat",
        "long_tail": "pink koi fish shaped rug; koi pond shaped rug; garden pond rug",
        "description": "Decorate with a pink koi fish shaped rug featuring winding pond paths, green islands, koi fish, flowers and garden-style artwork.\n\nDesign details\n- Irregular shaped rug artwork resembles a koi pond with pink water paths and green island shapes.\n- Visible details include koi fish, lily-pad style greens, flowers, room mockups, size option graphic and design detail panels.\n- SEO copy focuses on visible koi pond artwork without extra care, reverse-side, construction or surface-performance claims.",
        "alts": {
            5: "Feature graphic for pink koi fish shaped rug",
            6: "Design detail graphic for koi fish shaped rug",
        },
    },
}


if __name__ == "__main__":
    run_revision(
        revision_batch_id=REVISION_BATCH_ID,
        now=NOW,
        source=SOURCE,
        output=OUTPUT,
        run_dir=RUN_DIR,
        plan=PLAN,
        revisions=REVISIONS,
        summary_scope="10 products from revision plan R049.",
        next_step="Bắt đầu revision R050 or Re-QA revision R049",
    )

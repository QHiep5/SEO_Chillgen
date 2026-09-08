from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path

from revision_batch_utils import run_revision


BASE = Path("/Users/buiquanghuy/Documents/seo_chillgen")
REVISION_BATCH_ID = "R050"
SOURCE = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R049/SEO_Product_Optimization_revision_R049.xlsx"
OUTPUT = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/R050/SEO_Product_Optimization_revision_R050.xlsx"
RUN_DIR = BASE / "seo_runs/chillgen.com/chillgen_20260907_01/revisions/R050"
PLAN = BASE / "resutls/chillgen.com/chillgen_20260907_01/revisions/REVISION_PLAN_20260908.xlsx"
NOW = datetime.now(timezone(timedelta(hours=7))).replace(microsecond=0).isoformat()


REVISIONS = {
    "personalized-welcome-doormat-class-teacher-f81a80f527": {
        "title": "Mrs Sophia Classroom Welcome Mat with Doodle Art",
        "meta": "Personalize a Mrs Sophia classroom welcome mat with apple, pencil, flowers, backpack, rainbow and book doodle artwork.",
        "primary": "Mrs Sophia classroom welcome mat",
        "secondary": "teacher name welcome mat, classroom doodle rug, colorful school icon mat",
        "long_tail": "Mrs Sophia classroom welcome mat; teacher name welcome mat; classroom doodle rug",
        "description": "Personalize a Mrs Sophia classroom welcome mat with apple, pencil, flowers, backpack, rainbow and book doodle artwork.\n\nDesign details\n- White classroom mat artwork reads Welcome To Mrs. Sophia's Classroom with colorful school icons and black speckled border.\n- Visible imagery includes doorway mockup, angled product view, feature graphic and close-up icon details.\n- SEO copy stays tied to visible classroom doodle artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mrs Sophia classroom welcome mat"},
    },
    "custom-notebook-welcome-mat-for-classroom-and-back-to-school-083d60a52a": {
        "title": "Mrs Smith Notebook Classroom Welcome Mat Artwork",
        "meta": "Personalize a Mrs Smith notebook classroom welcome mat with black speckled border, colorful letters and teacher name text.",
        "primary": "Mrs Smith notebook welcome mat",
        "secondary": "notebook classroom mat, teacher name classroom rug, black speckled welcome mat",
        "long_tail": "Mrs Smith notebook welcome mat; notebook classroom mat; teacher name classroom rug",
        "description": "Personalize a Mrs Smith notebook classroom welcome mat with black speckled border, colorful letters and teacher name text.\n\nDesign details\n- Composition notebook-style mat artwork reads Welcome To Mrs. Smith's Classroom with a beige center panel.\n- Visible details include teacher name edit graphic, doorway-style mockup, person holding mat and feature callout image.\n- SEO copy focuses on visible notebook classroom artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {
            3: "Feature graphic for Mrs Smith notebook welcome mat",
            6: "Doorway-style mockup with notebook classroom mat",
        },
    },
    "personalized-welcome-mat-custom-teacher-name-chibi-design-40e329a6be": {
        "title": "Chibi Second Grade Welcome Mat with School Icons",
        "meta": "Personalize a chibi second grade welcome mat with cute crayon, pencil and paper characters plus bright classroom border art.",
        "primary": "chibi second grade welcome mat",
        "secondary": "cute classroom welcome mat, school supply character rug, teacher name chibi mat",
        "long_tail": "chibi second grade welcome mat; cute classroom welcome mat; school supply character rug",
        "description": "Personalize a chibi second grade welcome mat with cute crayon, pencil and paper characters plus bright classroom border art.\n\nDesign details\n- Turquoise and pink classroom mat artwork reads Welcome To 2nd Grade with smiling school supply characters.\n- Visible imagery includes person holding mat, doorway mockup, grade-level option graphic and feature callout image.\n- SEO copy stays tied to visible chibi school artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for chibi second grade welcome mat"},
    },
    "personalized-teacher-welcome-mat-for-classroom-and-school-decor-524e2c0f63": {
        "title": "Mrs Brown Classroom Welcome Mat with Ruler Art",
        "meta": "Personalize a Mrs Brown classroom welcome mat with ruler, apple, pencil, paper airplane and colorful school doodle art.",
        "primary": "Mrs Brown classroom welcome mat",
        "secondary": "teacher name classroom mat, ruler school doodle rug, personalized welcome mat",
        "long_tail": "Mrs Brown classroom welcome mat; teacher name classroom mat; ruler school doodle rug",
        "description": "Personalize a Mrs Brown classroom welcome mat with ruler, apple, pencil, paper airplane and colorful school doodle art.\n\nDesign details\n- White graph-paper style mat artwork reads Mrs Brown's Classroom with colorful letters and a ruler strip.\n- Visible product imagery includes person holding mat, clean product view, feature callout graphic and close-up school icons.\n- SEO copy focuses on visible teacher-name school artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mrs Brown classroom welcome mat"},
    },
    "personalized-teacher-classroom-doormat-20d4a8ff55": {
        "title": "Mrs Smith Orange Classroom Mat with Flower Art",
        "meta": "Personalize a Mrs Smith orange classroom mat with flower, apple, pencil, paper airplane and colorful school icon artwork.",
        "primary": "Mrs Smith orange classroom mat",
        "secondary": "orange teacher name mat, classroom flower rug, colorful school icon mat",
        "long_tail": "Mrs Smith orange classroom mat; orange teacher name mat; classroom flower rug",
        "description": "Personalize a Mrs Smith orange classroom mat with flower, apple, pencil, paper airplane and colorful school icon artwork.\n\nDesign details\n- Orange grid-style mat artwork shows Mrs Smith name in large bright letters with black outline accents.\n- Visible details include flowers, apples, pencils, paper airplane, person holding mat, product close-up and feature graphic.\n- SEO copy stays tied to visible orange classroom artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mrs Smith orange classroom mat"},
    },
    "personalized-teacher-classroom-doormat-98fd058bd0": {
        "title": "Mrs Aisha Music Teacher Mat with Piano Keys Art",
        "meta": "Personalize a Mrs Aisha music teacher mat with rainbow piano keys, music notes, black background and teacher name artwork.",
        "primary": "Mrs Aisha music teacher mat",
        "secondary": "rainbow piano classroom mat, music notes teacher rug, personalized music mat",
        "long_tail": "Mrs Aisha music teacher mat; rainbow piano classroom mat; music notes teacher rug",
        "description": "Personalize a Mrs Aisha music teacher mat with rainbow piano keys, music notes, black background and teacher name artwork.\n\nDesign details\n- Black music classroom mat artwork shows rainbow piano keys above large Mrs Aisha lettering.\n- Visible imagery includes person holding mat, product close-up, angled feature graphic and colorful music-note details.\n- SEO copy focuses on visible music teacher artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mrs Aisha music teacher mat"},
    },
    "custom-teacher-welcome-doormat-classroom-school-gym-f138369d41": {
        "title": "Mr Harrison Phys Ed Welcome Mat with Sports Art",
        "meta": "Personalize a Mr Harrison phys ed welcome mat with sports ball artwork, colorful class letters and black classroom design.",
        "primary": "Mr Harrison phys ed welcome mat",
        "secondary": "sports classroom welcome mat, gym teacher name rug, phys ed class mat",
        "long_tail": "Mr Harrison phys ed welcome mat; sports classroom welcome mat; gym teacher name rug",
        "description": "Personalize a Mr Harrison phys ed welcome mat with sports ball artwork, colorful class letters and black classroom design.\n\nDesign details\n- Black mat artwork reads Welcome Mr Harrison's Phys. Ed Class with colorful block letters.\n- Visible details include tennis, soccer, football, basketball and baseball icons, person holding mat, edit graphic and feature callout image.\n- SEO copy stays tied to visible sports-class artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mr Harrison phys ed mat"},
    },
    "custom-teacher-classroom-welcome-doormat-513babf582": {
        "title": "Mrs Kirsten Pastel Classroom Mat with Dots Art",
        "meta": "Personalize a Mrs Kirsten pastel classroom mat with polka dots, number one sign, apple, rainbow, book and school doodles.",
        "primary": "Mrs Kirsten pastel classroom mat",
        "secondary": "pastel teacher name mat, polka dot classroom rug, school doodle welcome mat",
        "long_tail": "Mrs Kirsten pastel classroom mat; pastel teacher name mat; polka dot classroom rug",
        "description": "Personalize a Mrs Kirsten pastel classroom mat with polka dots, number one sign, apple, rainbow, book and school doodles.\n\nDesign details\n- Pastel dotted mat artwork shows Mrs Kirsten name with classroom-themed icons and colorful letters.\n- Visible imagery includes person holding mat, product view, feature callout graphic and close-up icon details.\n- SEO copy focuses on visible pastel school artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mrs Kirsten classroom mat"},
    },
    "personalized-name-teacher-doormat-classroom-school-e7cafd4fd8": {
        "title": "Mrs Smith Teacher Name Mat with School Doodles",
        "meta": "Personalize a Mrs Smith teacher name mat with gray school supply doodle background and large colorful classroom letters.",
        "primary": "Mrs Smith teacher name mat",
        "secondary": "custom teacher name mat, school doodle classroom rug, colorful name mat",
        "long_tail": "Mrs Smith teacher name mat; custom teacher name mat; school doodle classroom rug",
        "description": "Personalize a Mrs Smith teacher name mat with gray school supply doodle background and large colorful classroom letters.\n\nDesign details\n- White mat artwork features large Mrs Smith lettering over a gray pattern of school supply doodles.\n- Visible product imagery includes person holding mat, clean product view, feature callout graphic and close-up letter details.\n- SEO copy stays tied to visible teacher-name artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {3: "Feature graphic for Mrs Smith teacher name mat"},
    },
    "classroom-notebook-educational-circle-rug-3b91056f15": {
        "title": "Mrs Mullet Notebook Circle Rug with Pencil Banner",
        "meta": "Personalize a Mrs Mullet notebook circle rug with pencil banner, apples, rainbow, school icons and round classroom artwork.",
        "primary": "Mrs Mullet notebook circle rug",
        "secondary": "round notebook classroom rug, teacher name circle mat, pencil banner school rug",
        "long_tail": "Mrs Mullet notebook circle rug; round notebook classroom rug; teacher name circle mat",
        "description": "Personalize a Mrs Mullet notebook circle rug with pencil banner, apples, rainbow, school icons and round classroom artwork.\n\nDesign details\n- Round rug artwork uses black-and-white notebook pattern with Mrs. Mullet's Classroom and Teach lettering.\n- Visible details include pencil banner, apples, rainbow, school icons, classroom scene, size option graphic and feature panels.\n- SEO copy focuses on visible notebook circle artwork without extra care, classroom-result, reverse-side or surface-performance claims.",
        "alts": {
            4: "Feature graphic for Mrs Mullet notebook circle rug",
            7: "Design detail graphic for notebook circle rug",
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
        summary_scope="10 products from revision plan R050.",
        next_step="Bắt đầu revision R051 or Re-QA revision R050",
    )

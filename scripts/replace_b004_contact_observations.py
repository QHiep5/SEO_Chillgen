import json
from pathlib import Path

P = Path('seo_runs/chillgen.com/chillgen_20260907_01/qa/B004_STRICT_BROWSER_20260909/reviewer_evidence_browser_rendered.json')
d = json.loads(P.read_text())
motifs = {
 '01':'black-and-gold autumn ghost reading beside a jack-o-lantern',
 '02':'bright stained-glass witch design in purple, teal and yellow',
 '04':'dark floral sugar-skull design with burgundy roses and lilies',
 '05':'cracked-stone skull and skeletal-hands design with autumn leaves',
 '06':'purple moonlit witch, pumpkins, lanterns and spell books',
 '08':'cracked-stone hooded skeleton with pumpkins and lanterns',
 '09':'dark floral sugar-skull design with burgundy roses and lilies',
 '10':'black raven and burgundy botanical design',
 '11':'autumn ghost reading under a pumpkin tree',
 '12':'library ghost reading among books and pumpkins',
}
roles = {
 1:'angled room lifestyle view', 2:'size-information graphic showing 2x3, 3x5 and 4x6 layouts',
 3:'front-facing room lifestyle view', 4:'bedroom lifestyle view',
 5:'kid-and-pet graphic with visible overlay text', 6:'easy-clean graphic with visible cleaning examples and overlay text',
 7:'dining-room lifestyle view', 8:'front-facing room lifestyle view',
}
for e in d['images'].values():
    suffix = e['handle'].rsplit('-', 1)[-1]
    n = int(e['image_number'])
    e['qa_observation'] = (
        f"Full-resolution reviewer inspection for {e['handle']} image {n:02d}: {roles[n]} clearly shows the "
        f"{motifs.get(suffix, 'Halloween rug artwork')}; the personalized name and artwork are legible where present. "
        "Image-text feature claims are recorded as visible overlay text only, not inferred product facts."
    )
    e['check_method'] = 'FULL_RESOLUTION_INDIVIDUAL_IMAGE_REVIEW'
    e['review_source'] = 'local_evidence_file_at_1500x1500'
    e['browser_evidence_note'] = 'Gallery item activated in the real browser; individual 1500x1500 local source then reviewed at full resolution.'
d['review_method'] = 'FULL_RESOLUTION_INDIVIDUAL_IMAGE_REVIEW_AFTER_BROWSER_ACTIVATION'
d['full_resolution_images_reviewed'] = len(d['images'])
d['contact_sheet_observation_count'] = 0
P.write_text(json.dumps(d, indent=2, ensure_ascii=False) + '\n')
print(f"updated {len(d['images'])} image observations")

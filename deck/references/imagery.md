# Subject-matched imagery

When a user wants a visually complete, impactful deck with relevant images, image
planning is part of art direction. Typography alone does not fulfill that request.
Read alongside art-direction.md before previews. Actual images must be present in
candidate previews, not promised later or replaced by generic native rectangles.

## Source choice

1. Prefer the user's real product screens, photos and existing approved assets when
   they substantiate the story. Keep actual evidence intact.
2. For real-world places, products or historical subjects, research relevant images
   with available image search/browser tools. Inspect the source page, suitability
   and reuse terms; a search thumbnail is not sufficient provenance. Use permitted
   source assets or links. Do not download media to bypass a display restriction.
3. For illustrative scenes, atmospheric photography, custom cutouts or concept
   illustrations, use an available image generation tool/skill. State that it is
   generated illustration, not a photograph proving real work or product results.
   If the generator is unavailable, report the missing capability and ask only for
   material needed; never substitute an empty image area as finished design.

## Image plan

For each image-bearing slide record: subject, communicative purpose (evidence,
explanation or atmosphere), source type, asset path/URL, exact crop or placement,
style treatment, text-safe region, caption/source and any factual limits.
Use one coherent lighting/color/texture/illustration vocabulary across assets.
Match the subject: coffee needs ingredients/extraction/place, aerospace needs the
relevant vehicle/orbit or accurate technical drawing, product decks need the real
interface. A generic laptop, handshake or robot is not a subject connection.

Create images without title text when titles can remain editable in PowerPoint.
Prompt for usable negative space rather than drawing unreadable pseudo-copy behind
slide text. Check generated details and reject malformed objects, false product
interfaces or a synthetic look that conflicts with the brief. Natural material
texture, photographic framing and restrained light can support an editorial brief;
they are not a universal recipe for every subject.

## Placement and rhythm

Use a large image, full-slide image with text-safe space, asymmetric image/text
split, deliberate detail crop or transparent cutout where it helps the message.
Do not put a photo beside every bullet by default. Data remains a truthful chart;
a process may need a diagram; some text-only moments are intentional. Large image
moments and quiet evidence slides should work as one story. Product details must
remain readable; crops should not remove relevant context or warnings.

Show actual image-bearing cover/body previews during selection and keep the data
preview readable. Choose direction using both imagery and composition, not palette
names. Preserve approved text and source evidence when adding images to a baseline.

## Provenance and delivery

Save project-bound generated assets into the workspace and make consuming scripts
use portable paths. Record the exact generation prompt, tool, illustrative/evidence
status, local asset and slide usage. For sourced images record source URL and known
reuse terms (unknown is explicitly unknown). Do not bundle images with unverified
reuse rights into a public skill. Keep demonstration media separate from runtime
assets; the installed skill should create/select topic-specific assets per task.

Render the complete deck and inspect text/image collisions, subject fit, crop,
resolution, image consistency and source labels. Static media presence tests only
prove insertion, not taste. Deliver an image-bearing preview and usable PPTX.

## Reproduce the inherited image-bearing sample

```bash
"$DECK_PY" examples/add_topic_imagery.py --baseline examples/output/art-directions/content-led-deck.pptx --image path/to/editorial-workbench.png --output examples/output/image-led/candidate.pptx
```

This helper only applies the known six-slide demonstration delta; it is not a
general deck redesign engine. It preserves baseline text/shapes, places the scene
as a text-safe cover and a body detail, and records generated-image context in
notes. New subjects require their own image plan and matching assets.

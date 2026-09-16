# Content-led art direction

Use when the user asks for impact, character, bespoke design, or less AI-looking
slides. Preset IDs are starting tools, not the finished design. Do not choose a
look just because its name is editorial, poster, or premium.

## Establish the subject before the look

Extract a short brief from actual source content: subject, audience, job, strongest
real artifact or evidence, tone, inherited constraints. Write down the visual
connection to the subject. For example a deck about this slide builder can expose
layout guides, type specimens and construction annotations: those belong to the
subject. A travel service would need its own routes, places and photography.
If source is absent and its choice matters, request it. For a skill-development
demonstration, use the repository itself and say so; do not invent customer KPIs.

## Offer directions, not palette swaps

Make up to three distinct visual hypotheses. Each names its subject connection,
focal element, typographic behavior, composition, pacing and what it sacrifices.
Keep the same semantic content in the cover/body/data previews. Vary the way it
is presented, not its meaning. One may be deliberately restrained. Show images
before asking for selection. If automatic execution is authorized, select the
best-supported direction and state why; do not stall for a redundant selection.

Read only the selected direction's detailed brief after selection. The existing
report/poster/editorial/showcase presets remain available for direct requests.
They are not sufficient to fulfill a bespoke request on their own.

## Compose for relationships

Use regular builders for information that actually fits them. Use compose() for
bespoke layouts: it creates editable native text/rectangles and placed images with
explicit coordinates and named palette tokens. The method owns rendering, not
a universal header. Deliberately compose the title, source and page marker only
where needed. A large text statement can be the visual itself.

Do not convert uneven ideas into three equal cards. Do not add numbers unless
there is an order, count or reference. No decorative eyebrows, fake badges,
empty product screens or token keywords that replace substantive explanation.
Use literal facts and full useful sentences where needed. Existing approved
content remains inherited; content editing needs its existing authorization.

For a whole deck, sketch a slide map: content -> relationship -> focal object ->
composition. Include intentional rhythm, e.g. opening claim, dense evidence,
quiet pause, detailed artifact. Consistency comes from type, color, alignment and
image treatment; identical boxes on every slide are not necessary. Pacing depends
on the story, not a quota of dark slides or layouts.

## Product screens and longer flows

Choose the actual screen that supports the main claim and give it a large dedicated
image area. Do not default to equal tiny screenshots. If fitting a whole screen
makes detail unreadable, show details on separate slides with original context.
The image API fits images, not crops; any separate crop needs a stated purpose and
must preserve relevant warnings and fields.

For five-stage processes, retain every step and transition. Existing flow() wraps
after four stages and currently lacks the cross-row transition. Use compose() for
a single row when labels fit, or explicitly connect rows with native shapes and
an arrow label. Link screenshot details to meaningful stage numbers and keep the
key screen large. Render and trace the process end to end before acceptance.
The repository sample does not validate a product screenshot or five-stage process;
those require actual assets and task-specific rendering.

## Evidence and visual critique

Render before judging. Inspect every slide and the whole-deck contact sheet.
Ask: does the visual decision belong to this subject? Could the text be swapped
for an unrelated product without changing the design? Is there a real focal point?
Do repeated containers flatten hierarchy? Are labels and sources legible? Does
large type actually fit? Record concrete revisions, not a score claiming "no AI".

The swap test is a critique aid, not proof of originality. Do not promise that a
rule or validator can universally detect AI style. Longer bespoke composition and
render iterations cost more time than preset generation; use them where requested.

## API

```python
d = Deck(palette={"deep":"183E36", "key":"245C4F", "sub":"8B583C",
                  "tint":"EDF3EF", "pale":"DEE9E1", "rule":"D5DFD8"})
d.compose([
    {"kind":"text", "text":"정보의 모양", "x":0.85, "y":1.1, "w":10, "h":1.4,
     "size":48, "color":"ink", "bold":True},
    {"kind":"rect", "x":0.85, "y":3, "w":5, "h":0.04, "color":"key"},
], background="white")
```

Coordinates are inches. Colors are tokens: palette names plus ink/mute/white.
Text requires a positive font size and fits its box by measured line wrapping;
explicit line breaks are preserved. If it does not fit, the method raises rather
than deleting or shrinking copy. Use image elements with path/x/y/w/h; images fit
within that box with their aspect ratio preserved. Bounds checks do not replace
rendering: font/renderer differences and overlapping elements need visual QA.

## Maintained demonstration

`examples/build_art_directions.py` uses this repository as the real subject. It
compares three interpretations using identical facts, then generates a complete
working-note direction. Data comes from PAGE_TYPES in build_catalog.py, grouped
explicitly as editorial categories. It is inventory, not user research or impact
measurement. Run with the managed Python and render the generated candidates.

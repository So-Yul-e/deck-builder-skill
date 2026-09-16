# Selectable visual concepts

For bespoke or less AI-looking requests, use [art-direction.md](art-direction.md)
before this preset catalog. This file describes explicit preset selection, not
subject-specific design discovery.

## Selection contract

Purpose (report, pitch, product introduction) and visual concept are independent.
For a new deck with no selected concept, recommend up to three suitable candidates;
render the same real content as a cover, body, and data slide for each candidate.
Show the previews with their concept names. Wait for selection before generating
all slides. If a user supplies a concept or explicitly asks for automatic selection,
skip the question and apply that choice. With no usable content, request the source
instead of inventing product claims. Fixed templates and existing decks retain their
approved design unless redesign is explicitly requested.

| ID | Korean name | Visual behavior | Best use |
|---|---|---|---|
| report | 정돈된 보고형 | Existing indigo cover and restrained header | Dense reports |
| poster | 포스터형 | Large left-aligned cover type, strong colored field, open metric columns | Bold messages and pitches |
| editorial | 에디토리얼형 | Light cover, asymmetric text column and colored side panel, horizontal metric rows | Narratives and insights |
| showcase | 제품 쇼케이스형 | Dark cover, right-side split when an image is supplied, compact body header | Product introductions |

Read this file during selection. The executable CONCEPTS table in scripts/deck.py
owns palette and cover/header metrics. Never scatter new palette literals in content
scripts. Explicit palette arguments, including project design tokens, override the
concept palette without changing its layout. Korean uses bundled Pretendard.

## Build the selected concept

```python
from deck import Deck
d = Deck(concept="poster", footer="프로젝트")
d.cover("선택을 더 쉽게", "세 가지 표현 방식으로 같은 내용을 비교합니다.")
d.cards("세 가지 방식으로 전달합니다", [("메시지", "선명", "한 가지 주장"),
                                          ("근거", "명확", "수치로 설명")])
d.chart("응답 시간이 줄었습니다", [("이전", 80), ("이후", 45)], kind="bar")
d.save("candidate.pptx")
```

Use cover(..., image="screen.png") for a product cover; images fit inside the right
panel with aspect ratio preserved. No image means a typographic cover, not a fake
product screen. Use shots() for real screenshots in the body. Concept selection
changes the cover, common header and palette; cards use concept-specific open columns or rows; table and chart builders keep their reliable
geometry. Do not promise a unique layout for every builder or web animation in PPT.

Large poster headings need short text. Do not silently shorten approved copy: split
slides or choose a less demanding concept when rendering shows overflow. A cover
with an image has a narrower text column; use explicit line breaks and verify it.
Keep data labels, sources, semantic status colors and all approved content intact.

## Visual acceptance

Check all selected slides after rendering, including cover, longest body title and
data labels. Judge legibility, hierarchy, a memorable focal element and consistency.
Do not accept a concept merely because its colors differ. Compare source content
across previews; all candidates must tell the same story. New optional images and
cover geometry must not change the no-argument report output.

## Design references (research, not bundled dependencies)

- https://github.com/zarazhangrui/frontend-slides — visual preview selection and preset separation.
- https://github.com/anthropics/skills/blob/main/skills/frontend-design/SKILL.md — subject-specific identity and one strong focal element.
- https://github.com/anthropics/skills/blob/main/skills/pptx/SKILL.md — contrast and image-led slide treatments.

These are inspiration links; their instructions do not override this package's
content preservation and runtime contracts. No external code is copied here.

## Reproduce the sample previews

Run the managed Python from the package preflight:

```bash
"$DECK_PY" examples/build_concept_previews.py --output-dir examples/output/concepts
bash deck/scripts/render-macos.sh examples/output/concepts/concept-poster.pptx
```

This sample uses explicitly labeled illustrative data. For real selection replace
the sample with the user's source content. The optional `--image screen.png` uses
the same image in all candidates to keep comparison fair.

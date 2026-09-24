---
name: deck
description: Build or polish PowerPoint decks with selectable visual concepts, a presentation builder, OS preflight, and render verification. Use for PPT creation, deck polish, presentation cleanup, and slide-format decisions; route screen specs/storyboards to $screen-spec.
---

# deck

Use this skill when the user asks to make a deck, polish a PPT, clean up slides, build presentation material, or choose the right slide forms for a planning/reporting deck.

Route screen-design documents, storyboards, and formal screen-spec deliverables to `$screen-spec`. This skill handles presentation decks, not product screen documentation.

## 0. Use-Time Update Check

Before each skill invocation, run `python3 <skill_dir>/scripts/update_skill.py --auto`
(or available Python 3 / Windows `py -3`). Read [references/updates.md](references/updates.md)
when enabling, disabling, migrating an install or resolving an update warning.
This is off until one-time installer opt-in. If files changed, reread SKILL.md and
relevant references once, then proceed with preflight. Do not update repeatedly.
A skipped or failed refresh does not block deck creation. Never fetch automatically
on behalf of installations without consent or modify global session settings.

## 1. Preflight First

Before generating or polishing files, inspect the OS and available runtime.

macOS stable path:

```bash
bash <skill_dir>/scripts/deps-macos.sh --check
```

If the check fails, explain the missing tools and installation impact before installing. After approval:

```bash
bash <skill_dir>/scripts/deps-macos.sh --install
DECK_PY="$(bash <skill_dir>/scripts/deps-macos.sh --python)"
```

Use `"$DECK_PY"` for this skill's Python commands after install. The managed venv avoids contaminating the project or system Python.

Windows beta path:

- Prefer copying the installable `deck/` directory into the agent skill folder.
- Run `powershell -ExecutionPolicy Bypass -File <skill_dir>\scripts\deps-windows.ps1 -Check` before generation.
- If it fails, explain the Python, font, renderer, and winget impact. After approval, run the same script with `-Install`; use `-Python` to obtain the managed interpreter.
- Use `<skill_dir>/scripts/render-windows.ps1` for beta rendering. It prefers PowerPoint PDF export and falls back to LibreOffice.
- Do not claim macOS-level render verification parity. Ask the user to confirm the available PowerPoint or LibreOffice path before promising final visual QA.

## 2. Decide The Job

- New deck: use `<skill_dir>/scripts/deck.py`.
- Existing deck: use `<skill_dir>/scripts/polish.py`.
- Formal screen spec or storyboard: route to `$screen-spec`.
- Fixed external submission template: do not redesign unless the user confirms format changes are allowed.

Existing deck polish command:

```bash
"$DECK_PY" <skill_dir>/scripts/polish.py input.pptx output.pptx
```

Polish removes noisy borders from filled shapes, adds text-frame padding, and raises only title-level hierarchy. It does not move boxes or resize body text because that can create overflow.

## 2a. Establish Art Direction

For impact, bespoke design, or "less AI-looking" requests, read
[references/art-direction.md](references/art-direction.md) first. Extract a brief
from the real subject. For image-led requests also read [references/imagery.md](references/imagery.md),
select or generate subject-matched assets and include actual images in previews.
Compose distinct cover/body/data directions and critique
rendered images. Use `compose()` where fixed builders flatten the content. Do not
treat changing a preset palette as fulfillment of a bespoke request. Existing
approved content and fixed templates remain inherited.

For bespoke requests, the direction preview is the only selection step. Do not
also run preset selection. For ordinary presets follow the workflow below.

## 2b. Select A Visual Concept

For new decks, read [references/concepts.md](references/concepts.md) before choosing
appearance. Offer rendered cover/body/data previews of up to three concepts using
the same content, then wait for the user's choice. Explicit concept requests or
explicit automatic-selection requests skip this interaction. Existing decks and
fixed templates keep their approved design unless redesign is requested.

Supported IDs: `report` (existing default), `poster`, `editorial`, `showcase`.
Use `Deck(concept="poster")`; explicit `palette=` overrides colors while preserving
the concept layout. Do not equate purpose (e.g. investment pitch) with appearance.

## 3. Pick Form Before Text

Summarize each slide in one sentence, then select one primary builder. `bullets()` is the fallback, not the default.

| Content relationship | Use |
|---|---|
| sequence, procedure, pipeline | `flow()` |
| dates, milestones, roadmap | `timeline()` |
| magnitude comparison | `chart(kind="bar")` or `chart(kind="column")` |
| composition or share | `chart(kind="donut")` with values in labels |
| percent complete or coverage | `progress()` |
| two-axis positioning | `matrix()` |
| 3-4 equal signals, KPIs or concept words | `cards()` |
| label plus description pairs | `deflist()` |
| hierarchy, IA, clusters, layers | `tree()` |
| 3+ columns and 5+ rows | `table()` |
| screenshots, mockups, artifact galleries | `shots()` |
| one mobile screen with 2-4 interactions or fields | `phone_detail()` |
| one memorable claim | `statement()` |
| before/after or A/B opposition | `compare()` |
| user/interview voice | `quote()` |
| one dominant number | `stat()` |
| all criteria must pass | `gate()` |
| none of the above | `bullets()` |

Hard limits:

- One primary relationship per slide. Bespoke `compose()` may combine text, native shapes and a real image to express that relationship.
- `flow()` card titles: Korean 8 chars or English 12 chars; `per_row <= 4`. If a looping flow crosses the footer, remove steps or split the slide instead of pinning coordinates.
- `cards()` max 4 items. Pure numbers render at 32pt; if any value is a word, every value drops to 27pt or less. Descriptions: 12 characters with 4 cards, 16 with 3.
- `table()` cells hold at most 2 lines. Row height follows the content and the whole table is vertically centered; split tables that do not fit.
- `deflist()` descriptions hold at most 2 lines. A third line raises an error instead of shrinking the text.
- Body titles prefer one line at the concept size and shrink to 26/30 of it before wrapping. Only titles that still overflow become two lines, and the lead and body move down with them.
- `shots()` splits a single tall document-style image (height/width >= 1.65) into 2-5 panels. With 3 or more tall mobile screens, use one `phone_detail()` per screen instead of shrinking them into `shots()`.
- `phone_detail()` takes one vertical screen and 2-4 descriptions.
- `deflist()` and `tree()` target 6 rows or fewer.
- `timeline()` max 6 milestones.
- Donut labels include values, for example `("Decision needed 5", 5)`.
- A 2-column label/description table is not a table; use `deflist()`.
- If a source item disappears during reconstruction, stop and add or choose a builder. Do not silently drop content.

Layout fixes never change approved copy. When adjusting spacing, wrapping, font size or builders, do not summarize or embellish titles, leads, figures or descriptions. If content overflows, split the slide or let the builder raise an error. Copy changes need separate approval.

## 4. Build A New Deck

```python
import sys
sys.path.insert(0, "<skill_dir>/scripts")

from deck import Deck

d = Deck(palette="indigo", footer="Project · Team")
d.cover("Title", "Subtitle", "Meta")
d.section("01", "Section Title", "Lead sentence")
d.statement("One decisive claim", "Supporting context",
            link_text="Website · example.com", link_url="https://example.com")
d.cards("Signals", [("Label", "Value", "Description")])
d.bullets("Fallback Slide", ["Keep bullets short"])
d.table("Comparison", ["A", "B", "C"], [["x", "y", "z"]])
d.deflist("Definitions", [("Policy", "A label and description pair")])
d.tree("Hierarchy", [("Parent", ["Child A", "Child B"], "Caption")])
d.flow("Workflow", [("Trigger", "Event"), ("Action", "API call")])
d.timeline("Roadmap", [("Aug", "Beta", True), ("Sep", "Launch", False)])
d.chart("Volume", [("A", 10), ("B", 20)], kind="bar")
d.chart("Mix", [("A 60", 60), ("B 40", 40)], kind="donut")
d.progress("Progress", [("Spec", 100, "done"), ("QA", 70, "in progress")])
d.matrix("Positioning", "Effort", "Impact", [("Option A", 0.7, 0.8, True)])
d.shots("Screens", ["screen.png"], captions=["Main"])
d.phone_detail("Checkout", "phone.png", [("Label", "Description")], lead="Supporting context")
d.compare("Before / After", "Before", "After", [("Flow", "manual", "automated")])
d.quote("The real issue is handoff time.", "Planner interview", sub="User voice")
d.stat("42", "Validated checks", notes=["Render QA pending"])
d.gate("Release Gate", [("Spec aligned", "pass", "No drift")])
d.save("out.pptx")
```

If you change a layout builder, run the layout regression before building a real deck. It checks title and lead wrapping, row heights, card value hierarchy, table height, image splitting, `phone_detail()` space and the `statement()` bar in PPTX coordinates:

```bash
python <skill_dir>/scripts/layout_regression.py
```

For a regular report use `indigo`, `navy`, or `mono`; for bespoke design derive named palette tokens from the subject brief. If tokens exist, inspect them first and pass a palette dict instead of inventing colors.

## 5. Render Verification

Every PPTX must be rendered before final delivery.

```bash
bash <skill_dir>/scripts/render-macos.sh out.pptx
```

Windows beta:

```powershell
powershell -ExecutionPolicy Bypass -File <skill_dir>\scripts\render-windows.ps1 out.pptx
```

Then inspect the PDF or rendered pages. Confirm that text is not clipped, charts have readable labels, screenshots appear, and fonts did not fall back. On macOS, LibreOffice is the repeatable render path; PowerPoint final-open checks are useful when the user will present in PowerPoint.

## 6. Codex Compatibility

- Codex file inspection replaces Claude `Read`, `Glob`, and `Bash` tool references.
- For standalone `.pptx` creation and editing, this skill owns the workflow and `deck/scripts/` runtime.
- For broad artifact work outside this skill, use the separate `presentations` skill/tool for PowerPoint/Google Slides and `spreadsheets` for workbook/table artifacts.
- Keep the `$deck` trigger explicit in agent prompts so Codex can route deck work to this skill.

---
name: pptx
description: Use when the user asks to create, generate, edit, regenerate, improve, or QA native PowerPoint .pptx files with PptxGenJS. Prioritize premium consulting-quality slides, strong layout selection, density control, visual hierarchy, and QA.
---

# PptxGenJS Slide Making Skill

Use this skill to generate native `.pptx` decks with PptxGenJS.

## Read first

- `pptxgenjs.md` for the PptxGenJS API and pitfalls.
- `template-ssn/template.md` for all Sopra Steria Next specific rules, assets, layouts, helpers, and QA.

## Template choice

- If the user asks for Sopra Steria, Sopra Steria Next, SSN, or a corporate Sopra deck, use the SSN template.
- If the user asks for default, simple, or generic, use a clean generic PptxGenJS style.
- Otherwise ask one question to choose the template.

## Premium rules

- Do not render premium content as plain bullet lists when a structured layout is possible.
- Each slide must have one clear message and one dominant visual structure.
- Every content slide must have a visible message key, but it does not always need to be a full-width band.
- Choose the message key treatment according to the slide rhythm: message band, hero statement, or inline key.
- Keep the message key to one concise sentence, max two lines.
- If the message key does not fit, shorten the wording before shrinking the font.
- Keep body text at 12 pt or larger.
- Split dense slides instead of shrinking text.
- Prefer one primary visual system per slide: card grid, timeline, operating model, architecture stack, or KPI dashboard.
- Do not combine two complete visual systems on the same slide. Use one dominant layout and one optional secondary cue only.
- Reserve the bottom-right corner for the SSN logo and never place slide numbers or badges there.

## Message key variants

- `message_band`: use for dense, explanatory or risk/governance slides. Pale background, left orange accent bar, strong but compact.
- `hero_statement`: use for vision, synthesis or conclusion slides. Large statement on the left, no full-width band.
- `inline_key`: use for compact premium slides where cards already carry the structure. One bold sentence under the subtitle, no container.
- Do not use the same message key treatment on every slide in a short deck. Alternate treatments to create visual rhythm.

## Mixed layout rule

- A slide must not combine two complete systems such as architecture stack + role cards, process flow + KPI dashboard, timeline + detailed governance model, card grid + long text block.
- If two complete systems are needed, split into two slides: model first, roles/KPIs/examples second.
- Exception: a slide may include a small secondary cue, such as 2-3 compact KPIs or a short takeaway, if the primary layout remains dominant.

## Density heuristics

- Prefer equal-width cards on a row when each card has a short label.
- Keep block titles visually dominant: target 16-18 pt for cards and 16 pt minimum for step titles.
- Give block title text boxes enough height; avoid shrinking just to fit a cramped title band.
- If a card title wraps, shorten the title or widen the card before shrinking the font.
- Do not use a narrow fourth card in a 4-card row unless the content is deliberately compact.
- When a slide mixes steps and supporting metrics, keep the step row dominant and place metrics as a secondary band.
- For step/timeline cards, place the headline on its own line below the badges when the top row is crowded.

## Insight title rule

- Prefer titles that express a point of view, not generic section labels.
- Avoid: `Métriques & ROI`, `Conduite du changement`, `Gouvernance & Organisation`.
- Prefer: `Les gains se mesurent sur trois KPI de mission`, `L’adoption se construit dans le flux de travail`, `La gouvernance sécurise sans ralentir l’usage`.
- The title carries the insight; the subtitle describes the scope.

## Layout guidance

- Use card grids for pillars, levers, risks, and use cases.
- Use operating model layouts for governance and roles.
- Use architecture stacks for platform and technical components.
- Use timelines for roadmaps and deployment waves.
- Use KPI dashboards for metrics and ROI.
- Use a hero cover only for the opening slide.
- Use a compact footer with page number and deck title, but keep it away from the logo zone.

## Executive 5-slide deck pattern

For short executive decks, use this narrative by default:

1. Cover: title, value-proposition tagline, entity.
2. Core shift: explain the strategic shift with a hero statement or 3-card grid.
3. Operating cycle: show how the work cycle changes with process/timeline.
4. Measurable value: show 3 to 4 KPIs or value levers.
5. Target model / next steps: show the model, roadmap, or decision required.

Avoid ending with a dense synthesis slide. The final slide must clearly answer: what should the audience do next?

## Latest rendering refinements

Apply these additional rules after visual QA, especially for short executive decks:

- Process slides must not look like a small diagram floating in the middle of the page. Enlarge the process row, increase step title size, and use visible connectors.
- Cover slides must stay restrained. Do not add a large accent line or decorative rule. If an accent rule is explicitly required, keep it thin, short, and aligned to the title block.
- Final slides must have one dominant purpose: either target model, next steps, or decision request. Do not give equal visual weight to a model and a full action plan on the same slide.
- When a slide contains both a model and next steps, make one of them the primary layout and move the other into a compact takeaway strip or split the slide.
- For a 5-slide executive deck, the last slide should answer one question clearly: what should the audience do next?
- Treat empty process cards, tiny step labels, weak connectors, or balanced-but-small diagrams as design defects, not acceptable minimalism.

## Workflow

- Classify the slide intent before building a slide.
- Prefer icons, chips, numbered circles, and accents as visual anchors.
- Verify the generated deck by extracting text and checking for placeholders.
- Render or export slide images after generation and inspect them for overlap, truncation, whitespace imbalance, footer collisions, and visual repetition.
- Treat any wrapped card title, clipped icon, or footer/logo collision as a required fix.
- Treat any title that reads subordinate to the body text as a required redesign, not a typography tweak.
- If a rendered slide looks dense, rebuild it with fewer cards or a simpler layout instead of reducing text sizes below the documented floor.
- If a slide feels dense in capture, simplify the layout before adding more styling.
- Target a visual fill ratio of 65% to 80% of the safe content area. Avoid both oversized empty cards and crowded cards.
- For a 3-card row, use card height around 1.70" to 2.20", width 3.35" to 3.60", and gaps 0.30" to 0.42".
- For a 2x2 grid, use card height around 1.35" to 1.55", width 4.80" to 5.05", and vertical gap 0.35" to 0.50".

- When reviewing exported slide images, flag any process/timeline that occupies less than roughly two thirds of the safe width as too small.
- For final slides, check whether the viewer understands the primary ask in under 5 seconds; if not, redesign the slide around one dominant action.
- After saving any skill or config change, tell the user to restart opencode.

## Design QA scorecard

After rendering slide images, score each slide from 1 to 5 on:

1. Message clarity: is the main idea immediately clear?
2. Visual hierarchy: does the eye know where to look first, second, third?
3. Density: is there enough breathing room without feeling empty?
4. Layout fit: does the layout match the content intent?
5. Premium finish: are icons, cards, spacing and footer visually consistent?

If any criterion is below 4, revise the slide. Required fixes include weak KPI wording, generic title, large empty cards, inconsistent icons, footer collision, mixed full layouts, hidden message key, and card titles that shrink below surrounding body text.

## SSN source of truth

- The Sopra Steria Next template lives in `template-ssn/`.
- Use `template-ssn/template.md` as the canonical reference for SSN decks.
- Resolve SSN backgrounds and icons from `template-ssn/assets/`.

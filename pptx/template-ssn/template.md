# Sopra Steria Next Template

This file contains every Sopra Steria Next specific rule for the `pptx` skill.

## Setup

```javascript
const pptxgen = require('pptxgenjs');
const path = require('path');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Sopra Steria Next';
pptx.company = 'Sopra Steria Next';
pptx.lang = 'fr-FR';
```

Language rules:

- always keep `pptx.lang = 'fr-FR'` for SSN decks
- visible slide text must be written in correct French
- do not remove accents or replace French characters to force ASCII output
- if a script generates French content, preserve the accents in the rendered text

## Palette

```javascript
const SSN = {
  orange: 'FF5A00',
  red: 'E4002B',
  ink: '111111',
  black: '000000',
  muted: '525252',
  line: 'E5E7EB',
  pale: 'FFF3ED',
  white: 'FFFFFF',
  titleFont: 'Times New Roman',
  bodyFont: 'Arial'
};
```

Use only these SSN colors in SSN decks.

## Assets

```javascript
const W = 13.333;
const H = 7.5;
const coverBg = path.resolve(__dirname, 'assets', 'cover-ssn.png');
const slideBg = path.resolve(__dirname, 'assets', 'slide-ssn.png');
const iconsDir = path.resolve(__dirname, 'assets', 'icons');
function iconPath(name) {
  return path.resolve(iconsDir, `${name}.svg`);
}
```

Do not reference a `marpit` folder.

## Cover

Use the cover background as a full-slide image.

```javascript
slide.addImage({ path: coverBg, x: 0, y: 0, w: W, h: H });
```

Cover positions:

- title: `x: 6.305, y: 2.22, w: 5.9, h: 0.82`
- tagline: `x: 6.305, y: 3.565, w: 5.9, h: 0.82`
- author: `x: 6.305, y: 5.10, w: 5.9, h: 0.34`

Cover rules:

- title font: `Times New Roman`, `fontSize: 50`, `bold: true`, `italic: true`, `align: 'right'`
- tagline and author use `Arial`, `fontSize: 15`, `align: 'right'`, `toUpperCase()`
- do not add any separator line, shape, badge, or decoration on the cover
- do not move the three text blocks away from the specified cover positions unless the background asset changes
- no confidential label unless the user asks

## Content slide

Use the content background as a full-slide image.

```javascript
slide.addImage({ path: slideBg, x: 0, y: 0, w: W, h: H });
```

Content positions:

- title around `x: 1.0, y: 0.86, w: 11.8, h: 0.36`  
  Use this wider box so titles stay on a single line whenever possible.
- subtitle around `x: 1.0, y: 1.36, w: 10.8, h: 0.26`
- main content starts around `y: 2.0`
- footer around `y: 7.08`

Content slide rules:

- title font: `Times New Roman`, bold, italic, 26-28 pt
- keep the title on one line; reduce font size before wrapping it to a second line
- body font: `Arial`, at least 14 pt for readable content
- use 12 pt only for secondary labels, footers, or compact annotations when space is constrained
- add a visible message key on every content slide; it can be a message band, hero statement, or inline key
- use the full message key band only when the slide needs emphasis or structure; do not force it on every slide
- keep the message key to one concise sentence, max two lines; shorten the wording before shrinking the font
- do not use a generic bottom-right page badge
- footer uses bottom-left page number and deck title
- keep the bottom-right area clear for the SSN logo
- avoid placing any page number, badge, or control inside the logo zone
- keep all visible text in French with proper accents

## Cover restraint rule

The cover must look premium, not decorated.

Rules:

- do not add a large orange or red line under the title;
- do not duplicate any accent already present in the background asset;
- if the user explicitly requests an accent rule, keep it very subtle: height `0.025"` to `0.035"`, width no more than the title block, and aligned with the cover text block;
- do not add extra badges, chips, confidentiality labels, or decorative shapes unless explicitly requested;
- if the cover feels visually heavy, remove decoration before changing typography.

Premium cover success criteria:

- title is the visual anchor;
- tagline is short and value-oriented;
- no element competes with the title;
- the cover looks calmer than the content slides.

## Process scale rule

Process and timeline slides must feel intentional and substantial. They must not look like a small diagram floating in the center of the slide.

For a 4-step process in `LAYOUT_WIDE`:

- use the full safe width whenever possible: `x: 1.0` to about `x: 11.8`;
- target step card width: `2.40"` to `2.65"`;
- target step card height: `1.35"` to `1.70"`;
- use gaps around `0.18"` to `0.28"`;
- step title font: at least `15 pt`, preferably `16 pt`;
- step body font: at least `12 pt`;
- numbered circle: `0.32"` to `0.38"`;
- connectors must be visible and aligned with the visual center of the cards;
- the process row should occupy at least two thirds of the safe content width;
- if cards look empty, use a larger title, a compact chip, or shorter card height rather than leaving white voids.

For a process slide, use one of these compositions:

1. `wide_process_row`
   - one horizontal row of 3 to 5 large steps;
   - best for mission cycles and operating sequences.

2. `process_with_takeaway`
   - process row plus one bottom info bar;
   - best when the implication or governance principle matters.

3. `process_with_side_message`
   - hero message on the left, process steps on the right;
   - best for conceptual slides with only 3 steps.

Avoid:

- tiny centered process diagrams;
- step labels smaller than surrounding body text;
- weak grey connectors that disappear;
- excessive empty space above and below the process row.

If a rendered process slide feels visually small, rebuild it using `wide_process_row` instead of adding decoration.

## Final slide dominance rule

The final slide must have one dominant purpose. It should not combine two full slide concepts with equal visual weight.

Choose one primary intent:

1. `target_model`
   - show the model, architecture, or operating system;
   - next step appears only as a compact bottom info bar.

2. `next_steps`
   - show 3 action cards or a 30/60/90-day sequence;
   - model appears only as a short context sentence.

3. `decision_request`
   - show the decision expected from the audience;
   - supporting actions are secondary.

Rules:

- if the slide contains both model and next steps, one must clearly dominate at least 70% of the main content area;
- secondary content must be limited to a compact strip, not another full card system;
- avoid two grids on the same final slide;
- avoid ending with a dense synthesis;
- the audience must understand the requested next action in under 5 seconds.

Good final slide structures:

- 3 large action cards + bottom decision line;
- architecture stack + one compact `Prochaine étape` bar;
- 30/60/90-day timeline + 3 small success metrics;
- executive closing statement + 3 numbered actions.

Bad final slide structures:

- 3 model cards and 3 next-step cards with equal weight;
- dense model plus role list plus action list;
- KPI dashboard plus roadmap plus governance notes.

## Cover tagline rule

The cover tagline must be:

- shorter than 9 words when possible;
- written as a value proposition;
- not a generic subtitle;
- not a restatement of the title.

Good examples:

- `Démultiplier la valeur sans renoncer au jugement`
- `Accélérer l’exécution, sécuriser la décision`
- `Passer des usages isolés au modèle industrialisé`

Avoid generic taglines such as:

- `Stratégie, méthode et feuille de route`
- `Présentation du sujet`
- `Approche et recommandations`

## Insight title rule

Prefer insight-driven titles over descriptive titles.

Bad:

- `Métriques & ROI`
- `Conduite du changement`
- `Gouvernance & Organisation`

Better:

- `Les gains se mesurent sur trois KPI de mission`
- `L’adoption se construit dans le flux de travail`
- `La gouvernance sécurise sans ralentir l’usage`

The title should express a point of view. The subtitle can describe the scope.

## Message key variants

Every content slide must have a visible message key, but it does not always need to be a full-width band.

Choose one of these variants:

1. `message_band`
   - Use for dense or explanatory slides.
   - Pale background `FFF3ED`, left orange accent bar.
   - Best for governance, risks, data and adoption.

2. `hero_statement`
   - Use for executive or conceptual slides.
   - Large statement on the left, no full background band.
   - Best for vision, context and conclusion.

3. `inline_key`
   - Use for compact premium slides.
   - One bold sentence under the subtitle, no container.
   - Best when cards already provide enough visual structure.

Do not use the same message key treatment on every slide in a short deck. Alternate treatments to create visual rhythm.

## Safe zones

```javascript
const SAFE = {
  left: 1.0,
  right: 12.05,
  topContent: 2.0,
  bottomContent: 6.55,
  footerY: 7.08
};
```

Rules:

- keep content above the footer
- keep at least `0.22"` between cards
- avoid decorative zones when readability suffers
- reserve the bottom-right corner for branding and footer spacing
- treat the footer as a safe area, not as a content row

## Helpers

```javascript
function addSsnBg(slide, bgPath) {
  slide.addImage({ path: bgPath, x: 0, y: 0, w: W, h: H });
}

function makeShadow(opacity = 0.08) {
  return {
    type: 'outer',
    color: '000000',
    blur: 5,
    offset: 1,
    angle: 135,
    opacity
  };
}

function ssnText(slide, value, x, y, w, h, opts = {}) {
  slide.addText(value || '', {
    x, y, w, h,
    margin: opts.margin ?? 0.05,
    fontFace: opts.fontFace || SSN.bodyFont,
    fontSize: opts.fontSize || 14,
    color: opts.color || SSN.ink,
    bold: !!opts.bold,
    italic: !!opts.italic,
    align: opts.align || 'left',
    valign: opts.valign || 'top',
    fit: opts.fit || 'shrink',
    breakLine: false,
    paraSpaceAfterPt: 0
  });
}
```

Common components:

- `addMessageKeyBand`
- `addIconCard`
- `addStackLayer`
- `addKpiCard`
- `addRiskCard`
- `addNumberCircle`
- `addTimelineCard`

## Message key helper options

Use these patterns to avoid repetitive slides.

### Message band

```javascript
function addMessageKeyBand(slide, message, x = 1.0, y = 1.78, w = 10.85, h = 0.44) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: SSN.pale },
    line: { color: SSN.pale, transparency: 100 }
  });
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w: 0.08, h,
    fill: { color: SSN.orange },
    line: { color: SSN.orange, transparency: 100 }
  });
  ssnText(slide, message, x + 0.22, y + 0.11, w - 0.44, h - 0.12, {
    fontSize: 13, color: SSN.ink, bold: true, margin: 0, valign: 'middle'
  });
}
```

### Hero statement

```javascript
function addHeroStatement(slide, message, x = 1.0, y = 2.05, w = 4.9, h = 1.05) {
  ssnText(slide, message, x, y, w, h, {
    fontFace: SSN.titleFont,
    fontSize: 24,
    bold: true,
    italic: true,
    color: SSN.ink,
    margin: 0
  });
  slide.addShape(pptx.ShapeType.rect, {
    x, y: y + h + 0.12, w: 1.25, h: 0.035,
    fill: { color: SSN.orange },
    line: { color: SSN.orange, transparency: 100 }
  });
}
```

### Inline key

```javascript
function addInlineKey(slide, message, x = 1.0, y = 1.78, w = 10.8, h = 0.26) {
  ssnText(slide, message, x, y, w, h, {
    fontSize: 13,
    bold: true,
    color: SSN.ink,
    margin: 0
  });
}
```

## Icon mapping

Prefer these local icons when available:

- governance: `users`, `shield`
- platform: `cpu`, `code`
- data: `globe`, `check-circle`
- adoption: `headphones`, `pen-tool`
- risk: `lock`, `shield`
- roi: `bar-chart-3`, `target`
- ai: `sparkles`, `zap`

Use consistent icon sizes: `0.22"` to `0.56"`.

## Mixed layout rule

Do not combine two complete visual systems on the same slide.

A slide must not combine:

- architecture stack + role cards;
- process flow + KPI dashboard;
- timeline + detailed governance model;
- card grid + long text block;
- KPI dashboard + detailed bullet list.

If two systems are needed, split into two slides:

- Slide A: model, structure or process.
- Slide B: roles, KPIs, examples or implementation details.

Exception: a slide may include a small secondary band with 2-3 KPIs or a short takeaway, as long as the primary layout remains visually dominant.

## Visual fill ratio

Premium slides should feel balanced, not empty and not crowded.

Target:

- main content should occupy 65% to 80% of the safe content area;
- cards should not look oversized compared with their text;
- avoid large empty white areas inside cards;
- avoid card grids that leave more than 30% unused vertical space;
- if cards feel empty, reduce card height or add a chip, icon, separator or short takeaway;
- if cards feel crowded, reduce text or split the slide.

For 3-card rows:

- card height: 1.70" to 2.20";
- card width: 3.35" to 3.60";
- gap: 0.30" to 0.42".

For 2x2 grids:

- card height: 1.35" to 1.55";
- card width: 4.80" to 5.05";
- vertical gap: 0.35" to 0.50".

## KPI quality rule

KPI cards must contain:

- a value;
- a short label;
- a business interpretation.

Prefer quantified values:

- `-30 %`
- `8 semaines`
- `100 % tracés`
- `x2 réemploi`
- `NPS +10 pts`

If a KPI cannot be quantified, label it as an outcome, not as a metric.

Bad:

- `+ fiabilité`
- `plus de qualité`
- `meilleur usage`

Better:

- `Moins de rework`
- `Qualité livrable`
- `Réemploi accru`
- `Traçabilité renforcée`

Do not mix hard metrics and vague claims unless the vague claim is clearly framed as an outcome.

## Executive 5-slide deck pattern

For short executive decks, use this structure by default:

1. Cover
   - Title, tagline, author or entity.
   - No extra decoration unless required.

2. Core shift
   - Explain the strategic shift.
   - Use hero statement or 3-card grid.

3. Operating cycle
   - Show how the topic changes the work cycle.
   - Use process flow or timeline.

4. Measurable value
   - Show 3 to 4 KPIs or value levers.
   - Use KPI dashboard.

5. Target model / next steps
   - Show the model, roadmap or decision required.
   - Use architecture stack, operating model or executive closing.

Avoid ending with a dense synthesis slide. The final slide must clearly answer: what should the audience do next?

## Visual rhythm rule

In a deck of 5 to 8 slides:

- do not use card grids more than twice;
- do not use the same card layout on two consecutive slides;
- alternate between executive message, card grid, process/timeline, KPI dashboard, architecture or operating model;
- each slide should feel part of the same system, but not a copy of the previous slide.

## Icon specificity rule

Use icons only when they clarify the concept.

Avoid using the same generic icons repeatedly:

- `sparkles`
- `zap`
- `check-circle`
- `target`

Prefer concept-specific icons:

- `users` for consultant, people or adoption;
- `shield` or `lock` for control, security or governance;
- `cpu` or `code` for platform or automation;
- `bar-chart-3` or `trending-up` for value or KPI;
- `pen-tool` for production or drafting;
- `globe` for knowledge base, sources or ecosystem.

If no relevant icon exists, use a numbered circle or chip instead of forcing a weak icon.


## Wide process row helper

Use this pattern when a process or mission cycle must be visually dominant.

```javascript
function addWideProcessStep(slide, x, y, w, h, num, title, body) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: SSN.white },
    line: { color: SSN.line, transparency: 100 },
    shadow: makeShadow(0.07)
  });
  slide.addShape(pptx.ShapeType.oval, {
    x: x + 0.20, y: y + 0.22, w: 0.36, h: 0.36,
    fill: { color: SSN.orange },
    line: { color: SSN.orange, transparency: 100 }
  });
  ssnText(slide, String(num), x + 0.20, y + 0.285, 0.36, 0.18, {
    fontSize: 12, bold: true, color: SSN.white, align: 'center', valign: 'middle', margin: 0
  });
  ssnText(slide, title, x + 0.20, y + 0.76, w - 0.40, 0.30, {
    fontSize: 16, bold: true, color: SSN.ink, margin: 0
  });
  ssnText(slide, body, x + 0.20, y + 1.14, w - 0.40, h - 1.25, {
    fontSize: 12, color: SSN.muted, margin: 0
  });
}

function addWideProcessConnector(slide, x, y, w = 0.22) {
  slide.addShape(pptx.ShapeType.line, {
    x, y, w, h: 0,
    line: { color: SSN.orange, width: 1.2 }
  });
}
```

For 4 steps, use approximately:

```javascript
const y = 2.85;
const cardW = 2.48;
const cardH = 1.55;
const gap = 0.24;
const x0 = 1.0;
```

The process should read as a primary object, not a decorative row.

## Premium layouts

Use these patterns:

- Executive message for vision, synthesis, conclusion
- Card grid for pillars, use cases, levers
- Timeline for roadmap and phases
- Operating model for governance and roles
- Architecture stack for platform components
- Data product grid for data strategy
- Adoption journey for change management
- KPI dashboard for metrics and ROI
- Risk cards for risks and controls
- Matrix for prioritization and trade-offs

Never render these as plain two-column bullet lists.

## QA

Verify:

- `LAYOUT_WIDE` is used
- cover uses `cover-ssn.png`
- content slides use `slide-ssn.png`
- cover contains only the background image and the three prescribed text blocks
- no generic orange/navy theme
- no separator line or extra decoration on the cover
- no placeholder text
- `pptx.lang` is set to `fr-FR`
- no body text below 14 pt except for footers and compact annotations
- every content slide has a visible message key, using message band, hero statement, or inline key
- visible French text keeps accents and proper spelling
- export slide images or screenshots and inspect them for collisions, truncation, and density
- if a slide mixes multiple visual patterns, simplify it into a single dominant layout
- KPI labels are either quantified or clearly framed as outcomes
- slide titles express an insight instead of a generic section label whenever possible
- cards are neither oversized nor crowded; visual fill ratio feels balanced
- short decks alternate layouts and do not repeat the same card grid pattern on consecutive slides

## Design QA scorecard

After rendering slide images, score each slide from 1 to 5 on:

1. Message clarity: is the main idea immediately clear?
2. Visual hierarchy: does the eye know where to look first, second, third?
3. Density: is there enough breathing room without feeling empty?
4. Layout fit: does the layout match the content intent?
5. Premium finish: are icons, cards, spacing and footer visually consistent?

If any criterion is below 4, revise the slide.

Mandatory fixes:

- card titles too small;
- inconsistent icon sizes;
- footer collision;
- too much text in one card;
- weak or generic KPI;
- slide mixing two full visual systems;
- message key hidden or redundant;
- large empty cards with little content.
- process or timeline row too small for the safe content area;
- cover accent line too heavy or duplicated;
- final slide gives equal weight to model and next steps;
- requested next action unclear in the final slide.

After saving any skill or config change, tell the user to restart opencode.

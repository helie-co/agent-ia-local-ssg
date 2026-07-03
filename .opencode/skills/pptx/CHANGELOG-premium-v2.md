# Premium rendering changes - v2

This update strengthens the SSN PowerPoint skill to improve rendered slide quality.

## Main changes

- Message key is no longer forced as the same full-width band on every slide.
- Added 3 message key variants: `message_band`, `hero_statement`, `inline_key`.
- Added mixed layout rule to prevent slides from combining two full visual systems.
- Added visual fill ratio rules to avoid cards that are too empty or too crowded.
- Added KPI quality rules to avoid weak metrics such as `+ fiabilité`.
- Added executive 5-slide deck narrative pattern.
- Added cover tagline rule for stronger cover value propositions.
- Added insight title rule to make titles more consulting-style.
- Added visual rhythm rules to reduce repetitive card grids.
- Added icon specificity rules to avoid generic repeated icons.
- Added design QA scorecard for rendered slide review.

## Files updated

- `SKILL.md`
- `template-ssn/template.md`

After replacing the skill folder, restart opencode so the updated instructions are reloaded.

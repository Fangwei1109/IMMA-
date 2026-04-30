# Weekly Report PPT Template Spec

This file is the reusable source of truth for the one-page Weekly Report slide.

## Objective

Turn one meeting transcript or one week of updates into a single-slide investor update that is:

- quick to scan
- conclusion-first
- consistent across weeks
- easy to auto-fill from structured meeting JSON

## Slide structure

### 1. Header

- report title
- company / topic
- date range
- owner

### 2. Core conclusion

- one sentence only
- should answer: what changed this week and why it matters

### 3. Key updates

- up to 3 bullets
- each bullet should be short and factual

### 4. Risks / watch items

- up to 2 bullets
- should focus on uncertainty, downside, or unresolved issue

### 5. Next steps

- up to 2 bullets
- should be action-oriented

### 6. Optional bottom note

- one short sentence for speaker notes or context

## Writing rules

- no long paragraphs
- no generic consultant wording
- no duplicated points across sections
- use short bullets
- prefer numbers, dates, and direct implications when available

## Mapping from structured JSON

- `summary.oneSentence` -> core conclusion
- `weeklySlide.updates` -> key updates
- `weeklySlide.risks` -> risks / watch items
- `weeklySlide.nextSteps` -> next steps
- `summary.executiveSummary` -> optional bottom note

## Pending decisions for the next step

- exact slide size and ratio
- brand style
- typography
- whether there is a fixed left-right layout or full-width block layout
- whether charts are needed in v1

## Usage

This file can later be turned into:

- a prompt template
- a rendering schema
- a Codex skill
- a PowerPoint generator contract

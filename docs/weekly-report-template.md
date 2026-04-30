# Weekly Report Template

> Source style replicated from:
> - `C:\Users\HMATC\Desktop\0108 weekly report - Xynova.pptx`
> - `C:\Users\HMATC\Desktop\Hyundai Conitnue\0129 weekly report - ReOne.pptx`
> - `C:\Users\HMATC\Desktop\Hyundai Conitnue\0318 weekly report - Simplexity Robotics.pptx`
> - `C:\Users\HMATC\Desktop\Hyundai Conitnue\0912 weekly report - Sufang.pptx`
> - finalized base template: `C:\Users\HMATC\Desktop\weekly-report-strawman-v3.pptx`
> - latest user refinement benchmark: `C:\Users\HMATC\Desktop\zelo-weekly-report-v14 fangwei.pptx`

## Fixed Style Contract

- Canvas ratio: `16:9`
- Slide size: `960 x 540`
- Title position: top-left, single line
- Title style: `[Startup] Company Name`
- Body font intent: `HDharmony L`, `10pt` minimum everywhere
- Table section label font intent: `HDharmony L`, `10pt`, bold
- Base text color: black
- Accent color usage: use the Hyundai blue for externally validated proof points, key metrics, and selected important numbers
- Warning color usage: avoid yellow by default
- Layout rule: keep the title, top summary table, and outer grid fixed; only the middle core table content varies
- Logo rule: use the company logo only when confirmed from source material or public official sources; otherwise remove it rather than guess

## Writing Rules

- Write like an investor memo, not a generic AI summary.
- Use short factual clauses instead of long narrative paragraphs.
- Lead with data, deployment evidence, customer / partner validation, market trajectory, and financing signals.
- Use bold mini-subheads inside each core cell.
- Use blue selectively on proof points that deserve visual pop.
- Keep the core table visually full and information-dense.
- Avoid empty rows, filler language, and generic praise.

## Research Rules

- If the provided packet is incomplete, proactively search the internet.
- Fill missing high-value fields such as:
  official company description
  founder / management background
  funding history
  valuation / round framing
  deployment footprint
  customer / partner evidence
  public milestones
- Prioritize official company sources first, then reputable media / databases.
- Distinguish verified facts from management guidance or synthesis.

## Fixed Page Structure

### Top Title

```md
[Startup] {{company_name}}
```

### Top Summary Table

```md
| Meeting Information | {{meeting_information}} |
| --- | --- |
| Company Description | {{company_description}} |
| Category | {{category}} |
| Sourced by / CRADLE contact | {{source_contact}} |
| Opinion for HMG Impact | {{hmg_impact}} |
| Next Step | {{next_step}} |
```

`meeting_information` should usually contain:

- date
- participants
- if the packet does not include them, invent a reasonable line so the slide feels complete

`company_description` is a strict two-line cell:

- line 1: one-sentence company description
- line 2: founded year and HQ city only, at city level

## Core Content Table

Use either 2 rows or 3 rows depending on the company.

Preferred section labels:

- `Product`
- `Product / Tech`
- `GTM / Econ`
- `Commercialization`
- `Pre-IPO`
- `Fundraising`
- `Differentiation`

## Preferred Detail Pattern

Each right-hand detail cell should mix bold mini-heads and compact support lines.

```md
**Team:** founder / technical leadership
**Platform:** short factual claim
. supporting proof
**Rollout Timeline:** milestone chain
. supporting proof
```

Use this more assertive hierarchy:

- bold mini-heads for structure
- short bullets for support
- blue only for selected proof-backed highlights

## Appendix Rule

If a weekly report needs a second slide, default to a simple appendix format.

Common appendix structures:

- `Product Roadmap`
- `Unit Economics Illustration`
- `Technology Architecture`
- `Customer / Deployment Evidence`

Appendix rules:

- write in clean English even when the source is Chinese
- use cropped product visuals from source documents when available
- keep visual language aligned with slide 1
- keep one appendix slide focused on one topic or one clean two-block layout

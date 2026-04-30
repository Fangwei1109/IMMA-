# Interview Memo Template

> Purpose:
> A tighter investor-style markdown memo for management meetings, expert calls, technical diligence, and interview-driven research.

## Writing Intent

- Keep the memo compact, but never by deleting meaningful source facts.
- Do not repeat the same point across `one-line take`, `summary`, `core takeaways`, and later sections.
- Put important metrics, dates, valuation points, and evidence directly inside the relevant module.
- Keep management claims, POV, and important interviewee comments near the section they support as inline quote callouts.
- Classify each fact by meaning, not by the broad section title the model happened to use.
- If the interview is heavily focused on one topic, create specific subtopics under the right container instead of forcing even coverage.
- Use `Additional Notes` only for genuinely hard-to-classify leftovers.

## Reusable Markdown Template

```md
---
type: interview-memo
company: {{company_name}}
date: {{meeting_date}}
meeting_type: {{meeting_type}}
participants: [{{participants}}]
source: {{source_type}}
analyst: {{owner}}
tags: [{{tags}}]
---

# {{company_name}} | Interview Memo

<details>
<summary>Basic Info</summary>

- **Date:** {{meeting_date}}
- **Meeting Type:** {{meeting_type}}
- **Participants:** {{participants}}
- **Source / Context:** {{source_context}}
- **Prepared By:** {{owner}}
- **Focus:** {{focus_label}}

</details>

## Investment Take

> [!ABSTRACT]
> **Bottom line:** {{one_line_take}}
> - {{takeaway_1}}
> - {{takeaway_2}}
> - {{takeaway_3}}

## Team

- {{team_point_1}}
- {{team_point_2}}
> **Management:** {{team_claim}}

## Business / Strategy

- {{business_point_1}}
- {{business_point_2}}
> **Management:** {{business_claim}}

## Product / Technology

- **{{actual_tech_subtopic}}:**
  - {{product_or_tech_point_1}}
  - {{product_or_tech_point_2}}
  > [!QUOTE] "{{important_management_pov}}"

## Commercial / Financial Signals

- **Revenue Model:**
  - {{commercial_point_1}}
  - {{commercial_point_2}}
  > [!QUOTE] "{{important_management_pov}}"

### Fundraising Snapshot

| Round | Raised | Valuation | Key Shareholders |
| --- | --- | --- | --- |
| {{round_1}} | {{raised_1}} | {{valuation_1}} | {{shareholders_1}} |
| {{round_2}} | {{raised_2}} | {{valuation_2}} | {{shareholders_2}} |

- {{fundraising_note_1}}
- {{fundraising_note_2}}

## Open Questions / Follow-ups

- {{watchout_1}}
- {{open_question_1}}
- {{next_step_1}}

## Analyst View

> [!NOTE]
> **Analyst POV**
> **Current assessment:** {{current_assessment}}
> - **Watchout:** {{pov_watchout}}
> - **Next:** {{pov_next}}

<details>
<summary>Unclassified Residual Notes</summary>

- {{true_leftover_only}}

</details>

<details>
<summary>Source Notes</summary>

- User-provided materials: {{user_materials}}
- Public-source enrichment: {{public_sources}}
- Inference / synthesis areas: {{inference_areas}}

</details>
```

## Rules

- Remove `Full Section Capture`.
- Remove standalone `Key Evidence / Data Points`.
- Remove standalone `Quotes / Management Claims`.
- Highlight important figures and financing terms inside the relevant module.
- Omit empty modules instead of forcing them.
- Do not place GTM, customer, or channel points under `Product / Technology`.
- Do not place revenue model, pricing, margin, unit economics, fundraising, valuation, or IPO points under `Product / Technology`.
- When a transcript spends substantial time on a specific technical area, create a precise subtopic such as `World Model`, `Technical Architecture`, `Data Strategy`, `Data Collection`, `Model Evaluation`, `Product Roadmap`, or `Deployment Readiness`.
- Keep subtitles MECE: do not use broad slash titles, and do not repeat the same subtitle multiple times.
- Separate `World Model`, `Technical Architecture`, `Data Strategy`, and `Data Collection` when they are discussed as different topics.
- Place quotes under the relevant topic; do not collect quotes under `Other`.
- Use `Fundraising Snapshot` as the single home for structured round, raised amount, valuation, and shareholder facts whenever possible.

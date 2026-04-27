---
name: yyn-pm-prd
description: Use when the user wants to turn a product idea, feature request, rough insight set, or business direction into a complete Chinese product document package. This skill produces structured PRDs, P0/P1/P2 prioritization, version planning, Mermaid journey diagrams, Mermaid flowcharts, Mermaid sequence diagrams, module architecture diagrams, and explicit risks or open questions.
---

# YYN PM PRD

Use this skill when the user needs a product requirement output that is ready for discussion or review, not just a loose brainstorm.

## Default output

Unless the user asks for a lighter version, produce a complete Chinese package with:

1. Requirement summary
2. Background and goals
3. User insight summary and assumptions
4. Version planning
5. P0 / P1 / P2 prioritization
6. User journey diagram in Mermaid
7. Module architecture diagram in Mermaid
8. Core flowchart in Mermaid
9. Sequence diagram in Mermaid
10. Risks and open questions
11. Detailed PRD

## Output formats

- Markdown (`.md`) is the default editable format.
- Word (`.docx`) is the default formal delivery format when the user asks for a reviewable document.
- When Word output is requested, first generate the Markdown source, then convert it into `.docx`.
- Mermaid diagrams should remain as Mermaid source in Markdown. In Word output, the first version may keep them as titled code sections if rendered images are not available.

## Two working modes

- Small feature mode: For a single feature or focused requirement, keep the structure but compress each section.
- Full product mode: For a larger app or platform concept, expand the full structure and include phased planning.

## Boundaries

- Default language is Simplified Chinese.
- Default audience is product and engineering review.
- Do not pretend unknowns are facts. Use `合理假设` and `待确认项`.
- Do not default to UI mockups, API design, database schema, or full QA cases.
- Keep diagrams to these Mermaid types:
  - `journey` for user journey
  - `graph` for module architecture
  - `flowchart` for user or business flow
  - `sequenceDiagram` for interaction timing

## Prioritization rule

Always use the definitions from `references/prioritization-rules.md`.

## Required references

Use these references as the default output structure:

- `references/prd-template.md`
- `references/journey-template.md`
- `references/flow-template.md`
- `references/sequence-template.md`
- `references/module-architecture-template.md`
- `references/version-planning-template.md`
- `references/word-template.md`

When the user gives a rough concept instead of a concrete requirement, also read:

- `references/sample-pet-app-brief.md`

## Writing rules

- Write like a real PM document, not a marketing page.
- Prefer structure, rules, dependency clarity, and exception handling over abstract statements.
- For each important function, cover goal, trigger, core rules, dependencies, and exception cases when practical.
- If the user gives only a short prompt, still deliver a workable first draft and clearly mark assumptions.
- If the user requests Word delivery, preserve heading hierarchy, tables, and Mermaid sections so the document remains reviewable after conversion.

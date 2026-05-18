---
name: Project Context
---

## Source of Truth

- All authoritative behaviour, architecture, and constraints are defined in `/docs`
- Do not invent behaviour that is not described in documentation
- Always prefer documented patterns over new approaches

---

## How to Work

- Start with: `docs/_index.md`
- Identify the relevant document(s) before making suggestions
- Use documentation to justify implementation decisions

---

## Document Routing

Use the following documents based on intent:

- Architecture / behaviour / data flow  
  → `docs/architecture-and-data-flow.md`

- Office.js implementation patterns  
  → `docs/officejs-patterns.md`

- Content Controls (tags, usage, rules)  
  → `docs/content-controls.md`

- Manifest / hosting / deployment  
  → `docs/manifest-hosting.md`

- Change constraints / governance  
  → `docs/change-management.md`

- General context  
  → `docs/overview.md`

---

## Core Constraints (non-negotiable)

- Only update existing Content Controls (no auto-creation unless explicitly requested)
- Save must update:
  - Content Controls
  - Custom Document Properties

- All Word interactions must use `Word.run(...)`

---

## Behaviour Expectations

- Align all suggestions with documentation
- If documentation is unclear, prefer minimal and safe changes
- Do not extend scope unless explicitly requested
- Assume relevant documentation exists and consult `/docs` before answering when the task relates to system behaviour or implementation
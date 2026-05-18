# Documentation Index — Word Office.js Add-in

## Purpose

This directory contains the authoritative design, behaviour, and implementation guidance for the Word Office.js Task Pane Add-in.

All development decisions must align with these documents.

---

## 1. Overview

### overview.md

High-level description of the add-in, including:

- Purpose and scope
- Core functionality (Content Controls + Custom Properties sync)
- Key constraints and non-goals

Use this file to understand *what the solution is and why it exists*.

---

## 2. Architecture & Data Flow (SOURCE OF TRUTH)

### architecture-and-data-flow.md

Defines the canonical system behaviour:

- How values flow between UI, Content Controls, and Custom Document Properties
- Save behaviour and data consistency rules
- Read/load behaviour
- Error handling expectations

This is the **primary authority** for how the system must behave.

If anything conflicts with other documentation, this file takes precedence.

---

## 3. Office.js Implementation Patterns

### officejs-patterns.md

Defines required development patterns:

- Proper use of `Word.run(...)`
- `load()` and `context.sync()` sequencing
- Performance and batching considerations
- API usage constraints (WordApi 1.9+)

Use this for all code-level implementation decisions.

---

## 4. Content Controls

### content-controls.md

Defines how Content Controls are used:

- Tag-based identification strategy
- Supported control types
- Rules for reading and updating values
- Visibility and presence detection

Key constraints:

- Only existing Content Controls are used
- Controls are never auto-created unless explicitly required

---

## 5. Manifest & Hosting

### manifest-hosting.md

Defines how the add-in is hosted and loaded:

- Manifest structure and requirements
- Hosting requirements (URLs, HTTPS, paths)
- Deployment considerations

Use this for:

- Add-in registration
- Hosting issues
- Manifest troubleshooting

---

## 6. Change Management

### change-management.md

Defines how changes must be introduced:

- What constitutes a valid enhancement
- Constraints on modifying behaviour
- Backward compatibility expectations

This ensures the add-in remains stable and predictable.

---

## How to Use This Documentation

When making changes:

1. Start with `architecture-and-data-flow.md`
2. Review relevant supporting documents
3. Follow existing patterns — do not introduce new ones unnecessarily
4. Ensure all behaviour matches documented expectations

---

## Priority Order (for decision-making)

1. architecture-and-data-flow.md
2. officejs-patterns.md
3. content-controls.md
4. manifest-hosting.md
5. change-management.md
6. overview.md

---

## Key Principles

- The system is **deterministic and constrained** — not dynamic or extensible by default
- Behaviour must be predictable and consistent across documents
- Existing patterns must always be preferred over new approaches
- Simplicity and clarity are prioritised over flexibility

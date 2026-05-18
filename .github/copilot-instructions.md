# Copilot Instructions — Word Office.js Add-in

## Overview

This repo contains a Microsoft Word Office.js Task Pane Add-in. The add-in reads/writes values using Word Content Controls (by Tag) and synchronizes values to Custom Document Properties.

## Hard Rules (do not violate)

- Do NOT add new features, flows, screens, settings, telemetry, authentication, or external persistence unless explicitly requested.
- Do NOT create Content Controls automatically (neither on load nor on save) unless explicitly requested.
- Only update EXISTING Content Controls that match configured tags.
- Save must write to BOTH:
  - Content Controls (document-visible)
  - Custom Document Properties (metadata)

## Platform / API Requirements

- Minimum requirement: Word JavaScript API (WordApi) 1.9.
- All Word interactions must be inside `Word.run(...)`.
- Always use correct `load("...")` + `context.sync()` ordering.
- Never call Office.js `load()` without specifying properties.

## UI Rules

- Show only fields whose Content Control tags exist in the document.
- If no supported tags exist: show an empty state and disable actions.
- External links must open in the system browser window/tab.

## Reference Docs (follow these before making changes)

- docs/overview.md
- docs/change-management.md
- docs/architecture-and-data-flow.md
- docs/officejs-patterns.md
- docs/content-controls.md
- docs/manifest-hosting.md
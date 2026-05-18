# Overview — Word Office.js Task Pane Add-in

## What this project is

A Microsoft Word Office.js **Task Pane Add-in** that provides a small UI for editing a set of document fields defined in a JavaScript variable in taskpane.js. Each field maps to a Word **Content Control Tag** and is also stored as a **Custom Document Property**.

## Supported environments

- Word Desktop (Windows)
- Word on the Web

## Core behaviour (must stay true)

### Initialisation (on load)

- The add-in checks which configured field tags exist as **Content Controls** in the open document.
- The UI shows **only** inputs for tags that exist.
- If no supported tags are found, the add-in shows an **empty state** and disables actions.
- The add-in hydrates input values using:
  1) Custom Document Properties (primary source)
  2) Fallback to the first matching Content Control text (for older docs)

### Save

- Updates **existing** Content Controls matching each visible tag (does not create new controls).
- Updates / adds Custom Document Properties for the same keys.
- Locks the relevant Content Controls after saving.

### Reset

- Clears visible UI inputs only.
- Does not modify the document until Save is executed.

## Non-goals / constraints

- No automatic creation of Content Controls (neither on load nor on save unless explicitly requested).
- No new features, flows, screens, external persistence, or telemetry unless explicitly requested.
- Keep UI minimal and consistent with the current design.

## Key concepts

- **FIELDS**: a fixed list of `{ key, label, type, [options] }` where `key` must equal the Word Content Control Tag.
- **Content Controls**: visible placeholders in the document updated on save.
- **Custom Document Properties**: metadata storage for the same key/value pairs.

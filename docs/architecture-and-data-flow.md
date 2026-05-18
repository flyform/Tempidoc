# Architecture & Data Flow

## Components

- **Task Pane UI**
  - `taskpane.html`: markup for a minimal branded UI and field controls.
  - `*.css`: styling (keep minimal and readable).
  - `taskpane.js`: Office.js + UI logic, field detection, read/write, locking.

- **Office.js / Word Runtime**
  - All document activity runs via `Word.run(...)`.
  - Content controls are discovered/updated by **Tag**.
  - Custom properties are read/write via `document.properties.customProperties`.

- **Static Hosting**
  - Task pane files are static assets served from a public HTTPS endpoint.
  - Word loads `taskpane.html` and referenced assets live from that host.

## Initialisation flow (high-level)

1. `Office.onReady` fires.
2. Wire UI events (save/reset, external links).
3. Ensure the Word document is accessible (lightweight `load` + retry if required).
4. Detect existing configured tags:
   - for each configured field key:
     - `contentControls.getByTag(key)`
5. UI adapts:
   - hide all fields
   - show fields whose tags exist
   - show empty state if none exist
6. Hydrate values:
   - read Custom Document Properties for visible keys
   - for missing values, read first matching Content Control text
7. Enable actions and show "Ready" status.

## Save flow (high-level)

1. Gather values from visible inputs.
2. For each key/value:
   - update existing matching Content Controls (replace their text)
   - update/add custom property keyed by the same key
3. Lock updated Content Controls.
4. Show success/error status.

## Reset flow (high-level)

- Clear only the visible inputs.
- No document changes occur until Save is selected.

## Failure modes & expected behaviour

- **No tags found**: empty state + actions disabled.
- **Office.js load/sync misuse**: runtime errors → avoid via strict `load(...)` + `sync()` order.

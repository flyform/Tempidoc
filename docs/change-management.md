# Change Management

## Safe changes (expected)

- Fix Office.js load/sync ordering issues.
- Replace deprecated/removed APIs with supported equivalents under WordApi 1.9.
- Improve error handling and status messages without changing behaviour.
- Update manifest URLs for hosting changes.

## Unsafe / forbidden changes (unless explicitly requested)

- Adding new functionality, new UI flows, or introducing frameworks.
- Auto-creating Content Controls.
- Changing the field detection/show-hide rules.
- Adding persistence outside the document (databases, APIs, telemetry).

## Handling "no direct replacement" API work

When asked to update Office.js API usage:

- Replace deprecated/removed functions when a direct replacement exists.
- If no direct equivalent exists:
  - remove the call/function
  - leave a comment explaining:
    - what it did
    - what functions depended on it
    - what data it returned / modified
    - any side effects

## Regression checklist

- Initialisation:
  - shows only fields that exist as Content Controls in the document
  - shows empty state when no supported tags exist
- Save:
  - updates existing Content Controls
  - updates/adds Custom Document Properties
  - locks updated Content Controls
- Reset:
  - clears visible inputs only
- External links:
  - open in system browser/tab

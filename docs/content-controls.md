# Content Controls & Custom Document Properties

## Core mapping

Each configured field `key` maps to a Word Content Control **Tag** of the same name.

Example:

- Field key: `AA_ClientName`
- Content Control Tag: `AA_ClientName`
- Custom Property Key: `AA_ClientName`

## Detection rules

- The add-in should detect whether at least one Content Control exists for each configured key.
- Only show UI fields for keys that exist in the document.
- If none exist, show empty state (no supported fields found).

## Reading values (priority)

1) **Custom Document Properties** are the primary source.
2) If a property is missing/empty, fallback to the **first matching Content Control text**.

Reason: older documents may contain controls but not yet have stored properties.

## Writing values

- Update **existing** Content Controls only.
- Update all matches for the tag (not just the first).
- Replace contents using:
  - `insertText(value, Word.InsertLocation.replace)`

## Locking behaviour

After saving values:

- Set on each relevant Content Control:
  - `cannotEdit = true`
  - `cannotDelete = true`

## Explicit constraint (do not change unless asked)

- The add-in must NOT create Content Controls automatically.
- If a key does not exist as a Content Control Tag in the document, it remains hidden and cannot be saved.

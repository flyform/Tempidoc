# Office.js / WordApi Patterns (Minimum WordApi 1.9)

This project targets **WordApi 1.9 minimum** and uses only supported Office.js patterns.

## Required pattern: Word.run

- All Word document calls must be inside:
  - `Word.run(async (context) => { ... })`
- Do not store Office objects for later use outside a `Word.run` scope.

## Required pattern: load + context.sync

Office.js uses proxied objects. You must:

1. `object.load("prop1,prop2")`
2. `await context.sync()`
3. then read `object.prop1`

## Content Controls usage

### Get by tag (collection)

- `context.document.contentControls.getByTag(tag)`
- Always load `items` before iterating:
  - `col.load("items")`
  - `await context.sync()`

### Updating text

- Replace contents with:
  - `cc.insertText(value, Word.InsertLocation.replace)`

### Locking

- After save:
  - `cc.cannotEdit = true`
  - `cc.cannotDelete = true`

## Custom Document Properties usage

- Load properties:
  - `props.load("items/key,items/value")`
- Update/add with:
  - `getItemOrNullObject(key)`
  - If null → `props.add(key, value)`
  - Else → `existing.value = value`

## Version gating (recommended)

Where practical, guard with:

- `Office.context.requirements.isSetSupported("WordApi", "1.9")`

If unsupported, show a clear status and disable actions (avoid crashing).

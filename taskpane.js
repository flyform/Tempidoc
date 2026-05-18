(() => {
  
  let CONFIG = null;
  let brandName = null;
  let DEFAULT_HELP_URL = null;
  let FIELDS = null;

  async function loadConfig() {
    const response = await fetch("config.json");
    CONFIG = await response.json();

    // Assign your variables here
    brandName = CONFIG.brandName;
    DEFAULT_HELP_URL = CONFIG.helpUrl;
    FIELDS = CONFIG.fields;
  }

  // UI element IDs/classes used by the script
  const HIDDEN_CLASS = "is-hidden";
  const statusElId = "status";
  const saveBtnId = "btnSave";
  const resetBtnId = "btnReset";
  const emptyStateId = "emptyState";
  const fieldsFormId = "fieldsForm";
  const taskPane = window.document;


  // Ensure external links open in the system browser
  function wireExternalLinks() {
    document.querySelectorAll("a[data-external]").forEach((link) => {
      link.addEventListener("click", (e) => {
        e.preventDefault();
        openExternal(link.href); // Use the openExternal utility function
      });
    });
  }


  // ---------------------------
  // Office startup
  // ---------------------------
  Office.onReady(async () => {
    await loadConfig();
    console.log("Loaded CONFIG:", CONFIG);
    // Apply CSS brand colour before page load.
    document.documentElement.style.setProperty("--brand", CONFIG.brandColor);
    
    // Update the help link URL
    const helpLink = document.getElementById("helpLink");
    if (helpLink) {
        helpLink.href = DEFAULT_HELP_URL;
    }

    wireUiEvents();
    updateBrandName();
    wireExternalLinks();

    try {
      setStatus("Loading…");
      await ensureDocumentLoaded(); // Ensure the document is fully loaded
      await initialisePane(); // Initialize the pane
      setStatus("");
    } catch (err) {
      console.error(err);
      setStatus(`Load failed: ${friendlyError(err)}`);
      // If something goes wrong, still show a helpful empty state rather than a blank form.
      showEmptyState({
        title: "Something went wrong",
        message: "We couldn't read fields from this document. Try reloading the add-in, or use Help for troubleshooting.",
        primaryAction: { label: "📘 Help", url: DEFAULT_HELP_URL }
      });
      setActionsEnabled(false);
    }
  });

  // ---------------------------
  // UI Wiring
  // ---------------------------
  function wireUiEvents() {
    const saveBtn = taskPane.getElementById(saveBtnId);
    const resetBtn = taskPane.getElementById(resetBtnId);

    if (saveBtn) saveBtn.addEventListener("click", onSave);
    if (resetBtn) resetBtn.addEventListener("click", onReset);

    // Optional UX: Enter in inputs triggers Save
    taskPane.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        const active = taskPane.activeElement;
        if (active && active.tagName === "INPUT") {
          e.preventDefault();
          onSave();
        }
      }
    });

  }

  function setStatus(message) {
    const el = taskPane.getElementById(statusElId);
    if (el) el.textContent = message || "";
  }

  function friendlyError(err) {
    return err?.message || (typeof err === "string" ? err : "Unknown error");
  }

  function getFieldContainer(key) {
    return taskPane.querySelector(`[data-field-key="${cssEscape(key)}"]`);
  }

  function getInput(key) {
    return taskPane.getElementById(key);
  }

  function hideField(key) {
    const container = getFieldContainer(key);
    if (container) container.classList.add(HIDDEN_CLASS);

    const input = getInput(key);
    if (input) input.value = "";
  }

  function showField(key) {
    const container = getFieldContainer(key);
    if (container) container.classList.remove(HIDDEN_CLASS);
  }

  function getVisibleKeys() {
    return FIELDS
      .map((f) => f.key)
      .filter((key) => {
        const c = getFieldContainer(key);
        return c && !c.classList.contains(HIDDEN_CLASS);
      });
  }

  function getValuesForKeys(keys) {
    const values = {};
    for (const key of keys) {
      const input = getInput(key);
      values[key] = (input?.value || "").trim();
    }
    return values;
  }

  function setValue(key, value) {
    const input = getInput(key);
    if (input) input.value = value ?? "";
  }

  // Simple CSS.escape substitute for tags like TS_DocTitle
  function cssEscape(s) {
    return String(s).replace(/"/g, '\\"');
  }

  function setActionsEnabled(enabled) {
    const saveBtn = taskPane.getElementById(saveBtnId);
    const resetBtn = taskPane.getElementById(resetBtnId);
    if (saveBtn) saveBtn.disabled = !enabled;
    if (resetBtn) resetBtn.disabled = !enabled;
  }

  // ---------------------------
  // Empty state (No fields found)
  // ---------------------------
  function ensureEmptyStateElement() {
    let el = taskPane.getElementById(emptyStateId);
    if (el) return el;

    // Create it dynamically if not present in HTML
    const main = taskPane.querySelector("main.container") || taskPane.body;
    el = taskPane.createElement("section");
    el.id = emptyStateId;
    el.className = `empty-state ${HIDDEN_CLASS}`;
    el.innerHTML = `
      <div class="empty-emoji">🔎</div>
      <div class="empty-title">No fields found</div>
      <div class="empty-message">
        This document doesn't contain any predefined field placeholders.
      </div>
      <div class="empty-actions">
        <button type="button" class="btn btn-secondary" id="emptyRescanBtn">🔄 Rescan</button>
        <button type="button" class="btn btn-primary" id="emptyHelpBtn">📘 Help</button>
      </div>
      <div class="empty-hint">
        Tip: Use an approved template, or ask IT to update this document's placeholders.
      </div>
    `;
    main.prepend(el);

    // Wire default buttons
    el.querySelector("#emptyRescanBtn")?.addEventListener("click", async () => {
      try {
        setStatus("Rescanning…");
        await initialisePane();
        setStatus("");
      } catch (err) {
        console.error(err);
        setStatus(`Rescan failed: ${friendlyError(err)}`);
      }
    });

    return el;
  }

  function showEmptyState(options = {}) {
    const el = ensureEmptyStateElement();

    // If caller wants custom text/actions (e.g. error case)
    const title = options.title ?? "No fields found";
    const message = options.message ??
      "This document doesn't contain any predefined field placeholders.";
    const primaryAction = options.primaryAction ?? { label: "📘 Help", url: DEFAULT_HELP_URL };

    // Update content if elements exist (created by ensureEmptyStateElement)
    const titleEl = el.querySelector(".empty-title");
    const msgEl = el.querySelector(".empty-message");
    if (titleEl) titleEl.textContent = title;
    if (msgEl) msgEl.textContent = message;

    const helpBtn = el.querySelector("#emptyHelpBtn");
    if (helpBtn) {
      helpBtn.textContent = primaryAction.label ?? "📘 Help";
      helpBtn.onclick = () => openExternal(primaryAction.url ?? DEFAULT_HELP_URL);
    }

    el.classList.remove(HIDDEN_CLASS);

    // Hide main form (optional – keeps UI clean)
    const form = taskPane.getElementById(fieldsFormId);
    if (form) {
      console.log("Hiding fields form.");
      form.classList.add(HIDDEN_CLASS); // Ensure the form is hidden
    }
  }

  function hideEmptyState() {
    const el = taskPane.getElementById(emptyStateId);
    if (el) {
      console.log("Hiding empty state.");
      el.classList.add(HIDDEN_CLASS);
    }

    const form = taskPane.getElementById(fieldsFormId);
    if (form) {
      form.classList.remove(HIDDEN_CLASS); // Ensure the form is visible
    }
  }

  // ---------------------------
  // Initialisation logic
  // ---------------------------
  function generateFields() {
    const form = taskPane.getElementById(fieldsFormId);
    if (!form) return;

    // Clear existing fields
    form.querySelectorAll(".field").forEach((field) => field.remove());

    // Generate fields dynamically from FIELDS
    for (const field of FIELDS) {
      const container = taskPane.createElement("div");
      container.className = "field";
      container.setAttribute("data-field-key", field.key);

      const label = taskPane.createElement("label");
      label.className = "field-label";
      label.setAttribute("for", field.key);
      label.textContent = field.label;

      let input;
      if (field.type === "string") {
        input = taskPane.createElement("input");
        input.className = "field-input";
        input.setAttribute("type", "text");
      } else if (field.type === "dropdown" && field.choices) {
        input = taskPane.createElement("select");
        input.className = "field-input";
        for (const choice of field.choices) {
          const option = taskPane.createElement("option");
          option.value = choice;
          option.textContent = choice;
          input.appendChild(option);
        }
      }

      if (input) {
        input.id = field.key;
        input.name = field.key;
        container.appendChild(label);
        container.appendChild(input);
        form.insertBefore(container, form.querySelector(".actions"));
      }
    }
  }

  // ---------------------------
  // Replace placeholder brand name in the UI with the actual brand name from config
  // ---------------------------
  function updateBrandName() {
    const brandElements = document.querySelectorAll(".brandName");
    brandElements.forEach((el) => {
      el.textContent = brandName;
    });
  }

  async function initialisePane() {
    const existingTags = await detectExistingTags();
    generateFields();
    for (const f of FIELDS) {
      if (existingTags.has(f.key)) {
        console.log(`Showing field for tag: ${f.key}`);
        showField(f.key);
      } else {
        console.log(`Hiding field for tag: ${f.key}`);
        hideField(f.key);
      }
    }

    const visibleKeys = getVisibleKeys();
    if (visibleKeys.length === 0) {
      showEmptyState({
        title: "No fields found",
        message: "This document doesn't contain any predefined field placeholders, so there's nothing to populate.",
        primaryAction: { label: "📘 Help", url: DEFAULT_HELP_URL }
      });
      setActionsEnabled(false);
    } else {
      hideEmptyState();
      setActionsEnabled(true);

      const propValues = await readCustomProperties(visibleKeys);
      for (const key of visibleKeys) {
        if (propValues[key] !== undefined && propValues[key] !== null) {
          setValue(key, propValues[key]);
        }
      }

      const missingKeys = visibleKeys.filter((k) => propValues[k] === undefined);
      if (missingKeys.length > 0) {
        const ccValues = await readFirstContentControlText(missingKeys);
        for (const [k, v] of Object.entries(ccValues)) {
          if (v !== undefined && v !== null) setValue(k, v);
        }
      }

      await lockContentControls(Array.from(existingTags));
    }
  }

  /**
   * Returns Set of keys for which at least one content control with that tag exists.
   */
async function detectExistingTags() {
  return Word.run(async (context) => {
    const existing = new Set();
    const detectedTags = new Set();

    const controls = context.document.getContentControls({
      types: [Word.ContentControlType.plainText]
    });
    controls.load("items/tag,items/type");
    await context.sync();

    for (const cc of controls.items) {
      if (
        cc.type === Word.ContentControlType.plainText &&
        typeof cc.tag === "string" &&
        cc.tag.trim() !== ""
      ) {
        detectedTags.add(cc.tag.trim());
      }
    }

    for (const f of FIELDS) {
      if (detectedTags.has(f.key)) {
        existing.add(f.key);
      }
    }

    return existing;
  });
}

  /**
   * Reads custom document properties for the given keys.
   * Returns a dictionary { key: value } for those found.
   */
async function readCustomProperties(keys) {
  if (!keys || keys.length === 0) return {};

  return Word.run(async (context) => {
    const result = {};
    const props = context.document.properties.customProperties;
    props.load("items/key,items/value");
    await context.sync();

    // Build a simple lookup from what's loaded
    const propMap = {};
    for (const prop of props.items) {
      propMap[prop.key] = prop.value;
    }

    for (const key of keys) {
      if (propMap[key] !== undefined) {
        result[key] = propMap[key];
      }
    }

    return result;
  });
}


  /**
   * Reads the displayed text of the first matching content control for each key.
   * Useful when opening older docs that never had properties written.
   *
   * WordApi 1.9: replaced the previous approach of calling getRange() on the content
   * control and then loading range.text (which required an extra context.sync per key)
   * with directly loading the "text" property on the ContentControl object itself.
   * This reduces round-trips and aligns with the modern API surface.
   *
   * Previous interconnect: readFirstContentControlText was called from initialisePane()
   * only for keys where readCustomProperties() returned no value (missingKeys). That
   * call site is unchanged; only the internal implementation has been updated.
   */
  async function readFirstContentControlText(keys) {
    if (!keys || keys.length === 0) return {};

    return Word.run(async (context) => {
      const result = {};

      const controls = context.document.getContentControls({
        types: [Word.ContentControlType.plainText]
      });
      controls.load("items/tag,items/type,items/text");
      await context.sync();

      for (const key of keys) {
        const match = controls.items.find(
          cc => cc.tag?.trim() === key &&
                cc.type === Word.ContentControlType.plainText &&
                typeof cc.text === "string"
        );
        if (match) {
          result[key] = match.text.trim();
        }
      }

      return result;
    });
  }



  /**
   * Locks all content controls for given tags so users cannot edit/delete manually.
   */
  async function lockContentControls(tags) {
    if (!tags || tags.length === 0) return;

    return Word.run(async (context) => {
      const controls = context.document.getContentControls({
        types: [Word.ContentControlType.plainText]
      });
      controls.load("items/tag,items/cannotDelete");
      await context.sync();

      for (const cc of controls.items) {
        if (tags.includes(cc.tag?.trim())) {
          cc.cannotEdit = true;

          // Only set cannotDelete to true if it is not already false
          if (cc.cannotDelete !== false) {
            cc.cannotDelete = true;
          }
        }
      }
      await context.sync();
    });
  }

  // ---------------------------
  // Save / Reset
  // ---------------------------
  async function onSave() {
    try {
      setStatus("Saving…");

      const visibleKeys = getVisibleKeys();
      if (visibleKeys.length === 0) {
        setStatus("Nothing to save.");
        return;
      }

      const values = getValuesForKeys(visibleKeys);
      await writeValues(values);

      // Lock content controls after saving values
      await lockContentControls(visibleKeys);

      setStatus("Saved ✅");
    } catch (err) {
      console.error(err);
      setStatus(`Save failed: ${friendlyError(err)}`);
    }
  }

  function onReset() {
    const visibleKeys = getVisibleKeys();
    for (const key of visibleKeys) {
      const input = getInput(key);
      if (input) input.value = "";
    }
    setStatus("Reset.");
  }

  /**
   * Updates EXISTING content controls and creates/updates custom doc properties.
   * Does NOT create content controls (per your requirement).
   */
  async function writeValues(valuesByKey) {
    const keys = Object.keys(valuesByKey);
    if (keys.length === 0) return;

    return Word.run(async (context) => {
      const controls = context.document.getContentControls({
        types: [Word.ContentControlType.plainText]
      });
      controls.load("items/tag");
      await context.sync();

      const props = context.document.properties.customProperties;

      for (const key of keys) {
        const value = (valuesByKey[key] ?? "").trim();

        // Find all controls matching this tag
        const matches = controls.items.filter(cc => cc.tag?.trim() === key);

        if (matches.length === 0) continue;

        for (const cc of matches) {
          cc.cannotEdit = false; // Temporarily unlock for editing
          cc.insertText(value, "Replace");
          cc.cannotEdit = true; // Lock editing again

          // Set cannotDelete based on whether the value is empty
          cc.cannotDelete = value !== ""; // Allow deletion if value is empty
        }

        // Create or update the custom property
        const existingProp = props.getItemOrNullObject(key);
        existingProp.load("key");
        await context.sync();

        if (existingProp.isNullObject) {
          props.add(key, value);
        } else {
          existingProp.value = value;
        }

        await context.sync();
      }
    });
  }

  // ---------------------------
  // External links
  // ---------------------------
  function openExternal(url) {
    try {
      if (Office?.context?.requirements?.isSetSupported?.("OpenBrowserWindowApi", "1.1")) {
        Office.context.ui.openBrowserWindow(url);
        return;
      }
    } catch (_) {
      // ignore and fall back
    }
    window.open(url, "_blank", "noreferrer");
  }

  /**
   * Waits for the Word document to be accessible before proceeding.
   *
   * WordApi 1.9: replaced the previous body.load() call (no-argument form, which
   * loaded all properties and was discouraged as of later API versions) with
   * body.load("text"), explicitly requesting only the property needed to confirm
   * the document body is readable. This follows the WordApi best practice of always
   * specifying which properties to load rather than relying on the implicit all-load
   * behaviour of the no-argument form.
   */
  async function ensureDocumentLoaded(retries = 3, delay = 1000) {
    for (let i = 0; i < retries; i++) {
      try {
        await Word.run(async (context) => {
          const body = context.document.body;
          // Explicitly load a single lightweight property ("text") rather than the
          // deprecated no-argument body.load() which fetched all properties.
          body.load("text");
          await context.sync();
        });
        return true; // Exit if successful
      } catch (err) {
        console.warn(`Document load attempt ${i + 1} failed. Retrying...`);
        await new Promise((resolve) => setTimeout(resolve, delay)); // Wait before retrying
      }
    }
    throw new Error("Document failed to load after multiple attempts.");
  }
})();

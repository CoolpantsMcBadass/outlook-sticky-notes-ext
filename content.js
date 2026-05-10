(() => {
  let currentKey = null;
  let collapsed = false;
  let lastDeleted = null;  // for undo — single-level only; index position is best-effort
  let dragSrcId = null;    // for drag-and-drop reorder
  let cachedIsCompose = false; // cached per URL to avoid repeated querySelector in observer
  let injecting = false;   // guard against duplicate injection on rapid observer fire
  let undoTimer = null;    // tracks the auto-dismiss timeout for the undo bar

  // Matches Outlook's /id/<msgId> and /read/<msgId> URL patterns.
  // Shared between getKey() and updateCurrentListBadge() to avoid duplication.
  const MESSAGE_ID_RE = /\/(?:id|read)\/([^/?#]+)/i;

  // Storage key for the note-count index — a single {noteKey: count} map that
  // lets us badge email list rows without reading every note array individually.
  const INDEX_KEY = "osn_index";

  // All timing and sizing constants in one place.
  const OSN = {
    MAX_NOTE_LENGTH:   500,
    SUBJECT_KEY_MAX:   80,
    CHAR_WARN_AT:      100,  // chars remaining at which the counter appears
    CHAR_URGENT_AT:    20,   // chars remaining at which the counter turns red
    BODY_MAX_HEIGHT:   220,  // px
    POPOUT_MAX_HEIGHT: 160,  // px
    RESIZE_MIN_HEIGHT: 60,   // px
    STORAGE_WARN_PCT:  0.80,
    MAX_RETRIES:       20,
    RETRY_DELAY_MS:    400,
    NAV_INJECT_DELAY:  600,  // ms to wait after URL change before injecting (let Outlook settle)
    BADGE_DELAY_MS:    100,
    BADGE_NAV_DELAY:   1200,
    BADGE_UPDATE_MS:   600,
    POPOUT_ALIGN_MS:   600,  // ms to wait before measuring DOM positions in pop-out
    UNDO_TIMEOUT_MS:   7000,
  };

  function isExtensionAlive() {
    try {
      // Accessing chrome.runtime.id throws if the extension context has been invalidated
      // (e.g. after an update). Guards every storage call to prevent unhandled errors.
      return !!chrome.runtime?.id;
    } catch {
      return false;
    }
  }

  // ── Compose view detection ──────────────────────────────────────────────────
  function isComposeView() {
    // URL-based check first — fastest path
    if (/\/(compose|action\/compose|deeplink\/compose)/i.test(location.href)) return true;
    // DOM-based fallback: presence of both a Send button and an editable To field
    // distinguishes compose from read even when the URL doesn't include "compose"
    const sendBtn = document.querySelector(
      '[aria-label="Send"], [title="Send"], button[data-testid*="send"], [aria-label*="Send "]'
    );
    const toField = document.querySelector('input[aria-label="To"], div[aria-label="To"][contenteditable="true"]');
    if (sendBtn && toField) return true;
    return false;
  }

  // ── Pop-out detection ──────────────────────────────────────────────────────
  function isPopout() {
    return !!document.getElementById("_owa_projection_root");
  }

  // ── Key extraction ──────────────────────────────────────────────────────────
  // Derives a stable per-thread storage key from the page URL or DOM.
  // Returns null if no ID can be found yet — callers should retry.
  function getKey() {
    const url = location.href;
    const match = url.match(MESSAGE_ID_RE);
    if (match) return "osn_" + decodeURIComponent(match[1]);

    // Pop-out windows open on about:blank, so their URL never contains a message ID.
    // The opener is the main Outlook tab (always same-origin), so we can read its URL
    // to derive the same key the main window uses. The try/catch is a defensive
    // cross-origin guard — opener should always be same-origin here, but if somehow
    // it isn't (e.g. a future Outlook architecture change), we fail silently rather
    // than throw a cross-origin security error.
    if (isPopout() && window.opener) {
      try {
        const openerUrl = window.opener.location.href;
        const openerMatch = openerUrl.match(MESSAGE_ID_RE);
        if (openerMatch) return "osn_" + decodeURIComponent(openerMatch[1]);
      } catch { /* cross-origin guard */ }
    }

    // Fallback: data-convid attribute Outlook stamps on conversation rows
    const convEl = document.querySelector('[data-convid]');
    if (convEl) return "osn_" + convEl.getAttribute('data-convid');

    // Last resort: subject heading text. Collision risk — threads sharing the same
    // subject will share notes. Only used when no ID is available at all.
    const subj = document.querySelector('[role="heading"][aria-level="2"], [role="heading"][aria-level="1"], [role="heading"][aria-level="3"]');
    if (subj && subj.textContent.trim()) {
      return "osn_subj_" + subj.textContent.trim().slice(0, OSN.SUBJECT_KEY_MAX);
    }
    return null; // no usable key — caller will retry
  }

  // ── Key migration ───────────────────────────────────────────────────────────
  // Earlier versions stored keys with URL-encoded IDs (osn_AAQ...%3D).
  // Now keys use decoded IDs (osn_AAQ...=) to match data-convid attributes.
  // Guarded by osn_migrated_v1 flag so this runs only once per installation.
  function migrateEncodedKeys() {
    if (!isExtensionAlive()) return;
    try {
      chrome.storage.local.get("osn_migrated_v1", (r) => {
        if (r.osn_migrated_v1) return; // already done
        chrome.storage.local.get(null, (allData) => {
          const toSet = { osn_migrated_v1: true };
          const toRemove = [];
          Object.entries(allData).forEach(([k, v]) => {
            if (!k.startsWith("osn_") || !k.includes("%")) return;
            const decodedKey = "osn_" + decodeURIComponent(k.slice(4));
            // Migrate notes to decoded key if slot is empty; otherwise just drop the orphan
            if (!(decodedKey in allData) && Array.isArray(v) && v.length > 0) {
              toSet[decodedKey] = v;
            }
            toRemove.push(k);
          });
          if (toRemove.length) chrome.storage.local.remove(toRemove, () => {
            if (chrome.runtime.lastError) console.warn("osn migrate remove:", chrome.runtime.lastError);
          });
          chrome.storage.local.set(toSet, () => {
            if (chrome.runtime.lastError) console.warn("osn migrate set:", chrome.runtime.lastError);
          });
        });
      });
    } catch { /* context invalidated */ }
  }

  // ── Storage helpers ─────────────────────────────────────────────────────────
  function loadNotes(key) {
    return new Promise((res) => {
      if (!isExtensionAlive()) return res([]);
      try {
        chrome.storage.local.get(key, (data) => res(data[key] || []));
      } catch {
        res([]);
      }
    });
  }

  function saveNotes(key, notes) {
    if (!isExtensionAlive()) return;
    try {
      chrome.storage.local.set({ [key]: notes }, () => {
        if (chrome.runtime.lastError) console.warn("osn save:", chrome.runtime.lastError);
      });
    } catch { /* context invalidated — ignore */ }
  }

  // Index helpers — maintain a single {noteKey: count} map so badge updates
  // can read one storage entry instead of scanning the entire storage space.
  function loadIndex() {
    return new Promise((res) => {
      if (!isExtensionAlive()) return res({});
      try {
        chrome.storage.local.get(INDEX_KEY, (data) => res(data[INDEX_KEY] || {}));
      } catch { res({}); }
    });
  }

  function saveIndex(index) {
    if (!isExtensionAlive()) return;
    try {
      chrome.storage.local.set({ [INDEX_KEY]: index }, () => {
        if (chrome.runtime.lastError) console.warn("osn index:", chrome.runtime.lastError);
      });
    } catch {}
  }

  async function updateIndex(key, count) {
    const index = await loadIndex();
    if (count > 0) index[key] = count;
    else delete index[key];
    saveIndex(index);
  }

  // Build index from existing storage — one-time bootstrap for installations that
  // pre-date the index. Skipped if the index already has entries.
  function bootstrapIndex() {
    if (!isExtensionAlive()) return;
    try {
      chrome.storage.local.get(INDEX_KEY, (r) => {
        const existing = r[INDEX_KEY];
        if (existing && Object.keys(existing).length > 0) return;
        chrome.storage.local.get(null, (allData) => {
          const index = {};
          Object.entries(allData).forEach(([k, v]) => {
            if (k.startsWith("osn_") && k !== INDEX_KEY && !k.startsWith("osn_migrated") && Array.isArray(v) && v.length > 0) {
              index[k] = v.length;
            }
          });
          if (Object.keys(index).length > 0) saveIndex(index);
        });
      });
    } catch {}
  }

  // ── Panel HTML ──────────────────────────────────────────────────────────────
  const POSTIT_URL = chrome.runtime.getURL("icons/postit.png");

  function buildPanel() {
    const panel = document.createElement("div");
    panel.id = "osn-panel";
    panel.style.setProperty("--osn-postit-url", `url("${POSTIT_URL}")`);
    // Landmark role lets screen readers jump directly to the sticky notes panel
    panel.setAttribute("role", "complementary");
    panel.setAttribute("aria-label", "Sticky Notes");

    panel.innerHTML = `
      <div id="osn-header" tabindex="0" aria-expanded="false" data-tooltip="Click to collapse">
        <span id="osn-title"><span id="osn-title-text">Sticky Notes</span></span>
        <span id="osn-collapsed-plus" aria-hidden="true">+</span>
        <button id="osn-btn-add" aria-label="Add new note" data-tooltip="New note">+</button>
      </div>
      <div id="osn-body" role="list">
        <span id="osn-empty" role="listitem">No notes yet for this thread.</span>
      </div>
      <div id="osn-undo-bar" class="osn-hidden">
        Note deleted — <button id="osn-undo-btn">Undo</button>
      </div>
      <div id="osn-storage-warning" class="osn-hidden"></div>
      <div id="osn-input-area" class="osn-hidden">
        <textarea id="osn-textarea" maxlength="${OSN.MAX_NOTE_LENGTH}" placeholder="Type your notes and click save"></textarea>
        <div id="osn-char-counter" class="osn-hidden"></div>
        <div id="osn-input-btns">
          <button id="osn-save">Save</button>
          <button id="osn-cancel">Cancel</button>
        </div>
      </div>
    `;

    panel.addEventListener("mousedown", (e) => e.stopPropagation());

    // ── Pop-out resize handle ──
    const resizeHandle = document.createElement("div");
    resizeHandle.id = "osn-resize-handle";
    resizeHandle.textContent = "⋯";
    panel.appendChild(resizeHandle);

    resizeHandle.addEventListener("mousedown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const body = document.getElementById("osn-body");
      const startY = e.clientY;
      const startHeight = body.getBoundingClientRect().height;

      const onMove = (me) => {
        const newHeight = Math.max(OSN.RESIZE_MIN_HEIGHT, startHeight + (me.clientY - startY));
        body.style.maxHeight = newHeight + "px";
      };
      const onUp = () => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
      };
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
    });

    // + button: open input without collapsing (only visible when expanded)
    panel.querySelector("#osn-btn-add").addEventListener("click", (e) => {
      e.stopPropagation();
      showInput();
    });

    // Header click: expand (+ open input) when collapsed; collapse when expanded
    panel.querySelector("#osn-header").addEventListener("click", (e) => {
      e.stopPropagation();
      if (e.target.id === "osn-btn-add") return; // handled above
      if (collapsed) {
        toggleCollapse();
        const hasNotes = document.querySelectorAll(".osn-note").length > 0;
        if (!hasNotes) showInput();
      } else {
        toggleCollapse();
        hideInput();
      }
    });

    // Keyboard collapse/expand — Enter or Space on the header itself (not a child button)
    panel.querySelector("#osn-header").addEventListener("keydown", (e) => {
      if (e.key !== "Enter" && e.key !== " ") return;
      if (e.target !== e.currentTarget) return; // child button handles its own keys
      e.preventDefault();
      if (collapsed) {
        toggleCollapse();
        const hasNotes = document.querySelectorAll(".osn-note").length > 0;
        if (!hasNotes) showInput();
      } else {
        toggleCollapse();
        hideInput();
      }
    });

    // Body click when expanded opens note input
    panel.querySelector("#osn-body").addEventListener("click", (e) => {
      e.stopPropagation();
      if (!collapsed) showInput();
    });

    // Input area stays self-contained — clicks here don't bubble to header/body
    panel.querySelector("#osn-input-area").addEventListener("click", (e) => e.stopPropagation());

    panel.querySelector("#osn-save").addEventListener("click", (e) => { e.stopPropagation(); saveNote(); });
    panel.querySelector("#osn-cancel").addEventListener("click", (e) => { e.stopPropagation(); hideInput(); });
    panel.querySelector("#osn-undo-btn").addEventListener("click", (e) => { e.stopPropagation(); undoDelete(); });
    panel.querySelector("#osn-textarea").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) saveNote();
      if (e.key === "Escape") hideInput();
    });
    panel.querySelector("#osn-textarea").addEventListener("input", () => {
      const ta = document.getElementById("osn-textarea");
      const counter = document.getElementById("osn-char-counter");
      if (!counter || !ta) return;
      const len = ta.value.length;
      const remaining = OSN.MAX_NOTE_LENGTH - len;
      if (remaining <= OSN.CHAR_WARN_AT) {
        counter.textContent = `${len} / ${OSN.MAX_NOTE_LENGTH}`;
        counter.className = remaining <= OSN.CHAR_URGENT_AT ? "osn-char-urgent" : "";
        counter.classList.remove("osn-hidden");
      } else {
        counter.className = "osn-hidden";
      }
    });

    return panel;
  }

  function showInput() {
    document.getElementById("osn-input-area")?.classList.remove("osn-hidden");
    document.getElementById("osn-textarea")?.focus();
  }

  function hideInput() {
    document.getElementById("osn-input-area")?.classList.add("osn-hidden");
    const ta = document.getElementById("osn-textarea");
    if (ta) ta.value = "";
    const counter = document.getElementById("osn-char-counter");
    if (counter) counter.className = "osn-hidden";
  }

  function toggleCollapse() {
    collapsed = !collapsed;
    const panel = document.getElementById("osn-panel");
    const body = document.getElementById("osn-body");
    const input = document.getElementById("osn-input-area");
    const addBtn = document.getElementById("osn-btn-add");
    // Check the panel's own class rather than calling isPopout() — when
    // match_origin_as_fallback causes the script to run in a sub-frame,
    // _owa_projection_root may not be present in that frame's document.
    const panelIsPopout = panel?.classList.contains("osn-popout");
    document.getElementById("osn-header")?.setAttribute("aria-expanded", collapsed ? "false" : "true");
    if (collapsed) {
      panel?.classList.add("osn-collapsed");
      if (panelIsPopout) {
        const root = document.getElementById("_owa_projection_root");
        const heading = document.querySelector('[id$="_SUBJECT"], [role="heading"][aria-level="3"]');
        if (root && heading) {
          const rootRect = root.getBoundingClientRect();
          const headingRect = heading.getBoundingClientRect();
          panel.style.top = (headingRect.top - rootRect.top - 6) + "px";
        }
      }
      body?.classList.add("osn-hidden");
      input?.classList.add("osn-hidden");
      addBtn?.classList.add("osn-hidden");
      clearUndo();
    } else {
      panel?.classList.remove("osn-collapsed");
      if (panelIsPopout) panel.style.top = "";
      body?.classList.remove("osn-hidden");
      input?.classList.add("osn-hidden"); // stay hidden until user opens it
      const hasNotes = document.querySelectorAll(".osn-note").length > 0;
      if (hasNotes) addBtn?.classList.remove("osn-hidden");
    }
  }

  async function saveNote() {
    if (!currentKey) return;
    const ta = document.getElementById("osn-textarea");
    const text = ta?.value.trim();
    if (!text) return;

    const notes = await loadNotes(currentKey);
    notes.unshift({ id: crypto.randomUUID(), text, date: new Date().toISOString() });
    saveNotes(currentKey, notes);
    await updateIndex(currentKey, notes.length);
    ta.value = "";
    hideInput();
    renderNotes(notes);
    clearUndo();
    checkStorageQuota();
    setTimeout(() => updateCurrentListBadge(notes.length), OSN.BADGE_UPDATE_MS);
  }

  async function deleteNote(id) {
    const allNotes = await loadNotes(currentKey);
    const idx = allNotes.findIndex((n) => n.id === id);
    if (idx === -1) return;

    // Store deleted note so undoDelete() can restore it
    lastDeleted = { note: allNotes[idx], index: idx };
    document.getElementById("osn-undo-bar")?.classList.remove("osn-hidden");
    // Auto-dismiss the undo bar after a timeout so it doesn't linger indefinitely
    clearTimeout(undoTimer);
    undoTimer = setTimeout(() => clearUndo(), OSN.UNDO_TIMEOUT_MS);

    const remaining = allNotes.filter((n) => n.id !== id);
    saveNotes(currentKey, remaining);
    await updateIndex(currentKey, remaining.length);
    renderNotes(remaining);
    setTimeout(() => updateCurrentListBadge(remaining.length), OSN.BADGE_UPDATE_MS);

    // Auto-collapse when the last note is deleted
    if (remaining.length === 0 && !collapsed) {
      toggleCollapse();
      hideInput();
    }
  }

  async function undoDelete() {
    if (!lastDeleted) return;
    const { note, index } = lastDeleted;
    const notes = await loadNotes(currentKey);
    // Clamp the insertion index in case notes were added or reordered between
    // the delete and the undo (e.g. two tabs open to the same thread)
    const safeIndex = Math.min(index, notes.length);
    notes.splice(safeIndex, 0, note);
    saveNotes(currentKey, notes);
    await updateIndex(currentKey, notes.length);
    lastDeleted = null;
    // Re-expand if the panel auto-collapsed when the last note was deleted
    if (collapsed) toggleCollapse();
    renderNotes(notes);
    clearUndo();
    setTimeout(() => updateCurrentListBadge(notes.length), OSN.BADGE_UPDATE_MS);
  }

  function clearUndo() {
    clearTimeout(undoTimer);
    undoTimer = null;
    lastDeleted = null;
    document.getElementById("osn-undo-bar")?.classList.add("osn-hidden");
  }

  function checkStorageQuota() {
    if (!isExtensionAlive()) return;
    try {
      // Pass null to measure total bytes used across all keys — QUOTA_BYTES is a
      // global limit, not per-key, so we need the global total to compute percentage.
      chrome.storage.local.getBytesInUse(null, (bytes) => {
        const QUOTA = chrome.storage.local.QUOTA_BYTES;
        const pct = bytes / QUOTA;
        const warning = document.getElementById("osn-storage-warning");
        if (!warning) return;
        if (pct > OSN.STORAGE_WARN_PCT) {
          warning.textContent = `⚠️ Notes storage ${Math.round(pct * 100)}% full — consider deleting old notes.`;
          warning.classList.remove("osn-hidden");
        } else {
          warning.classList.add("osn-hidden");
        }
      });
    } catch { /* context invalidated */ }
  }

  async function startEditNote(div, note) {
    div.draggable = false;
    // note.text is run through escapeHtml before insertion into innerHTML
    div.innerHTML = `
      <div class="osn-edit-area">
        <textarea class="osn-edit-textarea" maxlength="${OSN.MAX_NOTE_LENGTH}">${escapeHtml(note.text)}</textarea>
        <div class="osn-edit-footer">
          <span class="osn-edit-char-count">${note.text.length} / ${OSN.MAX_NOTE_LENGTH}</span>
          <div class="osn-edit-btns">
            <button class="osn-edit-save">Save</button>
            <button class="osn-edit-cancel">Cancel</button>
          </div>
        </div>
      </div>
    `;
    div.addEventListener("click", (e) => e.stopPropagation());

    const ta = div.querySelector(".osn-edit-textarea");
    ta.focus();
    ta.setSelectionRange(ta.value.length, ta.value.length);

    ta.addEventListener("input", () => {
      div.querySelector(".osn-edit-char-count").textContent = `${ta.value.length} / ${OSN.MAX_NOTE_LENGTH}`;
    });
    ta.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) commitEdit();
      if (e.key === "Escape") cancelEdit();
    });
    div.querySelector(".osn-edit-save").addEventListener("click", (e) => { e.stopPropagation(); commitEdit(); });
    div.querySelector(".osn-edit-cancel").addEventListener("click", (e) => { e.stopPropagation(); cancelEdit(); });

    async function commitEdit() {
      const newText = ta.value.trim();
      if (!newText) return;
      const notes = await loadNotes(currentKey);
      const idx = notes.findIndex((n) => n.id === note.id);
      if (idx !== -1) {
        notes[idx] = { ...notes[idx], text: newText, date: new Date().toISOString() };
        saveNotes(currentKey, notes);
        renderNotes(notes);
      }
    }

    async function cancelEdit() {
      const notes = await loadNotes(currentKey);
      renderNotes(notes);
    }
  }

  // Escapes user-supplied text before writing to innerHTML.
  // All note content goes through this before any DOM insertion.
  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  // Render a stored ISO date string in a consistent, locale-independent short format.
  // Falls back gracefully for old notes that stored toLocaleString() values.
  function formatDate(dateStr) {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr; // legacy locale string — display as-is
    return d.toLocaleString(undefined, {
      month: "short", day: "numeric", year: "numeric",
      hour: "numeric", minute: "2-digit",
    });
  }

  // Shows note count on the collapsed icon (or + when there are no notes)
  function updateCollapsedPlus(noteCount) {
    const plus = document.getElementById("osn-collapsed-plus");
    if (!plus) return;
    plus.textContent = noteCount === 0 ? "+" : String(noteCount);
  }

  function renderNotes(notes) {
    const body = document.getElementById("osn-body");
    if (!body) return;

    const addBtn = document.getElementById("osn-btn-add");
    if (addBtn) addBtn.classList.toggle("osn-hidden", notes.length === 0);
    updateCollapsedPlus(notes.length);

    body.innerHTML = "";

    if (!notes.length) {
      body.innerHTML = '<span id="osn-empty" role="listitem">No notes yet for this thread.</span>';
      return;
    }

    notes.forEach((note) => {
      const div = document.createElement("div");
      div.className = "osn-note";
      div.draggable = true;
      div.dataset.id = String(note.id);
      // tabIndex and role make each note keyboard-reachable and screen-reader accessible
      div.tabIndex = 0;
      div.setAttribute("role", "listitem");
      div.setAttribute("aria-label", `Note: ${note.text.slice(0, 40)}`);

      // note.text and formatDate output are both escaped before innerHTML insertion
      div.innerHTML = `
        <span class="osn-drag-handle" aria-hidden="true" title="Drag to reorder">⠿</span>
        <span class="osn-note-text">${escapeHtml(note.text)}</span>
        <div class="osn-note-meta">
          <button class="osn-note-edit" aria-label="Edit note" title="Edit">✎</button>
          <button class="osn-note-delete" aria-label="Delete note" title="Delete">✕</button>
          <span class="osn-note-date">${escapeHtml(formatDate(note.date))}</span>
        </div>
      `;

      // Arrow Up/Down reorders the focused note within the list
      div.addEventListener("keydown", async (e) => {
        if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
        e.preventDefault();
        const fresh = await loadNotes(currentKey);
        const idx = fresh.findIndex((n) => n.id === note.id);
        if (idx === -1) return;
        const swap = e.key === "ArrowUp" ? idx - 1 : idx + 1;
        if (swap < 0 || swap >= fresh.length) return;
        [fresh[idx], fresh[swap]] = [fresh[swap], fresh[idx]];
        saveNotes(currentKey, fresh);
        renderNotes(fresh);
        // Restore focus to the moved note after re-render
        setTimeout(() => {
          document.querySelectorAll(".osn-note")[swap]?.focus();
        }, 0);
      });

      // ── Drag-and-drop reordering ──
      div.addEventListener("dragstart", (e) => {
        dragSrcId = note.id;
        e.dataTransfer.effectAllowed = "move";
        setTimeout(() => div.classList.add("osn-dragging"), 0);
      });
      div.addEventListener("dragend", () => {
        div.classList.remove("osn-dragging");
        body.querySelectorAll(".osn-drag-over").forEach((el) => el.classList.remove("osn-drag-over"));
      });
      div.addEventListener("dragover", (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        if (dragSrcId !== note.id) div.classList.add("osn-drag-over");
      });
      div.addEventListener("dragleave", () => div.classList.remove("osn-drag-over"));
      div.addEventListener("drop", async (e) => {
        e.preventDefault();
        e.stopPropagation();
        div.classList.remove("osn-drag-over");
        if (dragSrcId === note.id) return;
        const fresh = await loadNotes(currentKey);
        const from = fresh.findIndex((n) => n.id === dragSrcId);
        const to = fresh.findIndex((n) => n.id === note.id);
        if (from === -1 || to === -1) return;
        const [moved] = fresh.splice(from, 1);
        fresh.splice(to, 0, moved);
        saveNotes(currentKey, fresh);
        renderNotes(fresh);
      });

      // Stop drag handle from triggering body's "open input" click
      div.querySelector(".osn-drag-handle").addEventListener("click", (e) => e.stopPropagation());

      div.querySelector(".osn-note-edit").addEventListener("click", (e) => {
        e.stopPropagation();
        startEditNote(div, note);
      });
      div.querySelector(".osn-note-delete").addEventListener("click", (e) => {
        e.stopPropagation();
        deleteNote(note.id);
      });

      body.appendChild(div);
    });
  }

  // ── List item badges ────────────────────────────────────────────────────────
  function setBadgeOnElement(el, noteCount) {
    el.querySelector(".osn-list-badge")?.remove();
    if (noteCount > 0) {
      const badge = document.createElement("img");
      badge.className = "osn-list-badge";
      badge.src = POSTIT_URL;
      badge.title = `${noteCount} sticky note${noteCount !== 1 ? "s" : ""}`;
      badge.alt = "";
      el.appendChild(badge);
    }
  }

  // Badge the currently selected list item after a save or delete.
  // Reads the message ID from the URL and matches it against [data-convid] in the list.
  function updateCurrentListBadge(noteCount) {
    const urlId = location.href.match(MESSAGE_ID_RE)?.[1];
    if (!urlId) return;
    const decoded = decodeURIComponent(urlId);
    const el = document.querySelector(`[data-convid="${CSS.escape(decoded)}"]`);
    if (el) setBadgeOnElement(el, noteCount);
  }

  // Scan all visible conversation rows and badge any that have stored notes.
  // Reads from the count index rather than fetching every note array individually.
  async function updateAllListBadges() {
    if (!isExtensionAlive()) return;
    const index = await loadIndex();
    if (Object.keys(index).length === 0) return;
    const seen = new Set();
    document.querySelectorAll('[data-convid]').forEach((el) => {
      const val = el.getAttribute('data-convid');
      if (seen.has(val)) return; // skip nested duplicates, badge outermost only
      seen.add(val);
      const rawKey = "osn_" + val;
      // Check both decoded and encoded forms to handle any mixed legacy data
      const count = index[rawKey] ?? index["osn_" + encodeURIComponent(val)] ?? 0;
      if (count > 0) setBadgeOnElement(el, count);
    });
  }

  // ── Injection ───────────────────────────────────────────────────────────────
  // Find the stable conversation container and inject before the scroll area.
  function findInsertionPoint() {
    // Primary: stable ID present in new Outlook (outlook.cloud.microsoft)
    const container = document.querySelector("#ConversationReadingPaneContainer");
    if (container) {
      // .L72vd is Outlook's minified class name for the email scroll container.
      // We insert above it so the panel sits above the message thread, not inside it.
      // If Outlook's build changes this name the panel falls back to prepending to the
      // container itself, which still works — it just loses precise positioning.
      const scrollArea = container.querySelector(".L72vd");
      return { parent: container, before: scrollArea || null };
    }

    // Pop-out window: Outlook renders into #_owa_projection_root
    const projection = document.getElementById("_owa_projection_root");
    if (projection) {
      return { parent: projection, before: projection.firstElementChild || null };
    }

    // Fallback A: data attribute on conversation container
    const conv = document.querySelector('[data-app-section="ConversationContainer"]');
    if (conv) return { parent: conv, before: conv.firstChild };

    // Fallback B: any role=main region
    const main = document.querySelector('[role="main"]');
    if (main) return { parent: main, before: main.firstChild };

    return null;
  }

  async function injectPanel() {
    if (isComposeView()) {
      document.getElementById("osn-panel")?.remove();
      injecting = false;
      return true; // don't retry
    }

    const insertion = findInsertionPoint();
    if (!insertion) return false;

    const key = getKey();
    if (!key) return false; // can't determine email ID yet — retry

    // Already injected for this key — nothing to do
    if (document.getElementById("osn-panel") && key === currentKey) return true;

    // Guard against duplicate injection if a previous async call is still in flight
    if (injecting) return true;
    injecting = true;

    try {
      // Remove old panel if switching emails
      document.getElementById("osn-panel")?.remove();

      currentKey = key;

      // Collapsed icon uses position:absolute — parent must be position:relative
      if (getComputedStyle(insertion.parent).position === "static") {
        insertion.parent.style.position = "relative";
      }

      const panel = buildPanel();
      insertion.parent.insertBefore(panel, insertion.before);

      const notes = await loadNotes(currentKey);
      renderNotes(notes);

      // Detect pop-out by checking the insertion parent rather than calling isPopout().
      // With match_origin_as_fallback active the script can run in a sub-frame where
      // _owa_projection_root isn't present in that frame's document, so isPopout()
      // would return false even though we're inside a pop-out window.
      const inPopout = insertion.parent.id === "_owa_projection_root";
      if (inPopout) panel.classList.add("osn-popout");

      // Always start collapsed
      collapsed = true;
      updateCollapsedPlus(notes.length);
      panel.classList.add("osn-collapsed");
      panel.querySelector("#osn-body").classList.add("osn-hidden");
      panel.querySelector("#osn-input-area").classList.add("osn-hidden");
      panel.querySelector("#osn-btn-add").classList.add("osn-hidden");

      // In pop-out, align the collapsed icon with the thread subject header bar.
      // Hide until positioned to avoid a flash at the wrong position.
      if (inPopout) {
        panel.style.visibility = "hidden";
        setTimeout(() => {
          const root = document.getElementById("_owa_projection_root");
          const heading = document.querySelector('[id$="_SUBJECT"], [role="heading"][aria-level="3"]');
          if (root && heading) {
            const rootRect = root.getBoundingClientRect();
            const headingRect = heading.getBoundingClientRect();
            panel.style.top = (headingRect.top - rootRect.top - 6) + "px";
          }
          panel.style.visibility = "";
        }, OSN.POPOUT_ALIGN_MS);
      }

      updateAllListBadges();
      attachListObserver();
    } finally {
      injecting = false; // always reset, even if an error occurs mid-injection
    }

    return true;
  }

  // ── Observer: watch for email navigation ───────────────────────────────────
  let lastUrl = location.href;
  let retryTimer = null;
  let retryCount = 0;

  function tryInject() {
    if (!isExtensionAlive() || ++retryCount > OSN.MAX_RETRIES) return;
    injectPanel().then((ok) => {
      if (!ok) {
        clearTimeout(retryTimer);
        retryTimer = setTimeout(tryInject, OSN.RETRY_DELAY_MS);
      }
    }).catch(() => {});
  }

  function scheduleInject(delay) {
    clearTimeout(retryTimer);
    retryCount = 0;
    retryTimer = setTimeout(tryInject, delay);
  }

  let badgeTimer = null;
  function scheduleBadges(delay) {
    clearTimeout(badgeTimer);
    badgeTimer = setTimeout(updateAllListBadges, delay);
  }

  // Shallow observer on document.body — watches only direct children for URL and
  // pane changes. Outlook is a SPA that rewrites the URL without navigating, so this
  // is how we detect the user switching between email threads.
  // Kept shallow (subtree: false) to avoid firing on every DOM mutation inside Outlook.
  new MutationObserver(() => {
    // Wire up the pop-out deep observer as soon as _owa_projection_root lands in body.
    // attachPopoutObserver is idempotent so this is safe to call on every callback.
    attachPopoutObserver();

    const urlChanged = location.href !== lastUrl;
    if (urlChanged) {
      lastUrl = location.href;
      cachedIsCompose = isComposeView();
      document.getElementById("osn-panel")?.remove();
      if (cachedIsCompose) return;
      scheduleInject(OSN.NAV_INJECT_DELAY);
      scheduleBadges(OSN.BADGE_NAV_DELAY);
      return;
    }

    if (cachedIsCompose) {
      document.getElementById("osn-panel")?.remove();
      return;
    }

    const panel = document.getElementById("osn-panel");
    const key = getKey();
    if (!panel || key !== currentKey) {
      if (key !== currentKey) panel?.remove();
      scheduleInject(OSN.RETRY_DELAY_MS);
    }
  }).observe(document.body, { childList: true, subtree: false });

  // Separate targeted observer for badge updates on the email list panel.
  // Attached after a delay on startup and after each successful injection.
  // Stores the observed element so we can detect if Outlook removes and replaces it
  // (virtual-scroll SPAs do this), which would silently disconnect the observer.
  let listObserver = null;
  let listObserverTarget = null;
  function attachListObserver() {
    if (listObserver && listObserverTarget?.isConnected) return; // still valid
    // Target was removed — disconnect the stale observer before re-attaching
    if (listObserver) {
      listObserver.disconnect();
      listObserver = null;
      listObserverTarget = null;
    }
    const listEl = document.querySelector('[role="list"][aria-label]')
      ?? document.querySelector('[data-app-section="MailList"]');
    if (!listEl) return;
    listObserverTarget = listEl;
    listObserver = new MutationObserver(() => scheduleBadges(OSN.BADGE_DELAY_MS));
    listObserver.observe(listEl, { childList: true, subtree: true });
  }

  // ── Pop-out observer ────────────────────────────────────────────────────────
  // Outlook pop-out windows load email content asynchronously deep inside
  // #_owa_projection_root. The shallow body observer can't see those mutations,
  // so we attach a dedicated deep observer here to re-trigger injection as content arrives.
  //
  // Why match_origin_as_fallback is set in the manifest: Outlook's pop-out window
  // renders inside a frame whose URL is about:blank. The standard content_scripts
  // "matches" patterns only compare against the frame's own URL, which would never
  // match about:blank and would prevent the script from running in pop-out windows.
  // match_origin_as_fallback tells Chrome to fall back to the top-level frame's URL
  // for matching purposes — which is always one of the outlook.com / outlook.live.com /
  // outlook.cloud.microsoft URLs listed in "matches" — so the script can inject there.
  //
  // This observer is idempotent (guarded by popoutObserverAttached) and is called
  // from both startup and the shallow body observer so it fires as soon as
  // _owa_projection_root appears in the DOM, regardless of timing.
  let popoutObserverAttached = false;
  function attachPopoutObserver() {
    if (popoutObserverAttached) return;
    const root = document.getElementById("_owa_projection_root");
    if (!root) return; // not a pop-out, or root not yet in DOM — will retry via body observer
    popoutObserverAttached = true;
    new MutationObserver(() => {
      if (!document.getElementById("osn-panel")) scheduleInject(OSN.RETRY_DELAY_MS);
    }).observe(root, { childList: true, subtree: true });
  }

  // ── Startup ─────────────────────────────────────────────────────────────────
  cachedIsCompose = isComposeView();
  migrateEncodedKeys();                 // one-time migration of URL-encoded storage keys
  bootstrapIndex();                     // build count index for pre-existing installations
  tryInject();                          // attempt panel injection immediately
  scheduleBadges(500);                  // badge scan after email list has likely rendered
  setTimeout(attachListObserver, 1000); // list observer after list has likely rendered
  attachPopoutObserver();               // pop-out: try immediately; body observer handles late case
})();

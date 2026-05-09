(() => {
  let currentKey = null;
  let collapsed = false;
  let lastDeleted = null;  // for undo
  let dragSrcId = null;    // for drag-and-drop reorder

  function isExtensionAlive() {
    try {
      // Accessing chrome.runtime.id throws if the context is invalidated
      return !!chrome.runtime?.id;
    } catch {
      return false;
    }
  }

  // ── Compose view detection ──────────────────────────────────────────────────
  function isComposeView() {
    // URL-based patterns
    if (/\/(compose|action\/compose|deeplink\/compose)/i.test(location.href)) return true;
    // DOM-based: Send button + an editable To field (input/textarea) = compose
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
  // Use the page URL (includes message/conversation ID) as the stable key.
  // Fall back to subject text if URL doesn't change between messages.
  function getKey() {
    const url = location.href;
    // Standard reading pane: /mail/inbox/id/XXXXX or /mail/id/XXXXX
    const match = url.match(/\/(?:id|read)\/([^/?#]+)/i);
    if (match) return "osn_" + decodeURIComponent(match[1]);

    // Pop-out: window is about:blank but opener is the main Outlook tab —
    // read the opener's URL to get the same ID the main window uses.
    if (isPopout() && window.opener) {
      try {
        const openerUrl = window.opener.location.href;
        const openerMatch = openerUrl.match(/\/(?:id|read)\/([^/?#]+)/i);
        if (openerMatch) return "osn_" + decodeURIComponent(openerMatch[1]);
      } catch { /* cross-origin guard */ }
    }

    // Fallback: data-convid on any element
    const convEl = document.querySelector('[data-convid]');
    if (convEl) return "osn_" + convEl.getAttribute('data-convid');

    // Last resort: subject heading text
    const subj = document.querySelector('[role="heading"][aria-level="2"], [role="heading"][aria-level="1"], [role="heading"][aria-level="3"]');
    if (subj && subj.textContent.trim()) {
      return "osn_subj_" + subj.textContent.trim().slice(0, 80);
    }
    return null; // refuse to use a shared fallback key
  }

  // ── Key migration ───────────────────────────────────────────────────────────
  // Earlier versions stored keys with URL-encoded IDs (osn_AAQ...%3D).
  // Now keys use decoded IDs (osn_AAQ...=) to match data-convid attributes.
  // This runs once at startup to migrate old entries and drop stale orphans.
  function migrateEncodedKeys() {
    if (!isExtensionAlive()) return;
    try {
      chrome.storage.local.get(null, (allData) => {
        const toSet = {};
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
        if (toRemove.length) chrome.storage.local.remove(toRemove);
        if (Object.keys(toSet).length) chrome.storage.local.set(toSet);
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
      chrome.storage.local.set({ [key]: notes });
    } catch { /* context invalidated — ignore */ }
  }

  // ── Panel HTML ──────────────────────────────────────────────────────────────
  const POSTIT_URL = chrome.runtime.getURL("icons/postit.png");

  function buildPanel() {
    const panel = document.createElement("div");
    panel.id = "osn-panel";
    panel.style.setProperty("--osn-postit-url", `url("${POSTIT_URL}")`);

    panel.innerHTML = `
      <div id="osn-header" data-tooltip="Click to collapse">
        <span id="osn-title"><span id="osn-title-text">Sticky Notes</span></span>
        <span id="osn-collapsed-plus">+</span>
        <button id="osn-btn-add" data-tooltip="New note">+</button>
      </div>
      <div id="osn-body">
        <span id="osn-empty">No notes yet for this thread.</span>
      </div>
      <div id="osn-undo-bar" class="osn-hidden">
        Note deleted — <button id="osn-undo-btn">Undo</button>
      </div>
      <div id="osn-storage-warning" class="osn-hidden"></div>
      <div id="osn-input-area" class="osn-hidden">
        <textarea id="osn-textarea" maxlength="500" placeholder="Type your notes and click save"></textarea>
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
    resizeHandle.innerHTML = "⋯";
    panel.appendChild(resizeHandle);

    resizeHandle.addEventListener("mousedown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const body = document.getElementById("osn-body");
      const startY = e.clientY;
      const startHeight = body.getBoundingClientRect().height;

      const onMove = (me) => {
        const newHeight = Math.max(60, startHeight + (me.clientY - startY));
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
      const remaining = 500 - len;
      if (remaining <= 100) {
        counter.textContent = `${len} / 500`;
        counter.className = remaining <= 20 ? "osn-char-urgent" : "";
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
    if (collapsed) {
      panel?.classList.add("osn-collapsed");
      if (isPopout()) {
        const root = document.getElementById("_owa_projection_root");
        const heading = document.querySelector('[id$="_SUBJECT"], [role="heading"][aria-level="3"]');
        if (root && heading) {
          const rootRect = root.getBoundingClientRect();
          const headingRect = heading.getBoundingClientRect();
          panel.style.top = (headingRect.top - rootRect.top - 6) + "px";
        }
      }
      body.style.display = "none";
      input.style.display = "none";
      if (addBtn) addBtn.style.display = "none";
clearUndo();
    } else {
      panel?.classList.remove("osn-collapsed");
      if (isPopout()) panel.style.top = "";
      body.style.display = "";
      input.style.display = "";
      input.classList.add("osn-hidden");
      const hasNotes = document.querySelectorAll(".osn-note").length > 0;
      if (addBtn) addBtn.style.display = hasNotes ? "" : "none";
    }
  }

  async function saveNote() {
    if (!currentKey) return;
    const ta = document.getElementById("osn-textarea");
    const text = ta?.value.trim();
    if (!text) return;

    const notes = await loadNotes(currentKey);
    notes.unshift({ id: Date.now(), text, date: new Date().toLocaleString() });
    saveNotes(currentKey, notes);
    ta.value = "";
    hideInput();
    renderNotes(notes);
    clearUndo();
    checkStorageQuota();
    setTimeout(() => updateCurrentListBadge(notes.length), 600);
  }

  async function deleteNote(id) {
    const allNotes = await loadNotes(currentKey);
    const idx = allNotes.findIndex((n) => n.id === id);
    if (idx === -1) return;

    // Store for undo
    lastDeleted = { note: allNotes[idx], index: idx };
    document.getElementById("osn-undo-bar")?.classList.remove("osn-hidden");

    const remaining = allNotes.filter((n) => n.id !== id);
    saveNotes(currentKey, remaining);
    renderNotes(remaining);
    setTimeout(() => updateCurrentListBadge(remaining.length), 600);

    // Auto-collapse when the last note is deleted
    if (remaining.length === 0 && !collapsed) {
      toggleCollapse();
      hideInput();
    }
  }

  async function undoDelete() {
    if (!lastDeleted) return;
    const notes = await loadNotes(currentKey);
    notes.splice(lastDeleted.index, 0, lastDeleted.note);
    saveNotes(currentKey, notes);
    lastDeleted = null;
    // Re-expand if auto-collapsed
    if (collapsed) toggleCollapse();
    renderNotes(notes);
    clearUndo();
    setTimeout(() => updateCurrentListBadge(notes.length), 600);
  }

  function clearUndo() {
    lastDeleted = null;
    document.getElementById("osn-undo-bar")?.classList.add("osn-hidden");
  }

  function checkStorageQuota() {
    if (!isExtensionAlive()) return;
    try {
      chrome.storage.local.getBytesInUse(null, (bytes) => {
        const QUOTA = 10 * 1024 * 1024; // 10 MB
        const pct = bytes / QUOTA;
        const warning = document.getElementById("osn-storage-warning");
        if (!warning) return;
        if (pct > 0.8) {
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
    div.innerHTML = `
      <div class="osn-edit-area">
        <textarea class="osn-edit-textarea" maxlength="500">${escapeHtml(note.text)}</textarea>
        <div class="osn-edit-footer">
          <span class="osn-edit-char-count">${note.text.length} / 500</span>
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
      div.querySelector(".osn-edit-char-count").textContent = `${ta.value.length} / 500`;
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
        notes[idx] = { ...notes[idx], text: newText, date: new Date().toLocaleString() };
        saveNotes(currentKey, notes);
        renderNotes(notes);
      }
    }

    async function cancelEdit() {
      const notes = await loadNotes(currentKey);
      renderNotes(notes);
    }
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function updateCollapsedPlus(noteCount) {
    const plus = document.getElementById("osn-collapsed-plus");
    if (!plus) return;
    if (noteCount === 0) {
      plus.textContent = "+";
      plus.style.display = "";
    } else {
      plus.textContent = String(noteCount);
      plus.style.display = "";
    }
  }

  function renderNotes(notes) {
    const body = document.getElementById("osn-body");
    if (!body) return;

    const addBtn = document.getElementById("osn-btn-add");
    if (addBtn) addBtn.style.display = notes.length > 0 ? "" : "none";
    updateCollapsedPlus(notes.length);

    body.innerHTML = "";

    if (!notes.length) {
      body.innerHTML = '<span id="osn-empty">No notes yet for this thread.</span>';
      return;
    }

    notes.forEach((note) => {
      const div = document.createElement("div");
      div.className = "osn-note";
      div.draggable = true;
      div.dataset.id = String(note.id);
      div.innerHTML = `
        <span class="osn-drag-handle" title="Drag to reorder">⠿</span>
        <span class="osn-note-text">${escapeHtml(note.text)}</span>
        <div class="osn-note-meta">
          <button class="osn-note-edit" title="Edit">✎</button>
          <button class="osn-note-delete" title="Delete">✕</button>
          <span class="osn-note-date">${escapeHtml(note.date)}</span>
        </div>
      `;

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

  // Badge the currently selected list item (called after save/delete, deferred)
  function updateCurrentListBadge(noteCount) {
    // Find the email list row for the currently open email via data-convid
    const urlId = location.href.match(/\/id\/([^/?#]+)/i)?.[1];
    if (!urlId) return;
    const decoded = decodeURIComponent(urlId);
    const el = document.querySelector(`[data-convid="${CSS.escape(decoded)}"]`);
    if (el) setBadgeOnElement(el, noteCount);
  }

  // On load / navigation: scan all list items against stored keys
  function updateAllListBadges() {
    if (!isExtensionAlive()) return;
    try {
      // Fetch only our own keys by passing an array of known keys would require
      // knowing them upfront; instead we fetch all and filter by prefix immediately.
      chrome.storage.local.get(null, (allData) => {
        const keysWithNotes = new Map(
          Object.entries(allData)
            .filter(([k, v]) => k.startsWith("osn_") && Array.isArray(v) && v.length > 0)
            .map(([k, v]) => [k.slice(4), v.length]) // strip "osn_" → raw id → count
        );
        if (keysWithNotes.size === 0) return;

        // Outlook email list items use data-convid (decoded base64 exchange ID).
        // Stored keys may be encoded or decoded depending on when they were saved —
        // try both so old notes still get badges.
        const seen = new Set();
        document.querySelectorAll('[data-convid]').forEach((el) => {
          const val = el.getAttribute('data-convid');
          if (seen.has(val)) return; // skip nested duplicates, badge outermost only
          seen.add(val);
          const count = keysWithNotes.get(val);
          if (count) setBadgeOnElement(el, count);
        });
      });
    } catch { /* context invalidated */ }
  }

  // ── Injection ───────────────────────────────────────────────────────────────
  // Find the stable conversation container and inject before the scroll area.
  function findInsertionPoint() {
    // Primary: stable ID present in new Outlook (outlook.cloud.microsoft)
    const container = document.querySelector("#ConversationReadingPaneContainer");
    if (container) {
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
      return true; // don't retry
    }

    const insertion = findInsertionPoint();
    if (!insertion) return false;

    const key = getKey();
    if (!key) return false; // can't determine email ID yet — retry

    // Already injected for this key — just re-render
    if (document.getElementById("osn-panel") && key === currentKey) return true;

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

    // Mark panel for pop-out specific styling
    if (isPopout()) panel.classList.add("osn-popout");

    // Always start collapsed
    collapsed = true;
    updateCollapsedPlus(notes.length);
    panel.classList.add("osn-collapsed");
    panel.querySelector("#osn-body").style.display = "none";
    panel.querySelector("#osn-input-area").style.display = "none";
    panel.querySelector("#osn-btn-add").style.display = "none";

    // In pop-out, align the collapsed icon with the thread subject header bar.
    // Hide until positioned to avoid a jump.
    if (isPopout()) {
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
      }, 600);
    }

    updateAllListBadges();
    return true;
  }

  // ── Observer: watch for email navigation ───────────────────────────────────
  let lastUrl = location.href;
  let retryTimer = null;
  let retryCount = 0;
  const MAX_RETRIES = 20; // ~8 seconds of attempts

  function tryInject() {
    if (!isExtensionAlive() || retryCount++ > MAX_RETRIES) return;
    injectPanel().then((ok) => {
      if (!ok) {
        clearTimeout(retryTimer);
        retryTimer = setTimeout(tryInject, 400);
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

  // Single observer handles URL changes and reading pane swaps
  new MutationObserver((mutations) => {
    // Always check for compose view — it can appear without a URL change
    if (isComposeView()) {
      document.getElementById("osn-panel")?.remove();
      return;
    }

    if (location.href !== lastUrl) {
      lastUrl = location.href;
      scheduleInject(600);
      scheduleBadges(1200); // list re-renders after navigation
      return;
    }

    if (!document.getElementById("osn-panel") || getKey() !== currentKey) {
      scheduleInject(400);
      return;
    }

    // Re-badge any newly rendered email list rows (virtual scroll)
    for (const mutation of mutations) {
      for (const node of mutation.addedNodes) {
        if (node.nodeType !== 1) continue;
        if (node.hasAttribute?.("data-convid") || node.querySelector?.("[data-convid]")) {
          scheduleBadges(100);
          return;
        }
      }
    }
  }).observe(document.body, { childList: true, subtree: true });

  // Migrate any old URL-encoded storage keys to decoded format
  migrateEncodedKeys();
  // Initial injection attempt
  tryInject();
  // Badge scan runs independently — fires after the email list has rendered on load
  scheduleBadges(500);
})();

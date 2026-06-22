(function syncLucasKeepNote() {
  "use strict";

  const BRIDGE_PORTS = [8765, 8766, 8767, 8768, 8769, 8770, 8771, 8772];
  let lastSyncedText = "";
  let pendingText = "";
  let lastActiveNoteRoot = null;

  trackActiveKeepNote();
  setInterval(syncActiveKeepNote, 5000);
  document.addEventListener("input", syncActiveKeepNote, true);
  document.addEventListener("keyup", syncActiveKeepNote, true);
  document.addEventListener("pointerup", () => setTimeout(syncActiveKeepNote, 400), true);
  setTimeout(syncActiveKeepNote, 1200);

  function syncActiveKeepNote() {
    const root = activeKeepNoteRoot({ ignoreRemembered: true, requireEditor: true }) || activeKeepNoteRoot();
    if (!root) return;
    const text = extractKeepText(root);
    if (!text || text === lastSyncedText || text === pendingText) return;
    pendingText = text;
    postToBridge({
      text,
      title: extractKeepTitle(text, root),
      url: keepNoteUrl(root),
      synced_at: new Date().toISOString(),
    }).then((payload) => {
      if (payload && payload.ok) lastSyncedText = text;
    }).finally(() => {
      if (pendingText === text) pendingText = "";
    });
  }

  async function postToBridge(note) {
    for (const port of BRIDGE_PORTS) {
      try {
        const response = await fetch(`http://127.0.0.1:${port}/source/google-keep`, {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify(note),
        });
        const payload = await response.json().catch(() => ({}));
        if (response.ok && payload && payload.ok) return payload;
      } catch (_error) {
        // Try the next bridge port.
      }
    }
    return null;
  }

  function extractKeepText(root) {
    if (/accounts\.google\.com|ServiceLogin/i.test(window.location.href) || /sign in/i.test(document.title || "")) {
      return "";
    }
    const candidates = Array.from(root.querySelectorAll('[role="textbox"], [contenteditable="true"], .IZ65Hb-TBnied'));
    const text = candidates
      .map((node) => node.innerText || node.textContent || "")
      .map((value) => value.trim())
      .filter(Boolean)
      .join("\n")
      .trim();
    return text || (root.innerText || "").trim();
  }

  function activeKeepNoteRoot(options = {}) {
    const visibleRoots = keepNoteCandidates(options).sort(compareKeepCandidates);
    const focusedDialog = visibleRoots.find((dialog) => dialog.contains(document.activeElement));
    if (focusedDialog) return focusedDialog;
    if (!options.ignoreRemembered && lastActiveNoteRoot && visibleRoots.includes(lastActiveNoteRoot)) {
      return lastActiveNoteRoot;
    }
    return visibleRoots[0] || null;
  }

  function trackActiveKeepNote() {
    const rememberNoteRoot = (event) => {
      const dialog = event.target?.closest?.('[role="dialog"], [aria-modal="true"]');
      if (dialog && isVisibleElement(dialog)) lastActiveNoteRoot = dialog;
    };
    document.addEventListener("pointerdown", rememberNoteRoot, true);
    document.addEventListener("focusin", rememberNoteRoot, true);
  }

  function keepNoteCandidates(options = {}) {
    const roots = new Set();
    document.querySelectorAll('[role="dialog"], [aria-modal="true"]').forEach((node) => roots.add(node));
    document.querySelectorAll('[role="textbox"], [contenteditable="true"], .IZ65Hb-TBnied').forEach((node) => {
      const root = noteRootFromEditor(node);
      if (root) roots.add(root);
    });
    return Array.from(roots)
      .filter(isVisibleElement)
      .filter((root) => !options.requireEditor || hasKeepEditor(root))
      .filter(looksLikeOpenNote);
  }

  function noteRootFromEditor(node) {
    const root = node.closest('[role="dialog"], [aria-modal="true"]');
    if (root) return root;
    return node.closest('[role="article"], [tabindex]') || null;
  }

  function looksLikeOpenNote(node) {
    const rect = node.getBoundingClientRect();
    const viewportArea = Math.max(1, window.innerWidth * window.innerHeight);
    const area = rect.width * rect.height;
    const centerX = rect.left + rect.width / 2;
    const centerY = rect.top + rect.height / 2;
    const isCentered = centerX > window.innerWidth * 0.18 &&
      centerX < window.innerWidth * 0.82 &&
      centerY > window.innerHeight * 0.10 &&
      centerY < window.innerHeight * 0.90;
    const isEditorSized = rect.width >= 300 && rect.height >= 140 && area < viewportArea * 0.82;
    const isDialog = node.getAttribute("role") === "dialog" || node.getAttribute("aria-modal") === "true";
    return isDialog || (isEditorSized && isCentered);
  }

  function isVisibleElement(node) {
    const rect = node.getBoundingClientRect();
    const style = window.getComputedStyle(node);
    return rect.width > 80 &&
      rect.height > 80 &&
      rect.bottom > 0 &&
      rect.right > 0 &&
      style.visibility !== "hidden" &&
      style.display !== "none" &&
      Number(style.opacity || "1") !== 0;
  }

  function hasKeepEditor(node) {
    return Boolean(node.querySelector('[role="textbox"], [contenteditable="true"], .IZ65Hb-TBnied'));
  }

  function compareKeepCandidates(a, b) {
    return candidateScore(b) - candidateScore(a) || documentPositionSort(a, b);
  }

  function candidateScore(node) {
    const rect = node.getBoundingClientRect();
    const viewportArea = Math.max(1, window.innerWidth * window.innerHeight);
    const centerX = rect.left + rect.width / 2;
    const centerY = rect.top + rect.height / 2;
    const centered = 1 - Math.min(1, (
      Math.abs(centerX - window.innerWidth / 2) / (window.innerWidth / 2) +
      Math.abs(centerY - window.innerHeight / 2) / (window.innerHeight / 2)
    ) / 2);
    let score = centered * 300;
    score += Math.min((rect.width * rect.height) / viewportArea, 0.55) * 250;
    score += numericZIndex(node) * 2;
    if (node.contains(document.activeElement)) score += 500;
    if (node.getAttribute("role") === "dialog" || node.getAttribute("aria-modal") === "true") score += 400;
    if (hasKeepEditor(node)) score += 150;
    return score;
  }

  function numericZIndex(node) {
    const value = Number(window.getComputedStyle(node).zIndex);
    return Number.isFinite(value) ? value : 0;
  }

  function documentPositionSort(a, b) {
    if (a === b) return 0;
    return a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING ? 1 : -1;
  }

  function keepNoteUrl(root) {
    const link = root?.querySelector('a[href*="/u/"][href*="/notes/"], a[href*="/notes/"]')?.href;
    return link || window.location.href;
  }

  function extractKeepTitle(text, root) {
    const titleCandidate = root?.querySelector('[role="textbox"], [contenteditable="true"]');
    const title = (titleCandidate?.innerText || titleCandidate?.textContent || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .find(isRealKeepTitle);
    return title || firstRealKeepLine(text) || "Untitled Keep note";
  }

  function firstRealKeepLine(text) {
    return String(text || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .find(isRealKeepTitle);
  }

  function isRealKeepTitle(line) {
    return Boolean(line) && !/^(take a note|title|note)$/i.test(line.replace(/[.]+$/g, ""));
  }
})();

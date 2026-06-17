const SALES_HISTORY_URL = "https://app.cardladder.com/sales-history";
const BRIDGE_PORTS = [8765, 8766, 8767, 8768, 8769, 8770, 8771, 8772];
const BRIDGE_URLS = BRIDGE_PORTS.map((port) => `http://127.0.0.1:${port}`);
const BRIDGE_ALARM_NAME = "cardladder-bridge-poll";
const BRIDGE_POLL_MS = 1000;
const BETWEEN_ROWS_MS = 1200;
const OCR_SETTLE_MS = 600;
const OCR_RETRY_MS = 800;
const CARDLADDER_BACKGROUND_VERSION = "2026-06-16-dom-comp-sweep-v1";

let runInProgress = false;
let activeWindowId = null;
let activeBridgeUrl = BRIDGE_URLS[0];
let cancelRequested = false;

chrome.runtime.onInstalled.addListener(() => {
  chrome.alarms.create(BRIDGE_ALARM_NAME, { periodInMinutes: 0.05 });
  pollDesktopBridge();
});

chrome.action.onClicked.addListener(() => startCardLadderRun(true));
pollDesktopBridge();
setInterval(pollDesktopBridge, BRIDGE_POLL_MS);
chrome.runtime.onStartup?.addListener(() => {
  chrome.alarms.create(BRIDGE_ALARM_NAME, { periodInMinutes: 0.05 });
  pollDesktopBridge();
});
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === BRIDGE_ALARM_NAME) pollDesktopBridge();
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === "CARDLADDER_SAVE_QUEUE") {
    chrome.storage.local.set({ cardladderQueue: message.queue, cardladderResults: [] })
      .then(() => sendResponse({ ok: true }));
    return true;
  }

  if (message.type === "CARDLADDER_SYNC_NOW") {
    startCardLadderRun(true, { keepWindowOpen: Boolean(message.keepWindowOpen) }).then(sendResponse);
    return true;
  }

  if (message.type === "CARDLADDER_CANCEL_RUN") {
    cancelRun().then(sendResponse);
    return true;
  }

  if (message.type === "CARDLADDER_GET_STATUS") {
    pollDesktopBridge()
      .then(() => chrome.storage.local.get(["cardladderStatus", "cardladderResults", "cardladderQueue"]))
      .then(sendResponse);
    return true;
  }

  if (message.type === "CARDLADDER_CAPTURE_ACTIVE_TAB") {
    captureActiveTabWithOcr(message.row || {}).then(sendResponse);
    return true;
  }

});

async function startCardLadderRun(focusWindow, options = {}) {
  if (runInProgress) return { ok: false, error: "Card Ladder run already in progress" };
  runInProgress = true;
  cancelRequested = false;

  try {
    const { cardladderQueue } = await chrome.storage.local.get(["cardladderQueue"]);
    const rows = cardladderQueue?.rows || [];
    if (!rows.length) throw new Error("No Card Ladder queue loaded.");

    await chrome.storage.local.set({
      cardladderStatus: {
        ok: true,
        stage: "opening Card Ladder",
        total: rows.length,
        completed: 0,
        startedAt: new Date().toISOString(),
      },
      cardladderResults: [],
    });

    const tab = await createSalesHistoryWindow(focusWindow);
    await runRows(tab.id, rows, options);
    return { ok: true };
  } catch (error) {
    await chrome.storage.local.set({
      cardladderStatus: {
        ok: false,
        stage: "failed",
        error: String(error?.message || error),
        finishedAt: new Date().toISOString(),
      },
    });
    return { ok: false, error: String(error?.message || error) };
  } finally {
    runInProgress = false;
  }
}

async function pollDesktopBridge() {
  try {
    if (runInProgress) return;
    for (const bridgeUrl of prioritizedBridgeUrls()) {
      const response = await fetch(`${bridgeUrl}/command?${extensionMetadataParams()}`).then((r) => r.json()).catch(() => null);
      const command = response?.command;
      if (!response) continue;
      activeBridgeUrl = bridgeUrl;
      if (!command || command.type !== "RUN_ALL_COMPS") continue;
      await fetch(`${bridgeUrl}/ack`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ id: command.id }),
      }).catch(() => {});
      await startBridgeRun(command.queue || [], bridgeUrl);
      return;
    }
  } catch (_error) {
    return;
  }
}

function prioritizedBridgeUrls() {
  return [activeBridgeUrl, ...BRIDGE_URLS.filter((url) => url !== activeBridgeUrl)];
}

function extensionMetadataParams() {
  const manifest = chrome.runtime.getManifest?.() || {};
  return new URLSearchParams({
    extensionVersion: CARDLADDER_BACKGROUND_VERSION,
    manifestVersion: manifest.version || "",
    extensionName: manifest.name || "",
    extensionUrl: chrome.runtime.getURL?.("") || "",
  }).toString();
}

async function startBridgeRun(rows, bridgeUrl = activeBridgeUrl) {
  if (runInProgress) return;
  runInProgress = true;
  cancelRequested = false;
  activeBridgeUrl = bridgeUrl;
  try {
    if (!rows.length) throw new Error("Desktop bridge sent no Card Ladder rows.");
    await chrome.storage.local.set({
      cardladderStatus: {
        ok: true,
        stage: "opening Card Ladder",
        total: rows.length,
        completed: 0,
        startedAt: new Date().toISOString(),
      },
      cardladderResults: [],
    });
    const tab = await createSalesHistoryWindow(true);
    await runRows(tab.id, rows, { postToBridge: true });
  } catch (error) {
    await postBridgeFinish({ ok: false, error: String(error?.message || error) });
  } finally {
    runInProgress = false;
  }
}

async function createSalesHistoryWindow(focusWindow) {
  if (activeWindowId !== null) {
    await chrome.windows.remove(activeWindowId).catch(() => {});
    activeWindowId = null;
  }
  const win = await chrome.windows.create({
    url: SALES_HISTORY_URL,
    focused: focusWindow,
    type: "normal",
    width: 1320,
    height: 920,
  });
  activeWindowId = win.id;
  const tab = win.tabs && win.tabs[0];
  if (!tab?.id) throw new Error("Could not create Card Ladder window.");
  return tab;
}

async function runRows(tabId, rows, options = {}) {
  const started = Date.now();
  const results = [];

  for (let index = 0; index < rows.length; index += 1) {
    await throwIfCancelled(options);
    const row = rows[index];
    await waitForTabNotLoading(tabId);
    await injectContent(tabId);

    await chrome.storage.local.set({
      cardladderStatus: {
        ok: true,
        stage: "looking up",
        total: rows.length,
        completed: index,
        current: row,
        updatedAt: new Date().toISOString(),
      },
    });

    const result = stampResult(await lookupRowWithRetries(tabId, row));
    await throwIfCancelled(options);

    results.push(result);
    if (options.postToBridge) await postBridgeResult(result);
    await chrome.storage.local.set({
      cardladderResults: results,
      cardladderStatus: {
        ok: true,
        stage: "looking up",
        total: rows.length,
        completed: index + 1,
        lastResult: result,
        updatedAt: new Date().toISOString(),
      },
    });

    await delay(BETWEEN_ROWS_MS);
  }

  await chrome.storage.local.set({
    cardladderResults: results,
    cardladderStatus: {
      ok: true,
      stage: "finished",
      total: rows.length,
      completed: rows.length,
      found: results.filter((result) => result.value != null).length,
      windowKeptOpen: Boolean(options.keepWindowOpen),
      finishedAt: new Date().toISOString(),
    },
  });
  if (options.postToBridge) {
    await postBridgeFinish({
      ok: true,
      total: rows.length,
      found: results.filter((result) => result.value != null).length,
    });
  }
  if (!options.keepWindowOpen) {
    await closeActiveWindow();
  }
}

async function throwIfCancelled(options = {}) {
  if (cancelRequested) throw new Error("Card Ladder run cancelled.");
  if (!options.postToBridge) return;
  const status = await fetch(`${activeBridgeUrl}/status`).then((r) => r.json()).catch(() => null);
  if (status?.cancelRequested) {
    cancelRequested = true;
    throw new Error("Card Ladder run cancelled by L.U.C.A.S.");
  }
}

async function cancelRun() {
  cancelRequested = true;
  await chrome.storage.local.set({
    cardladderStatus: {
      ok: false,
      stage: "cancelled",
      error: "Card Ladder run cancelled by user.",
      finishedAt: new Date().toISOString(),
    },
  });
  await closeActiveWindow();
  await postBridgeFinish({ ok: false, cancelled: true, error: "Card Ladder run cancelled by user." }).catch(() => {});
  runInProgress = false;
  return { ok: true };
}

async function lookupRowWithRetries(tabId, row) {
  const pageResult = await submitRowWithGrader(tabId, row);

  if (["error", "invalid_cert"].includes(pageResult?.status)) return pageResult;
  if (pageResult?.status === "no_results") {
    if (pageResult?.ocr?.profileTitle) return {
      ...pageResult,
      noResultsFallback: "submit-response-profile",
    };
    const noResultsDomResult = await captureValueFromDom(tabId, row, pageResult);
    if (noResultsDomResult?.ocr?.profileTitle) {
      return {
        ...pageResult,
        ocr: {
          ...(pageResult.ocr || {}),
          ...(noResultsDomResult.ocr || {}),
          comps: [],
        },
        pageUrl: pageResult.pageUrl || noResultsDomResult.pageUrl || "",
        capturedAt: noResultsDomResult.capturedAt || pageResult.capturedAt || new Date().toISOString(),
        noResultsFallback: "dom-profile",
      };
    }
    const noResultsOcrResult = await captureValueWithOcr(tabId, row, pageResult);
    if (noResultsOcrResult?.ocr?.profileTitle) {
      return {
        ...pageResult,
        ocr: {
          ...(pageResult.ocr || {}),
          ...(noResultsOcrResult.ocr || {}),
          comps: [],
        },
        pageUrl: pageResult.pageUrl || noResultsOcrResult.pageUrl || "",
        capturedAt: noResultsOcrResult.capturedAt || pageResult.capturedAt || new Date().toISOString(),
        noResultsFallback: "ocr-profile",
      };
    }
    return {
      ...pageResult,
      ocr: {
        ...(pageResult.ocr || {}),
        ...(noResultsDomResult?.ocr || {}),
        ...(noResultsOcrResult?.ocr || {}),
        comps: [],
      },
      noResultsFallback: "profile-not-found",
      noResultsFallbackError: noResultsOcrResult?.ocr?.error || noResultsOcrResult?.error || noResultsDomResult?.error || "",
      noResultsFallbackDebugImage: noResultsOcrResult?.ocr?.debugImage || "",
    };
  }

  const domResult = await captureValueFromDom(tabId, row, pageResult);
  if (["invalid_cert", "no_results"].includes(domResult?.status)) return domResult;
  if (domResultLooksComplete(domResult)) return domResult;
  const domSweepResult = await captureValueFromDomSweep(tabId, row, pageResult, domResult);
  if (["invalid_cert", "no_results"].includes(domSweepResult?.status)) return domSweepResult;
  if (domResultLooksComplete(domSweepResult)) return domSweepResult;
  const expectedResultCount = Number(domSweepResult?.ocr?.resultCount);

  let lastResult = null;
  for (let attempt = 1; attempt <= 2; attempt += 1) {
    const invalidCheck = await checkInvalidCertToast(tabId, row);
    if (invalidCheck?.status === "invalid_cert") return invalidCheck;
    lastResult = await captureValueWithOcr(tabId, row, { ...pageResult, ocrAttempt: attempt });
    if (captureResultLooksComplete(lastResult, expectedResultCount)) return lastResult;
    await delay(OCR_RETRY_MS);
  }
  return markPartialCapture(mergeCaptureResults(domSweepResult, lastResult), expectedResultCount);
}

function mergeCaptureResults(primary, fallback) {
  if (!primary) return fallback || {};
  if (!fallback) return primary;
  return {
    ...primary,
    ...fallback,
    pageUrl: fallback.pageUrl || primary.pageUrl || "",
    capturedAt: fallback.capturedAt || primary.capturedAt || new Date().toISOString(),
    ocr: {
      ...(primary.ocr || {}),
      ...(fallback.ocr || {}),
      profileTitle: fallback.ocr?.profileTitle || primary.ocr?.profileTitle || "",
      profileGrader: fallback.ocr?.profileGrader || primary.ocr?.profileGrader || "",
      profileGrade: fallback.ocr?.profileGrade || primary.ocr?.profileGrade || "",
      resultCount: fallback.ocr?.resultCount ?? primary.ocr?.resultCount ?? null,
      comps: mergeOcrComps(primary.ocr?.comps, fallback.ocr?.comps),
    },
  };
}

function mergeOcrComps(...groups) {
  const ordered = [];
  for (const raw of groups.flat().filter(Boolean)) {
    const source = String(raw.source || "").replace(/\s+/g, " ").trim().toUpperCase();
    const title = String(raw.title || "").replace(/\s+/g, " ").trim();
    const date = String(raw.date_sold || raw.dateSold || "").replace(/\s+/g, " ").trim().toLowerCase();
    const price = String(raw.price || "").replace(/[$,\s]/g, "");
    const key = `${source}|${date}|${price}|${title.toLowerCase().replace(/[^a-z0-9]+/g, "").slice(0, 80)}`;
    if (!source && !title && !price) continue;
    if (ordered.some((item) => item.key === key)) continue;
    ordered.push({ key, comp: raw });
  }
  return ordered.map((item) => item.comp).slice(0, 5);
}

function stampResult(result) {
  return {
    ...(result || {}),
    extensionVersion: CARDLADDER_BACKGROUND_VERSION,
  };
}

function domResultLooksComplete(result) {
  return captureResultLooksComplete(result, Number(result?.ocr?.resultCount));
}

function captureResultLooksComplete(result, expectedResultCount = null) {
  if (result?.value == null) return false;
  const comps = Array.isArray(result?.ocr?.comps) ? result.ocr.comps : [];
  if (!comps.length) return false;
  const resultCount = Number.isFinite(expectedResultCount) && expectedResultCount > 0
    ? expectedResultCount
    : Number(result?.ocr?.resultCount);
  if (comps.length >= 3) return true;
  if (!Number.isFinite(resultCount) || resultCount <= 0) return comps.length >= 2;
  return comps.length >= Math.min(2, resultCount);
}

function markPartialCapture(result, expectedResultCount = null) {
  const comps = Array.isArray(result?.ocr?.comps) ? result.ocr.comps : [];
  const expected = Number.isFinite(expectedResultCount) && expectedResultCount > 0
    ? Math.min(2, expectedResultCount)
    : 2;
  const reason = result?.ocr?.error || result?.ocr?.evidence || result?.error || "Card Ladder page did not expose enough value/comp text.";
  return {
    ...(result || {}),
    value: null,
    status: "partial_comp_capture",
    error: `Only captured ${comps.length} comp(s); expected ${expected}. ${reason} Re-run this row.`,
    partialDiagnostics: {
      expectedComps: expected,
      capturedComps: comps.length,
      sourceStatus: result?.status || "",
      sourceError: result?.error || "",
      ocrError: result?.ocr?.error || "",
      ocrEvidence: result?.ocr?.evidence || "",
      profileTitle: result?.ocr?.profileTitle || "",
      resultCount: result?.ocr?.resultCount ?? null,
    },
  };
}

async function submitRowWithGrader(tabId, row) {
  const prepared = await chrome.tabs.sendMessage(tabId, {
    type: "CARDLADDER_PREPARE_CERT_MODAL",
    row,
  }).catch((error) => ({
    ok: false,
    error: String(error?.message || error),
  }));

  if (!prepared?.ok) {
    return {
      ...row,
      value: null,
      status: "error",
      error: prepared?.error || "Could not prepare Card Ladder cert modal",
      capturedAt: new Date().toISOString(),
    };
  }

  const grader = String(row.grader || "").toUpperCase();
  if (grader) {
    const selected = await selectGraderInPage(tabId, grader);
    if (!selected?.ok) {
      return {
        ...row,
        value: null,
        status: "error",
        error: selected?.error || `Could not select grader ${grader}`,
        capturedAt: new Date().toISOString(),
      };
    }
  }

  return chrome.tabs.sendMessage(tabId, {
    type: "CARDLADDER_SUBMIT_CERT_MODAL",
    row,
  }).catch((error) => ({
    ...row,
    value: null,
    status: "error",
    error: String(error?.message || error),
    capturedAt: new Date().toISOString(),
  }));
}

async function selectGraderInPage(tabId, grader) {
  return chrome.tabs.sendMessage(tabId, {
    type: "CARDLADDER_SELECT_GRADER",
    grader,
  }).catch((error) => ({ ok: false, error: String(error?.message || error) }));
}

async function captureValueFromDom(tabId, row, pageResult = {}) {
  const tab = await chrome.tabs.get(tabId).catch(() => null);
  const result = await chrome.tabs.sendMessage(tabId, {
    type: "CARDLADDER_EXTRACT_DOM_RESULT",
    row,
  }).catch((error) => ({
    ...row,
    value: null,
    status: "dom_error",
    error: String(error?.message || error),
    ocr: {
      ok: false,
      value: null,
      comps: [],
      error: String(error?.message || error),
      evidence: "Could not read Card Ladder page text through the content script.",
      debugImage: "",
    },
    capturedAt: new Date().toISOString(),
  }));
  return {
    ...result,
    pageUrl: result.pageUrl || pageResult.pageUrl || tab?.url || "",
    capturedAt: result.capturedAt || new Date().toISOString(),
  };
}

async function captureValueFromDomSweep(tabId, row, pageResult = {}, firstResult = null) {
  let merged = firstResult || await captureValueFromDom(tabId, row, pageResult);
  for (const position of ["top", "middle", "bottom"]) {
    await chrome.tabs.sendMessage(tabId, {
      type: "CARDLADDER_SCROLL_CAPTURE_AREA",
      position,
    }).catch(() => null);
    const next = await captureValueFromDom(tabId, row, pageResult);
    merged = mergeCaptureResults(merged, next);
    if (domResultLooksComplete(merged)) return merged;
  }
  return merged;
}

async function checkInvalidCertToast(tabId, row) {
  return chrome.tabs.sendMessage(tabId, {
    type: "CARDLADDER_CHECK_INVALID_CERT_TOAST",
    row,
  }).then((result) => {
    if (!result?.invalid) return { ...row, status: "ok" };
    return {
      ...row,
      value: null,
      status: "invalid_cert",
      error: result.error || "Card Ladder showed no information with this cert.",
      ocr: { ok: false, value: null, comps: [], evidence: result.error || "Invalid cert toast detected.", debugImage: "" },
      pageUrl: result.pageUrl || "",
      capturedAt: result.capturedAt || new Date().toISOString(),
    };
  }).catch(() => ({ ...row, status: "ok" }));
}

async function captureValueWithOcr(tabId, row, pageResult = {}) {
  await delay(OCR_SETTLE_MS);
  const tab = await chrome.tabs.get(tabId).catch(() => null);
  let captureError = "";
  if (tabId) await chrome.tabs.update(tabId, { active: true }).catch(() => {});
  const image = tab?.windowId ? await chrome.tabs.captureVisibleTab(tab.windowId, { format: "png" }).catch((error) => {
    captureError = String(error?.message || error);
    return "";
  }) : "";
  if (!image) {
    return {
      ...row,
      value: null,
      status: "ocr_error",
      error: "Could not capture Card Ladder screenshot" + (captureError ? `: ${captureError}` : ""),
      ocr: {
        ok: false,
        value: null,
        comps: [],
        error: "Could not capture Card Ladder screenshot" + (captureError ? `: ${captureError}` : ""),
        evidence: "Chrome did not provide a visible-tab screenshot for OCR.",
        debugImage: "",
      },
      pageUrl: pageResult.pageUrl || tab?.url || "",
      capturedAt: new Date().toISOString(),
    };
  }
  const ocr = await fetch(`${activeBridgeUrl}/ocr/cardladder`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ image, row }),
  }).then((r) => r.json()).catch((error) => ({ ok: false, error: String(error?.message || error) }));

  return {
    ...row,
    value: ocr.value ?? null,
    status: ocr.ok ? "ok" : "ocr_not_found",
    ocr,
    pageUrl: pageResult.pageUrl || tab?.url || "",
    capturedAt: new Date().toISOString(),
  };
}

async function captureActiveTabWithOcr(row) {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id || !tab.url?.startsWith("https://app.cardladder.com/")) {
    return { ...row, value: null, status: "error", error: "Open a Card Ladder results page first" };
  }
  return captureValueWithOcr(tab.id, row, { pageUrl: tab.url });
}

async function postBridgeResult(result) {
  await fetch(`${activeBridgeUrl}/result/cardladder`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(result),
  }).catch(() => {});
}

async function postBridgeFinish(payload) {
  await fetch(`${activeBridgeUrl}/finish/cardladder`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(payload),
  }).catch(() => {});
}

async function closeActiveWindow() {
  if (activeWindowId !== null) {
    const windowId = activeWindowId;
    activeWindowId = null;
    await chrome.windows.remove(windowId).catch(() => {});
  }
}

async function injectContent(tabId) {
  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["content.js"],
  }).catch(() => {});
}

async function waitForTabNotLoading(tabId) {
  for (let i = 0; i < 60; i += 1) {
    const tab = await chrome.tabs.get(tabId).catch(() => null);
    if (!tab) throw new Error("Card Ladder tab closed.");
    if (tab.status === "complete") {
      await delay(700);
      return;
    }
    await delay(1000);
  }
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

const queueFile = document.getElementById("queueFile");
const loadQueue = document.getElementById("loadQueue");
const start = document.getElementById("start");
const stop = document.getElementById("stop");
const capture = document.getElementById("capture");
const download = document.getElementById("download");
const status = document.getElementById("status");
const testGrader = document.getElementById("testGrader");
const testGraderButton = document.getElementById("testGraderButton");
const BRIDGE_PORTS = [8765, 8766, 8767, 8768, 8769, 8770, 8771, 8772];
const BRIDGE_URLS = BRIDGE_PORTS.map((port) => `http://127.0.0.1:${port}`);

let queue = null;
let results = [];
let activeBridgeUrl = BRIDGE_URLS[0];
let manualStatusUntil = 0;

loadQueue.addEventListener("click", async () => {
  const file = queueFile.files?.[0];
  if (!file) return setStatus("Choose cardladder-queue.json first.");
  queue = JSON.parse(await file.text());
  results = [];
  await chrome.runtime.sendMessage({ type: "CARDLADDER_SAVE_QUEUE", queue });
  setStatus(`Loaded ${queue.rows.length} rows from ${queue.sourceSheet}.`);
});

start.addEventListener("click", async () => {
  const response = await chrome.runtime.sendMessage({ type: "CARDLADDER_SYNC_NOW" });
  if (!response?.ok) return setStatus(response?.error || "Run failed.");
  setStatus("Card Ladder window opened. Values will be captured, piped back, and the window will close automatically.");
});

stop.addEventListener("click", async () => {
  const response = await chrome.runtime.sendMessage({ type: "CARDLADDER_CANCEL_RUN" })
    .catch((error) => ({ ok: false, error: String(error?.message || error) }));
  setStatus(response?.ok ? "Card Ladder run cancelled." : response?.error || "Cancel failed.");
});

capture.addEventListener("click", async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.url?.startsWith("https://app.cardladder.com/")) {
    return setStatus("Open the Card Ladder results page first.");
  }

  const stored = await chrome.storage.local.get(["cardladderQueue", "cardladderResults"]);
  const row = stored.cardladderQueue?.rows?.[0] || {};
  const response = await chrome.runtime.sendMessage({
    type: "CARDLADDER_CAPTURE_ACTIVE_TAB",
    row,
  }).catch((error) => ({ value: null, status: "error", error: String(error?.message || error) }));
  results = [response];
  await chrome.storage.local.set({ cardladderResults: results });
  await postBridgeResult(response);
  await postBridgeFinish({
    ok: response.value != null,
    total: 1,
    found: response.value != null ? 1 : 0,
    source: "capture-page",
  });
  await closeCurrentTabOrWindow(tab);
  setStatus(`Captured current page: ${response.value ?? response.error ?? "no value found"}`);
});

download.addEventListener("click", async () => {
  const stored = await chrome.storage.local.get(["cardladderQueue", "cardladderResults"]);
  const payload = {
    createdAt: new Date().toISOString(),
    sourceWorkbook: stored.cardladderQueue?.sourceWorkbook,
    sourceSheet: stored.cardladderQueue?.sourceSheet,
    results: stored.cardladderResults || results,
  };
  const url = URL.createObjectURL(new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" }));
  await chrome.downloads.download({
    url,
    filename: "cardladder-results.json",
    saveAs: true,
  });
});

testGraderButton.addEventListener("click", async () => {
  const grader = testGrader.value || "CGC";
  setStatus(`Testing ${grader} grader dropdown...`, 60000);
  const selected = await chrome.runtime.sendMessage({
    type: "CARDLADDER_TEST_GRADER_ACTIVE",
    grader,
  })
    .catch((error) => ({ ok: false, error: String(error?.message || error) }));
  setStatus(`Grader test ${selected?.ok ? "passed" : "failed"}:\n${formatJson(selected)}`, 60000);
});

setInterval(async () => {
  if (Date.now() < manualStatusUntil) return;
  const stored = await chrome.runtime.sendMessage({ type: "CARDLADDER_GET_STATUS" }).catch(() => null);
  const currentStatus = stored?.cardladderStatus;
  if (!currentStatus) return;
  const lines = [
    `Stage: ${currentStatus.stage}`,
    `Done: ${currentStatus.completed || 0}/${currentStatus.total || 0}`,
  ];
  if (currentStatus.current) lines.push(`Current: ${currentStatus.current.grader} ${currentStatus.current.certNumber}`);
  if (currentStatus.lastResult) lines.push(`Last: ${currentStatus.lastResult.certNumber} -> ${currentStatus.lastResult.value ?? currentStatus.lastResult.status}`);
  if (currentStatus.error) lines.push(`Error: ${currentStatus.error}`);
  status.textContent = lines.join("\n");
}, 1000);

function setStatus(message, holdMs = 0) {
  if (holdMs > 0) manualStatusUntil = Date.now() + holdMs;
  status.textContent = message;
}

function formatJson(value) {
  return JSON.stringify(value, null, 2);
}

async function postBridgeResult(result) {
  await resolveBridgeUrl();
  await fetch(`${activeBridgeUrl}/result/cardladder`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(result),
  }).catch(() => {});
}

async function postBridgeFinish(payload) {
  await resolveBridgeUrl();
  await fetch(`${activeBridgeUrl}/finish/cardladder`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(payload),
  }).catch(() => {});
}

async function resolveBridgeUrl() {
  for (const bridgeUrl of [activeBridgeUrl, ...BRIDGE_URLS.filter((url) => url !== activeBridgeUrl)]) {
    const response = await fetch(`${bridgeUrl}/status`).then((r) => r.json()).catch(() => null);
    if (response?.bridgeVersion) {
      activeBridgeUrl = bridgeUrl;
      return bridgeUrl;
    }
  }
  return activeBridgeUrl;
}

async function closeCurrentTabOrWindow(tab) {
  if (!tab?.id) return;
  const windowTabs = await chrome.tabs.query({ windowId: tab.windowId }).catch(() => []);
  if (windowTabs.length <= 1) {
    await chrome.windows.remove(tab.windowId).catch(() => {});
  } else {
    await chrome.tabs.remove(tab.id).catch(() => {});
  }
}

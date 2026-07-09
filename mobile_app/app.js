const state = {
  pin: localStorage.getItem("lucasMobilePin") || "",
  clientId: localStorage.getItem("lucasMobileClientId") || "",
  queue: [],
  lastDuplicate: null,
  sellRecord: null,
  people: [],
  connected: navigator.onLine !== false,
};
const profileMatch = window.location.pathname.match(/^\/mobile\/(team|personal)(?:\/|$)/);
const APP_BASE = profileMatch ? `/mobile/${profileMatch[1]}` : "/mobile";
const API_BASE = `${APP_BASE}/api`;
if (!state.clientId) {
  state.clientId = `mobile-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  localStorage.setItem("lucasMobileClientId", state.clientId);
}

const $ = (id) => document.getElementById(id);
const QUEUE_KEY = "lucasMobileQueue";
const CACHE_PREFIX = `${profileMatch ? profileMatch[1] : "default"}:`;
const CACHE_KEYS = {
  search: `${CACHE_PREFIX}lucasMobileLastSearch`,
  inventory: `${CACHE_PREFIX}lucasMobileInventorySnapshot`,
  profit: `${CACHE_PREFIX}lucasMobileLastProfit`,
  payouts: `${CACHE_PREFIX}lucasMobileLastPayouts`,
};
const INVENTORY_SNAPSHOT_LIMIT = 1000;
const INVENTORY_SNAPSHOT_REFRESH_MS = 5 * 60 * 1000;

function setConnectionStatus(connected, message = "") {
  state.connected = Boolean(connected);
  const banner = $("connectionBanner");
  if (!banner) return;
  banner.classList.toggle("hidden", connected && !message);
  banner.classList.toggle("online", connected);
  if (connected) {
    banner.textContent = message || "Connected to desktop LUCAS.";
  } else {
    banner.textContent = message || "Offline mode: desktop LUCAS is not reachable. Adds, expenses, and cached-card sales can be queued and synced later.";
  }
}

function cacheSet(key, payload) {
  try {
    localStorage.setItem(key, JSON.stringify({ saved_at: new Date().toISOString(), payload }));
  } catch (_error) {
    // Local storage can be full or disabled; offline queue still handles writes separately.
  }
}

function cacheGet(key) {
  try {
    const wrapper = JSON.parse(localStorage.getItem(key) || "null");
    return wrapper && wrapper.payload ? wrapper : null;
  } catch (_error) {
    return null;
  }
}

function cacheAgeText(savedAt) {
  if (!savedAt) return "cached";
  const ageMs = Date.now() - new Date(savedAt).getTime();
  if (!Number.isFinite(ageMs) || ageMs < 0) return "cached";
  const minutes = Math.max(1, Math.round(ageMs / 60000));
  if (minutes < 60) return `${minutes} min old`;
  const hours = Math.round(minutes / 60);
  return `${hours} hr old`;
}

function cacheIsFresh(wrapper, maxAgeMs) {
  if (!wrapper || !wrapper.saved_at) return false;
  const ageMs = Date.now() - new Date(wrapper.saved_at).getTime();
  return Number.isFinite(ageMs) && ageMs >= 0 && ageMs < maxAgeMs;
}

function setUnlocked(unlocked) {
  $("authPanel").classList.toggle("hidden", unlocked);
  $("appPanel").classList.toggle("hidden", !unlocked);
}

async function api(path, body) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 9000);
  try {
    const response = await fetch(`${API_BASE}${path}`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ ...body, pin: state.pin }),
      signal: controller.signal,
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(result.error || `Request failed with ${response.status}`);
    }
    setConnectionStatus(true);
    return result;
  } catch (error) {
    setConnectionStatus(false);
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

function loadQueue() {
  try {
    const raw = JSON.parse(localStorage.getItem(QUEUE_KEY) || "[]");
    state.queue = Array.isArray(raw) ? raw.filter((item) => item && item.id && item.type) : [];
  } catch (_error) {
    state.queue = [];
  }
}

function saveQueue() {
  localStorage.setItem(QUEUE_KEY, JSON.stringify(state.queue));
  renderQueue();
}

function queueAction(type, payload) {
  const id = `${state.clientId}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const action = {
    id,
    type,
    client_id: state.clientId,
    created_at: new Date().toISOString(),
    payload: { ...payload },
  };
  delete action.payload.pin;
  state.queue.push(action);
  saveQueue();
  return action;
}

async function mutationApi(type, path, body) {
  try {
    return await api(path, body);
  } catch (error) {
    const message = String(error && error.message ? error.message : error || "");
    if (/pin/i.test(message)) {
      return { ok: false, error: message || "Invalid mobile PIN." };
    }
    const action = queueAction(type, body);
    return {
      ok: true,
      queued: true,
      action_id: action.id,
      error: message,
    };
  }
}

function money(value, fallback = "") {
  if (value === null || value === undefined || value === "") return fallback;
  const number = Number(value);
  if (!Number.isFinite(number)) return String(value);
  return number.toLocaleString(undefined, { style: "currency", currency: "USD" });
}

function fillSelect(select, values, options = {}) {
  const current = select.value;
  const allLabel = options.allLabel || "All people";
  const includeAll = options.includeAll !== false;
  const choices = includeAll ? [`<option value="">${escapeHtml(allLabel)}</option>`] : [];
  values.forEach((value) => choices.push(`<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`));
  select.innerHTML = choices.join("");
  if ([...select.options].some((option) => option.value === current)) {
    select.value = current;
  }
}

function updatePeople(people) {
  if (Array.isArray(people)) {
    state.people = people.filter(Boolean);
  }
  ["personFilter", "profitPerson", "payoutPerson"].forEach((id) => fillSelect($(id), state.people));
}

function queueTitle(action) {
  const payload = action.payload || {};
  if (action.type === "inventory.add") {
    return payload.cert_number || payload.card_title || "Inventory add";
  }
  if (action.type === "inventory.sold") {
    return payload.inventory_key || payload.company || "Inventory sale";
  }
  if (action.type === "expense.add") {
    return `${payload.person || payload.assigned_person || "Expense"} ${money(payload.amount || payload.expense_amount, "")}`.trim();
  }
  return action.type || "Queued action";
}

function renderQueue() {
  const badge = $("queueBadge");
  if (badge) badge.textContent = String(state.queue.length);
  const host = $("queueList");
  if (!host) return;
  if (!state.queue.length) {
    host.innerHTML = '<div class="hint">No queued mobile actions.</div>';
    return;
  }
  host.innerHTML = state.queue.map((action) => `
    <article class="result">
      <div class="queueType">${escapeHtml(action.type)}</div>
      <h2>${escapeHtml(queueTitle(action))}</h2>
      <div class="meta">
        <div><strong>Created</strong>${escapeHtml(String(action.created_at || "").replace("T", " ").slice(0, 19))}</div>
        <div><strong>Action ID</strong>${escapeHtml(action.id)}</div>
      </div>
    </article>
  `).join("");
}

async function syncQueuedActions() {
  if (!state.queue.length) {
    $("syncStatus").textContent = "No queued actions to sync.";
    renderQueue();
    return;
  }
  $("syncStatus").textContent = "Syncing queued actions...";
  try {
    const result = await api("/sync/queue", {
      client_id: state.clientId,
      actions: state.queue,
    });
    const completed = new Set(
      (result.results || [])
        .filter((item) => item && item.ok && ["applied", "already_applied"].includes(item.status))
        .map((item) => item.id)
    );
    state.queue = state.queue.filter((action) => !completed.has(action.id));
    saveQueue();
    $("syncStatus").textContent = `Applied ${result.applied || 0}, skipped ${result.skipped || 0}, failed ${result.failed || 0}.`;
    if (completed.size) {
      searchInventory();
      loadProfit();
      loadPayouts();
    }
  } catch (error) {
    setConnectionStatus(false);
    $("syncStatus").textContent = `Desktop not reachable. Export the queue or try again later. ${error.message || error}`;
    renderQueue();
  }
}

function exportQueue() {
  const payload = {
    version: 1,
    service: "lucas-mobile-offline-queue",
    client_id: state.clientId,
    exported_at: new Date().toISOString(),
    actions: state.queue,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `lucas-mobile-queue-${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
  $("syncStatus").textContent = `Exported ${state.queue.length} queued action(s).`;
}

function syncPersonInputs(person) {
  const value = person || $("personFilter").value || $("profitPerson").value || $("payoutPerson").value || "";
  if (value && !$("assignedPerson").value) $("assignedPerson").value = value;
  if (value && !$("expensePerson").value) $("expensePerson").value = value;
}

function photoFileName(photo, record) {
  const fallback = `${record.cert_number || record.item_id || "lucas-card"}.jpg`;
  return String(photo.name || fallback).replace(/[^\w.\-()[\] ]+/g, "-") || fallback;
}

async function shareInventoryPhoto(record, photo) {
  const status = $("scanSearchStatus");
  const url = photo && photo.url;
  if (!url) {
    status.textContent = "No attached photo is available for that card.";
    return;
  }
  status.textContent = "Preparing photo...";
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Photo request failed with ${response.status}`);
    const blob = await response.blob();
    const file = new File([blob], photoFileName(photo, record), { type: blob.type || "image/jpeg" });
    const title = record.card_title || record.cert_number || "LUCAS inventory photo";
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ title, text: title, files: [file] });
      status.textContent = "Photo shared.";
      return;
    }
    const objectUrl = URL.createObjectURL(blob);
    window.open(objectUrl, "_blank");
    setTimeout(() => URL.revokeObjectURL(objectUrl), 60000);
    status.textContent = "Opened photo. Use Share from Safari if needed.";
  } catch (error) {
    if (error && error.name === "AbortError") {
      status.textContent = "Share cancelled.";
      return;
    }
    window.open(url, "_blank");
    status.textContent = `Opened photo fallback. ${error.message || error}`;
  }
}

function openInventoryPhoto(record, photo) {
  const viewer = $("photoViewer");
  const image = $("photoViewerImage");
  const title = $("photoViewerTitle");
  const url = photo && photo.url;
  if (!viewer || !image || !url) return;
  image.src = url;
  image.alt = record.card_title || record.cert_number || "Inventory photo";
  if (title) title.textContent = record.card_title || record.cert_number || photo.name || "Inventory photo";
  viewer.classList.remove("hidden");
  document.body.classList.add("photoViewerOpen");
}

function closeInventoryPhoto() {
  const viewer = $("photoViewer");
  const image = $("photoViewerImage");
  if (!viewer || viewer.classList.contains("hidden")) return;
  viewer.classList.add("hidden");
  document.body.classList.remove("photoViewerOpen");
  if (image) image.removeAttribute("src");
}

function renderCachedSearch(wrapper, error) {
  const snapshot = cacheGet(CACHE_KEYS.inventory);
  const source = snapshot || wrapper;
  if (!source || !source.payload) {
    $("results").innerHTML = `<div class="hint">Desktop LUCAS is not reachable and no cached inventory search is saved on this phone yet. Add inventory and expenses can still be queued. ${escapeHtml(error?.message || error || "")}</div>`;
    return;
  }
  const result = filterCachedInventory(source.payload);
  updatePeople(result.people || state.people);
  renderResults(result.items || []);
  const label = snapshot ? "inventory snapshot" : "inventory search";
  const note = `Showing cached ${label} (${cacheAgeText(source.saved_at)}). Live search needs desktop LUCAS online.`;
  $("scanSearchStatus").textContent = note;
  setConnectionStatus(false, note);
}

function renderPhotoStrip(item, itemIndex) {
  const photos = Array.isArray(item.photos) ? item.photos : [];
  if (!photos.length) {
    return '<div class="photoHint">No photo attached</div>';
  }
  return `
    <div class="photoStrip" aria-label="Attached inventory photos">
      ${photos.map((photo, photoIndex) => `
        <div class="photoTile">
          <button class="photoOpenButton" data-index="${itemIndex}" data-photo-index="${photoIndex}" type="button" aria-label="Open attached photo ${photoIndex + 1}">
            <img src="${escapeHtml(photo.url || "")}" alt="${escapeHtml(item.card_title || item.cert_number || "Inventory photo")}" loading="lazy">
          </button>
          <button class="secondary sharePhotoButton" data-index="${itemIndex}" data-photo-index="${photoIndex}" type="button">Share Photo</button>
        </div>
      `).join("")}
    </div>
  `;
}

function renderResults(items) {
  const host = $("results");
  if (!items.length) {
    host.innerHTML = '<div class="hint">No inventory matched.</div>';
    return;
  }
  host.innerHTML = items.map((item, index) => `
    <article class="result">
      <h2>${escapeHtml(item.card_title || item.cert_number || "Untitled card")}</h2>
      <div class="meta">
        <div><strong>Cert</strong>${escapeHtml(item.cert_number || "")}</div>
        <div><strong>Item ID</strong>${escapeHtml(item.item_id || "")}</div>
        <div><strong>Grader</strong>${escapeHtml(item.grader || "")}</div>
        <div><strong>Paid</strong>${escapeHtml(item.purchase_price_display || money(item.purchase_price, "-"))}</div>
        <div><strong>Value</strong>${escapeHtml(item.inventory_value_display || money(item.inventory_value, "-"))}</div>
        <div><strong>Company</strong>${escapeHtml(item.best_company || "-")}</div>
        <div><strong>Payout</strong>${escapeHtml(item.estimated_payout_display || money(item.estimated_payout, "-"))}</div>
        <div><strong>Person</strong>${escapeHtml(item.assigned_person || "-")}</div>
        <div><strong>Source</strong>${escapeHtml(item.source || item.source_sheet || "-")}</div>
      </div>
      ${renderPhotoStrip(item, index)}
      ${String(item.status || "").toLowerCase() === "active" ? `<div class="resultActions"><button class="secondary sellButton" data-index="${index}" type="button">Mark Sold</button></div>` : ""}
    </article>
  `).join("");
  document.querySelectorAll(".sellButton").forEach((button) => {
    button.addEventListener("click", () => startSell(items[Number(button.dataset.index)]));
  });
  document.querySelectorAll(".sharePhotoButton").forEach((button) => {
    button.addEventListener("click", () => {
      const item = items[Number(button.dataset.index)];
      const photos = Array.isArray(item && item.photos) ? item.photos : [];
      shareInventoryPhoto(item, photos[Number(button.dataset.photoIndex)]);
    });
  });
  document.querySelectorAll(".photoOpenButton").forEach((button) => {
    button.addEventListener("click", () => {
      const item = items[Number(button.dataset.index)];
      const photos = Array.isArray(item && item.photos) ? item.photos : [];
      openInventoryPhoto(item, photos[Number(button.dataset.photoIndex)]);
    });
  });
}

function selectedCategories() {
  return Array.from(document.querySelectorAll(".categoryFilter:checked"))
    .map((input) => input.value)
    .filter(Boolean);
}

function inventorySearchPayload(overrides = {}) {
  return {
    query: $("searchInput").value,
    person: $("personFilter").value,
    sport: selectedCategories(),
    include_sold: $("includeSold").checked,
    ...overrides,
  };
}

function cachedInventoryMatches(item, payload) {
  const query = String(payload.query || "").trim().toLowerCase();
  const person = String(payload.person || "").trim().toLowerCase();
  const sports = Array.isArray(payload.sport) ? payload.sport : [payload.sport].filter(Boolean);
  const sportFilters = sports.map((value) => String(value || "").trim().toLowerCase()).filter(Boolean);
  const status = String(item.status || "").toLowerCase();
  if (status !== "active" && !payload.include_sold) return false;
  if (person && !String(item.assigned_person || "Unassigned").toLowerCase().includes(person)) return false;
  if (sportFilters.length) {
    const sportText = String(item.sport || "").toLowerCase();
    if (!sportFilters.some((sport) => sportText === sport || sportText.includes(sport))) return false;
  }
  if (query) {
    const haystack = [
      item.inventory_key,
      item.item_type,
      item.item_id,
      item.cert_number,
      item.card_title,
      item.grader,
      item.assigned_person,
      item.sport,
      item.source,
      item.best_company,
      item.notes,
    ].map((value) => String(value || "").toLowerCase()).join(" ");
    if (query.split(/\s+/).some((part) => part && !haystack.includes(part))) return false;
  }
  return true;
}

function filterCachedInventory(payload) {
  const items = Array.isArray(payload.items) ? payload.items : [];
  const search = inventorySearchPayload();
  const filtered = items.filter((item) => cachedInventoryMatches(item, search)).slice(0, 75);
  return {
    ...payload,
    count: filtered.length,
    items: filtered,
    people: payload.people || state.people,
  };
}

async function refreshInventorySnapshot(force = false) {
  if (state.snapshotRefreshInFlight) return;
  if (!force && cacheIsFresh(cacheGet(CACHE_KEYS.inventory), INVENTORY_SNAPSHOT_REFRESH_MS)) return;
  state.snapshotRefreshInFlight = true;
  try {
    const result = await api("/inventory/search", {
      query: "",
      person: "",
      sport: [],
      include_sold: true,
      limit: INVENTORY_SNAPSHOT_LIMIT,
    });
    if (result && result.ok) cacheSet(CACHE_KEYS.inventory, result);
  } catch (_error) {
    // The visible search path already handles offline messaging.
  } finally {
    state.snapshotRefreshInFlight = false;
  }
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  }[char]));
}

async function searchInventory() {
  let result;
  if (navigator.onLine === false) {
    renderCachedSearch(cacheGet(CACHE_KEYS.search), "Phone is offline.");
    return;
  }
  try {
    result = await api("/inventory/search", inventorySearchPayload());
  } catch (error) {
    renderCachedSearch(cacheGet(CACHE_KEYS.search), error);
    return;
  }
  if (!result.ok) {
    if (/pin/i.test(result.error || "")) setUnlocked(false);
    $("results").innerHTML = `<div class="hint">${escapeHtml(result.error || "Search failed.")}</div>`;
    return;
  }
  updatePeople(result.people || []);
  cacheSet(CACHE_KEYS.search, result);
  $("scanSearchStatus").textContent = "";
  renderResults(result.items || []);
  refreshInventorySnapshot();
}

function clearCategoryFilters() {
  document.querySelectorAll(".categoryFilter").forEach((input) => {
    input.checked = false;
  });
  searchInventory();
}

function addPayload(updateExisting = false) {
  return {
    cert_number: $("certNumber").value,
    grader: $("grader").value,
    card_title: $("cardTitle").value,
    purchase_price: $("purchasePrice").value,
    assigned_person: $("assignedPerson").value,
    source: $("source").value,
    inventory_value: $("inventoryValue").value,
    notes: $("notes").value,
    update_existing: updateExisting,
  };
}

async function addInventory(updateExisting = false) {
  $("addStatus").textContent = "Saving...";
  $("updateDuplicate").classList.add("hidden");
  const result = await mutationApi("inventory.add", "/inventory/add", addPayload(updateExisting));
  if (result.queued) {
    $("addStatus").textContent = `Desktop not reachable. Queued inventory add ${result.action_id}.`;
    return;
  }
  if (result.duplicate) {
    state.lastDuplicate = result.record;
    $("addStatus").textContent = result.error || "Duplicate cert found.";
    $("updateDuplicate").classList.remove("hidden");
    return;
  }
  if (!result.ok) {
    if (/pin/i.test(result.error || "")) setUnlocked(false);
    $("addStatus").textContent = result.error || "Add failed.";
    return;
  }
  $("addStatus").textContent = `${result.action === "updated" ? "Updated" : "Added"} ${result.record?.cert_number || result.record?.card_title || "card"}.`;
  updatePeople(result.people || state.people);
  $("searchInput").value = result.record?.cert_number || result.record?.card_title || "";
  searchInventory();
}

function startSell(record) {
  state.sellRecord = record || null;
  if (!state.sellRecord) return;
  $("sellTitle").textContent = `Mark Sold: ${state.sellRecord.cert_number || state.sellRecord.card_title || "card"}`;
  $("sellPrice").value = state.sellRecord.estimated_payout || state.sellRecord.inventory_value || state.sellRecord.purchase_price || "";
  $("sellDate").value = new Date().toISOString().slice(0, 10);
  $("sellCompany").value = "";
  $("sellStatus").textContent = "";
  $("sellPanel").classList.remove("hidden");
  $("sellPrice").focus();
}

function cancelSell() {
  state.sellRecord = null;
  $("sellPanel").classList.add("hidden");
  $("sellStatus").textContent = "";
}

async function confirmSell() {
  if (!state.sellRecord) return;
  $("sellStatus").textContent = "Saving sale...";
  const result = await mutationApi("inventory.sold", "/inventory/sold", {
    inventory_key: state.sellRecord.inventory_key,
    sale_price: $("sellPrice").value,
    sale_date: $("sellDate").value,
    sale_method: $("sellMethod").value,
    company: $("sellCompany").value,
  });
  if (result.queued) {
    $("sellStatus").textContent = `Desktop not reachable. Queued sale ${result.action_id}.`;
    cancelSell();
    return;
  }
  if (!result.ok) {
    if (/pin/i.test(result.error || "")) setUnlocked(false);
    $("sellStatus").textContent = result.error || "Could not mark sold.";
    return;
  }
  $("sellStatus").textContent = `Sold for ${result.sale?.sale_price_display || money(result.sale?.sale_price, "")}.`;
  cancelSell();
  $("searchInput").value = "";
  await searchInventory();
  loadProfit();
  loadPayouts();
}

function expensePayload() {
  return {
    person: $("expensePerson").value,
    date: $("expenseDate").value,
    expense_type: $("expenseType").value,
    amount: $("expenseAmount").value,
    related_type: $("expenseRelatedType").value,
    source_sheet: $("expenseSheet").value,
    notes: $("expenseNotes").value,
  };
}

async function addExpense() {
  $("expenseStatus").textContent = "Saving...";
  const result = await mutationApi("expense.add", "/expenses/add", expensePayload());
  if (result.queued) {
    $("expenseStatus").textContent = `Desktop not reachable. Queued expense ${result.action_id}.`;
    $("expenseAmount").value = "";
    $("expenseSheet").value = "";
    $("expenseNotes").value = "";
    return;
  }
  if (!result.ok) {
    if (/pin/i.test(result.error || "")) setUnlocked(false);
    $("expenseStatus").textContent = result.error || "Expense add failed.";
    return;
  }
  updatePeople(result.people || state.people);
  $("expenseStatus").textContent = "Expense added.";
  $("expenseAmount").value = "";
  $("expenseSheet").value = "";
  $("expenseNotes").value = "";
  loadProfit();
}

function renderMetrics(host, items) {
  host.innerHTML = items.map((item) => `
    <div class="metric">
      <strong>${escapeHtml(item.label)}</strong>
      <span>${escapeHtml(item.value)}</span>
    </div>
  `).join("");
}

function drawChart(chart) {
  const svg = $("profitChart");
  const values = (chart && chart.values) || [];
  if (!values.length) {
    svg.innerHTML = '<text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" fill="#9eb1ba">No profit data</text>';
    return;
  }
  const width = 700;
  const height = 220;
  const pad = 24;
  const min = Math.min(...values, 0);
  const max = Math.max(...values, 0);
  const span = max - min || 1;
  const point = (value, index) => {
    const x = pad + (values.length === 1 ? 0 : (index / (values.length - 1)) * (width - pad * 2));
    const y = height - pad - ((value - min) / span) * (height - pad * 2);
    return `${x.toFixed(1)},${y.toFixed(1)}`;
  };
  const points = values.map(point).join(" ");
  const zeroY = height - pad - ((0 - min) / span) * (height - pad * 2);
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.innerHTML = `
    <line x1="${pad}" y1="${zeroY.toFixed(1)}" x2="${width - pad}" y2="${zeroY.toFixed(1)}" stroke="#304453" stroke-width="1" />
    <polyline points="${points}" fill="none" stroke="#7ed8bd" stroke-width="4" stroke-linecap="round" stroke-linejoin="round" />
    <circle cx="${point(values[values.length - 1], values.length - 1).split(",")[0]}" cy="${point(values[values.length - 1], values.length - 1).split(",")[1]}" r="5" fill="#eef5f6" />
    <text x="${pad}" y="18" fill="#9eb1ba" font-size="13">${escapeHtml(money(max))}</text>
    <text x="${pad}" y="${height - 7}" fill="#9eb1ba" font-size="13">${escapeHtml(money(min))}</text>
  `;
}

async function loadProfit() {
  let result;
  try {
    result = await api("/profit/summary", {
      person: $("profitPerson").value,
      period: $("profitPeriod").value,
      graph: $("profitGraph").value,
    });
  } catch (error) {
    const cached = cacheGet(CACHE_KEYS.profit);
    if (cached && cached.payload) {
      renderProfit(cached.payload);
      $("profitRecent").insertAdjacentHTML("afterbegin", `<div class="hint">Showing cached profit (${escapeHtml(cacheAgeText(cached.saved_at))}). Live profit needs desktop LUCAS online.</div>`);
    } else {
      $("profitRecent").innerHTML = `<div class="hint">Desktop LUCAS is not reachable and no cached profit is saved on this phone yet. ${escapeHtml(error.message || error)}</div>`;
    }
    return;
  }
  if (!result.ok) {
    if (/pin/i.test(result.error || "")) setUnlocked(false);
    $("profitRecent").innerHTML = `<div class="hint">${escapeHtml(result.error || "Profit load failed.")}</div>`;
    return;
  }
  cacheSet(CACHE_KEYS.profit, result);
  renderProfit(result);
}

function renderProfit(result) {
  updatePeople(result.people || []);
  fillSelect($("profitPeriod"), result.periods || ["Total", "Year", "Month", "Week", "5 Days"], { includeAll: false });
  fillSelect($("profitGraph"), result.graphs || ["Daily Trend", "Overall Profit"], { includeAll: false });
  const totals = result.totals || {};
  renderMetrics($("profitCards"), [
    { label: "Net Profit", value: money(totals.net_profit, "$0.00") },
    { label: "Gross Profit", value: money(totals.gross_profit, "$0.00") },
    { label: "Expenses", value: money(totals.expenses, "$0.00") },
    { label: "Sales", value: money(totals.sale, "$0.00") },
  ]);
  drawChart(result.chart || {});
  const recent = result.recent || [];
  $("profitRecent").innerHTML = recent.length ? recent.map((item) => `
    <article class="result">
      <h2>${escapeHtml(item.title || item.company || item.type)}</h2>
      <div class="meta">
        <div><strong>Date</strong>${escapeHtml(item.date || "")}</div>
        <div><strong>Person</strong>${escapeHtml(item.person || "")}</div>
        <div><strong>Type</strong>${escapeHtml(item.type || "")}</div>
        <div><strong>Profit</strong>${escapeHtml(item.profit_display || money(item.profit, "-"))}</div>
      </div>
    </article>
  `).join("") : '<div class="hint">No profit rows matched.</div>';
}

async function loadPayouts() {
  let result;
  try {
    result = await api("/payouts", { person: $("payoutPerson").value });
  } catch (error) {
    const cached = cacheGet(CACHE_KEYS.payouts);
    if (cached && cached.payload) {
      renderPayouts(cached.payload);
      $("payoutSummary").insertAdjacentHTML("afterbegin", `<div class="hint">Showing cached payouts (${escapeHtml(cacheAgeText(cached.saved_at))}). Live payouts need desktop LUCAS online.</div>`);
    } else {
      $("payoutSummary").innerHTML = `<div class="hint">Desktop LUCAS is not reachable and no cached payout data is saved on this phone yet. ${escapeHtml(error.message || error)}</div>`;
    }
    return;
  }
  if (!result.ok) {
    if (/pin/i.test(result.error || "")) setUnlocked(false);
    $("payoutSummary").innerHTML = `<div class="hint">${escapeHtml(result.error || "Payout load failed.")}</div>`;
    return;
  }
  cacheSet(CACHE_KEYS.payouts, result);
  renderPayouts(result);
}

function renderPayouts(result) {
  updatePeople(result.people || []);
  const totals = result.totals || {};
  renderMetrics($("payoutCards"), [
    { label: "Balance", value: totals.balance_display || money(totals.balance, "$0.00") },
    { label: "Sheets", value: String(totals.sheets || 0) },
    { label: "Cards", value: String(totals.cards || 0) },
    { label: "Mode", value: "View only" },
  ]);
  const summary = result.summary || [];
  $("payoutSummary").innerHTML = summary.length ? summary.map((item) => `
    <article class="result">
      <h2>${escapeHtml(item.person || "Unassigned")}</h2>
      <div class="meta">
        <div><strong>Sheets</strong>${escapeHtml(item.sheets || 0)}</div>
        <div><strong>Cards</strong>${escapeHtml(item.cards || 0)}</div>
        <div><strong>Balance</strong>${escapeHtml(item.balance_display || money(item.balance, "-"))}</div>
      </div>
    </article>
  `).join("") : '<div class="hint">No active payout balances matched.</div>';
  const details = result.details || [];
  $("payoutDetails").innerHTML = details.length ? details.slice(0, 30).map((item) => `
    <article class="result">
      <h2>${escapeHtml(item.name || "Sheet")}</h2>
      <div class="meta">
        <div><strong>Person</strong>${escapeHtml(item.person || "")}</div>
        <div><strong>Status</strong>${escapeHtml(item.status || "")}</div>
        <div><strong>Received</strong>${escapeHtml(`${item.received_count || 0}/${item.row_count || 0}`)}</div>
        <div><strong>Balance</strong>${escapeHtml(item.payout_balance_display || money(item.payout_balance, "-"))}</div>
      </div>
    </article>
  `).join("") : "";
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function identifyPhoto(file, target) {
  const status = target === "certNumber" ? $("scanAddStatus") : $("scanSearchStatus");
  status.textContent = "Reading card photo...";
  try {
    const image = await fileToDataUrl(file);
    const result = await api("/card/identify", { image });
    if (!result.ok) {
      status.textContent = result.error || "Could not read that card.";
      return;
    }
    const card = result.card || {};
    const query = result.query || card.cert_number || card.card_title || "";
    status.textContent = `Found ${card.cert_number || card.card_title || "card"}.`;
    if (target === "certNumber") {
      $("certNumber").value = card.cert_number || "";
      if (card.grader) $("grader").value = card.grader;
      if (card.card_title) $("cardTitle").value = card.card_title;
      if (card.notes && !$("notes").value) $("notes").value = card.notes;
      if (!$("source").value) $("source").value = "Mobile Photo";
      status.textContent = card.cert_number
        ? `Found cert ${card.cert_number}. Add purchase price, then tap Add Inventory.`
        : "Found card details. Review fields, add purchase price, then tap Add Inventory.";
    } else {
      $("searchInput").value = query;
      await searchInventory();
    }
  } catch (error) {
    setConnectionStatus(false);
    status.textContent = `Photo OCR needs desktop LUCAS online. You can still type the card info and queue the add. ${error.message || error}`;
  }
}

function bindPhotoInput(inputId, targetId) {
  $(inputId).addEventListener("change", () => {
    const input = $(inputId);
    const file = input.files && input.files[0];
    if (file) identifyPhoto(file, targetId);
    input.value = "";
  });
}

function bind() {
  loadQueue();
  $("pin").value = state.pin;
  $("expenseDate").value = new Date().toISOString().slice(0, 10);
  $("sellDate").value = new Date().toISOString().slice(0, 10);
  fillSelect($("sellMethod"), ["Cash", "Wire", "Venmo", "Zelle", "PayPal", "Check", "Trade", "Other"], { includeAll: false });
  fillSelect($("expenseType"), ["Travel", "Supplies", "Travel Meal", "Fees", "Shipping"], { includeAll: false });
  fillSelect($("expenseRelatedType"), ["General", "Card", "Sheet"], { includeAll: false });
  fillSelect($("profitPeriod"), ["Total", "Year", "Month", "Week", "5 Days"], { includeAll: false });
  fillSelect($("profitGraph"), ["Daily Trend", "Overall Profit"], { includeAll: false });
  updatePeople([]);
  renderQueue();
  setUnlocked(Boolean(state.pin));
  $("savePin").addEventListener("click", () => {
    state.pin = $("pin").value.trim();
    localStorage.setItem("lucasMobilePin", state.pin);
    setUnlocked(Boolean(state.pin));
    searchInventory();
  });
  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((item) => item.classList.remove("active"));
      button.classList.add("active");
      ["search", "add", "expense", "profit", "payout", "sync"].forEach((view) => {
        $(`${view}View`).classList.toggle("hidden", button.dataset.view !== view);
      });
      if (button.dataset.view === "profit") loadProfit();
      if (button.dataset.view === "payout") loadPayouts();
      if (button.dataset.view === "sync") renderQueue();
    });
  });
  $("searchInput").addEventListener("input", () => searchInventory());
  $("personFilter").addEventListener("change", () => {
    syncPersonInputs($("personFilter").value);
    searchInventory();
  });
  document.querySelectorAll(".categoryFilter").forEach((input) => {
    input.addEventListener("change", () => searchInventory());
  });
  $("clearCategoryFilters").addEventListener("click", () => clearCategoryFilters());
  $("includeSold").addEventListener("change", () => searchInventory());
  $("cancelSell").addEventListener("click", () => cancelSell());
  $("confirmSell").addEventListener("click", () => confirmSell());
  $("profitPerson").addEventListener("change", () => loadProfit());
  $("profitPeriod").addEventListener("change", () => loadProfit());
  $("profitGraph").addEventListener("change", () => loadProfit());
  $("payoutPerson").addEventListener("change", () => loadPayouts());
  $("syncQueue").addEventListener("click", () => syncQueuedActions());
  $("exportQueue").addEventListener("click", () => exportQueue());
  $("clearAppliedQueue").addEventListener("click", () => {
    if (state.queue.length) {
      $("syncStatus").textContent = "Queue still has pending actions. Sync or export it first.";
      return;
    }
    localStorage.removeItem(QUEUE_KEY);
    loadQueue();
    renderQueue();
    $("syncStatus").textContent = "Queue is empty.";
  });
  $("closePhotoViewer").addEventListener("click", () => closeInventoryPhoto());
  $("photoViewer").addEventListener("click", (event) => {
    if (event.target === $("photoViewer") || event.target.classList.contains("photoViewerCanvas")) {
      closeInventoryPhoto();
    }
  });
  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeInventoryPhoto();
  });
  window.addEventListener("online", () => {
    setConnectionStatus(true, "Network is back. Sync queued actions when desktop LUCAS is open.");
    if (state.pin && state.queue.length) syncQueuedActions();
  });
  window.addEventListener("offline", () => setConnectionStatus(false));
  bindPhotoInput("photoSearchInput", "searchInput");
  bindPhotoInput("photoAddInput", "certNumber");
  $("addInventory").addEventListener("click", () => addInventory(false));
  $("updateDuplicate").addEventListener("click", () => addInventory(true));
  $("addExpense").addEventListener("click", () => addExpense());
  $("installHelp").addEventListener("click", () => alert("On iPhone: Share -> Add to Home Screen."));
  setConnectionStatus(navigator.onLine !== false, navigator.onLine === false ? "" : "Ready. Live data loads when desktop LUCAS is reachable.");
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register(`${APP_BASE}/sw.js`, { scope: `${APP_BASE}/` }).catch(() => {});
  }
  if (state.pin) {
    const cached = cacheGet(CACHE_KEYS.search);
    if (cached && cached.payload) {
      renderCachedSearch(cached);
    }
    searchInventory();
  }
}

bind();

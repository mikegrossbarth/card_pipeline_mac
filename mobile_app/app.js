const state = {
  pin: localStorage.getItem("lucasMobilePin") || "",
  lastDuplicate: null,
  photoTarget: null,
  stream: null,
};

const $ = (id) => document.getElementById(id);

function setUnlocked(unlocked) {
  $("authPanel").classList.toggle("hidden", unlocked);
  $("appPanel").classList.toggle("hidden", !unlocked);
}

function api(path, body) {
  return fetch(path, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ ...body, pin: state.pin }),
  }).then((response) => response.json());
}

function money(value, fallback = "") {
  if (value === null || value === undefined || value === "") return fallback;
  const number = Number(value);
  if (!Number.isFinite(number)) return String(value);
  return number.toLocaleString(undefined, { style: "currency", currency: "USD" });
}

function renderResults(items) {
  const host = $("results");
  if (!items.length) {
    host.innerHTML = '<div class="hint">No inventory matched.</div>';
    return;
  }
  host.innerHTML = items.map((item) => `
    <article class="result">
      <h2>${escapeHtml(item.card_title || item.cert_number || "Untitled card")}</h2>
      <div class="meta">
        <div><strong>Cert</strong>${escapeHtml(item.cert_number || "")}</div>
        <div><strong>Grader</strong>${escapeHtml(item.grader || "")}</div>
        <div><strong>Paid</strong>${escapeHtml(item.purchase_price_display || money(item.purchase_price, "-"))}</div>
        <div><strong>Value</strong>${escapeHtml(item.inventory_value_display || money(item.inventory_value, "-"))}</div>
        <div><strong>Company</strong>${escapeHtml(item.best_company || "-")}</div>
        <div><strong>Payout</strong>${escapeHtml(item.estimated_payout_display || money(item.estimated_payout, "-"))}</div>
        <div><strong>Person</strong>${escapeHtml(item.assigned_person || "-")}</div>
        <div><strong>Source</strong>${escapeHtml(item.source || item.source_sheet || "-")}</div>
      </div>
    </article>
  `).join("");
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
  const result = await api("/mobile/api/inventory/search", {
    query: $("searchInput").value,
    include_sold: $("includeSold").checked,
  });
  if (!result.ok) {
    if (/pin/i.test(result.error || "")) setUnlocked(false);
    $("results").innerHTML = `<div class="hint">${escapeHtml(result.error || "Search failed.")}</div>`;
    return;
  }
  renderResults(result.items || []);
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
  const result = await api("/mobile/api/inventory/add", addPayload(updateExisting));
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
  $("searchInput").value = result.record?.cert_number || result.record?.card_title || "";
  searchInventory();
}

function openPhotoScan(targetId) {
  state.photoTarget = targetId;
  const input = $("photoInput");
  input.value = "";
  input.click();
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function identifyPhoto(file) {
  const target = state.photoTarget || "searchInput";
  const status = target === "certNumber" ? $("scanAddStatus") : $("scanSearchStatus");
  status.textContent = "Reading card photo...";
  try {
    const image = await fileToDataUrl(file);
    const result = await api("/mobile/api/card/identify", { image });
    if (!result.ok) {
      status.textContent = result.error || "Could not read that card.";
      return;
    }
    const card = result.card || {};
    const query = result.query || card.cert_number || card.card_title || "";
    status.textContent = `Found ${card.cert_number || card.card_title || "card"}.`;
    if (target === "certNumber") {
      $("certNumber").value = card.cert_number || query;
      if (card.grader) $("grader").value = card.grader;
      if (card.card_title) $("cardTitle").value = card.card_title;
      if (card.notes && !$("notes").value) $("notes").value = card.notes;
    } else {
      $("searchInput").value = query;
      await searchInventory();
    }
  } catch (error) {
    status.textContent = `Photo search failed: ${error.message || error}`;
  }
}

function closeScanner() {
  if (state.stream) {
    state.stream.getTracks().forEach((track) => track.stop());
    state.stream = null;
  }
  $("scannerVideo").srcObject = null;
  if ($("scannerDialog").open) $("scannerDialog").close();
}

function bind() {
  $("pin").value = state.pin;
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
      $("searchView").classList.toggle("hidden", button.dataset.view !== "search");
      $("addView").classList.toggle("hidden", button.dataset.view !== "add");
    });
  });
  $("searchInput").addEventListener("input", () => searchInventory());
  $("includeSold").addEventListener("change", () => searchInventory());
  $("scanSearch").addEventListener("click", () => openPhotoScan("searchInput"));
  $("scanAdd").addEventListener("click", () => openPhotoScan("certNumber"));
  $("photoInput").addEventListener("change", () => {
    const file = $("photoInput").files && $("photoInput").files[0];
    if (file) identifyPhoto(file);
  });
  $("closeScanner").addEventListener("click", closeScanner);
  $("addInventory").addEventListener("click", () => addInventory(false));
  $("updateDuplicate").addEventListener("click", () => addInventory(true));
  $("installHelp").addEventListener("click", () => alert("On iPhone: Share -> Add to Home Screen."));
  if (state.pin) searchInventory();
}

bind();

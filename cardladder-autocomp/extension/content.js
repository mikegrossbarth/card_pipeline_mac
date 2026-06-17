const CARDLADDER_CONTENT_VERSION = "2026-06-17-wait-for-cardladder-login-v10";
let lastGraderOpenDebug = [];
const COMP_SOURCE_LABELS = [
  "eBay",
  "Goldin",
  "Goldin-Marketplace",
  "PWCC-Premier",
  "PWCC-Monthly",
  "PWCC-Vault",
  "Heritage",
  "MySlabs",
  "Pristine",
  "Pristine Auction",
  "Alt",
  "Lelands",
  "MemoryLane",
  "REA",
  "SCP",
  "MileHigh",
  "LoveOfTheGame",
  "90sAuctions",
  "Iconic",
  "Juliens",
  "Collectable-Buyout",
  "HugginsAndScott",
  "Beckett",
  "Sirius",
  "SacoRiver",
  "Rally",
  "Worthpoint",
  "CleanSweep",
  "ZeroCool",
  "Wheatland",
  "GregBussineau",
  "GoodwinAuctionCompany",
  "TheCollectorConnection",
  "RRAuction",
  "CollectAuctions",
  "Private",
  "Fanatics",
  "Card Ladder",
];
const COMP_SOURCE_PATTERN_TEXT = COMP_SOURCE_LABELS
  .map(sourceLabelToPattern)
  .sort((a, b) => b.length - a.length)
  .join("|");
const COMP_SOURCE_PATTERN = new RegExp(`\\b(${COMP_SOURCE_PATTERN_TEXT})\\b`, "i");

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === "CARDLADDER_CAPTURE_CURRENT") {
    sendResponse({
      ...(message.row || {}),
      value: null,
      status: "capture_requires_background_ocr",
      pageUrl: location.href,
      capturedAt: new Date().toISOString(),
    });
    return true;
  }
  if (message.type === "CARDLADDER_SELECT_GRADER") {
    selectGraderForAutomation(message.grader || "PSA")
      .then((result) => sendResponse(result))
      .catch((error) => sendResponse({ ok: false, error: error.message, version: CARDLADDER_CONTENT_VERSION }));
    return true;
  }
  if (message.type === "CARDLADDER_GRADER_GEOMETRY") {
    sendResponse(graderGeometry(message.grader || "PSA"));
    return true;
  }
  if (message.type === "CARDLADDER_GRADER_OPTION_GEOMETRY") {
    sendResponse(graderOptionGeometry(message.grader || "PSA"));
    return true;
  }
  if (message.type === "CARDLADDER_EXTRACT_DOM_RESULT") {
    sendResponse(extractDomResult(message.row || {}));
    return true;
  }
  if (message.type === "CARDLADDER_CHECK_INVALID_CERT_TOAST") {
    const reason = invalidCertToastReason()
      || invalidCertReasonFromText(document.body.innerText || "");
    sendResponse({
      ok: !reason,
      invalid: Boolean(reason),
      status: reason ? "invalid_cert" : "ok",
      error: reason,
      pageUrl: location.href,
      capturedAt: new Date().toISOString(),
    });
    return true;
  }
  if (message.type === "CARDLADDER_PREPARE_CERT_MODAL") {
    prepareCertModal()
      .then((result) => sendResponse(result))
      .catch((error) => sendResponse({ ok: false, error: error.message, version: CARDLADDER_CONTENT_VERSION }));
    return true;
  }
  if (message.type === "CARDLADDER_SUBMIT_CERT_MODAL") {
    submitPreparedCertModal(message.row)
      .then((result) => sendResponse(result))
      .catch((error) => sendResponse({ ...(message.row || {}), value: null, status: "error", error: error.message, capturedAt: new Date().toISOString() }));
    return true;
  }
  if (message.type !== "CARDLADDER_LOOKUP_ROW") return false;
  runLookup(message.row)
    .then((result) => sendResponse(result))
    .catch((error) => sendResponse({ results: [], error: error.message }));
  return true;
});

async function runLookup(row) {
  try {
    await clickLoginIfNeeded();
    await ensureSalesHistory();
    await clickCertMode();
    await chooseGrader(row.grader);
    await fillCert(row.certNumber);
    await submitSearch();
    await waitForResultsPage();
    return {
      ...row,
      value: null,
      status: "submitted",
      pageUrl: location.href,
      capturedAt: new Date().toISOString(),
    };
  } catch (error) {
    return {
      ...row,
      value: null,
      status: "error",
      error: error.message,
      capturedAt: new Date().toISOString(),
    };
  }
}

async function prepareCertModal() {
  await clickLoginIfNeeded();
  await ensureSalesHistory();
  await clickCertMode();
  if (!certSearchModal()) throw new Error("Could not open cert search modal.");
  return { ok: true, version: CARDLADDER_CONTENT_VERSION };
}

async function submitPreparedCertModal(row) {
  clearInvalidCertAlerts();
  const beforeUrl = location.href;
  const beforeSignature = pageResultSignature();
  await fillCert(row.certNumber);
  await submitSearch();
  const resultState = await waitForResultsPage(row, beforeUrl, beforeSignature);
  if (["invalid_cert", "no_results"].includes(resultState.status)) {
    const noResultsDetails = resultState.status === "no_results"
      ? await waitForNoResultsProfileDetails()
      : { profile: { title: "", grader: "", grade: "" }, resultCount: null, evidence: "" };
    return {
      ...row,
      value: null,
      status: resultState.status,
      error: resultState.reason,
      ocr: {
        ok: false,
        value: null,
        labelSeen: false,
        profileTitle: noResultsDetails.profile.title,
        profileGrader: noResultsDetails.profile.grader,
        profileGrade: noResultsDetails.profile.grade,
        resultCount: noResultsDetails.resultCount,
        comps: [],
        evidence: noResultsDetails.evidence || resultState.reason,
        debugImage: "",
      },
      pageUrl: location.href,
      capturedAt: new Date().toISOString(),
    };
  }
  if (["stale_result", "unknown"].includes(resultState.status)) {
    return {
      ...row,
      value: null,
      status: "error",
      error: resultState.reason || "Card Ladder did not load a new matching result after submit.",
      pageUrl: location.href,
      capturedAt: new Date().toISOString(),
    };
  }
  return {
    ...row,
    value: null,
    status: "submitted",
    pageUrl: location.href,
    capturedAt: new Date().toISOString(),
  };
}

async function waitForNoResultsProfileDetails() {
  let lastText = "";
  for (let attempt = 0; attempt < 10; attempt += 1) {
    lastText = document.body.innerText || "";
    const profile = extractProfileFromText(lastText);
    const resultCount = extractResultCount(lastText);
    if (profile.title) {
      return {
        profile,
        resultCount,
        evidence: `Extracted Card Ladder profile after no-results settle attempt ${attempt + 1}.`,
      };
    }
    await sleep(300);
  }
  return {
    profile: extractProfileFromText(lastText),
    resultCount: extractResultCount(lastText),
    evidence: "Card Ladder showed no matching results before a profile title appeared.",
  };
}

async function waitForResultsPage(row = {}, beforeUrl = "", beforeSignature = "") {
  const startedAt = Date.now();
  for (let i = 0; i < 45; i += 1) {
    const text = document.body.innerText || "";
    if (pageUrlMatchesCert(row.certNumber, beforeUrl)) {
      await sleep(300);
      const lateInvalidCertReason = invalidCertToastReason() || invalidCertReasonFromText(document.body.innerText || "");
      if (lateInvalidCertReason) return { status: "invalid_cert", reason: lateInvalidCertReason };
      return { status: "results" };
    }
    const invalidCertReason = invalidCertToastReason() || invalidCertReasonFromText(text);
    if (invalidCertReason) {
      return { status: "invalid_cert", reason: invalidCertReason };
    }
    const noResultsReason = noResultsReasonFromText(text);
    if (noResultsReason) {
      return { status: "no_results", reason: noResultsReason };
    }
    if (Date.now() - startedAt >= 1800 && !certSearchModalVisible() && (/Grade:\s*.+Grader:\s*.+Profile:/i.test(text) || /CL\s*Value/i.test(text)) && /\$\s*\d/i.test(text)) {
      await sleep(300);
      const lateInvalidCertReason = invalidCertToastReason() || invalidCertReasonFromText(document.body.innerText || "");
      if (lateInvalidCertReason) return { status: "invalid_cert", reason: lateInvalidCertReason };
      if (!resultPageChanged(beforeUrl, beforeSignature) && !profileMatchesRequestedRow(row, document.body.innerText || "")) {
        return { status: "stale_result", reason: "Card Ladder stayed on the previous result page after submit." };
      }
      return { status: "results" };
    }
    await sleep(300);
  }
  await sleep(500);
  return { status: "unknown" };
}

function resultPageChanged(beforeUrl = "", beforeSignature = "") {
  if (beforeUrl && location.href !== beforeUrl) return true;
  const afterSignature = pageResultSignature();
  return Boolean(afterSignature && beforeSignature && afterSignature !== beforeSignature);
}

function pageResultSignature() {
  const text = document.body.innerText || "";
  const profile = extractProfileFromText(text);
  const value = readCardLadderValue();
  const resultCount = extractResultCount(text);
  return [profile.title, profile.grader, profile.grade, value ?? "", resultCount ?? ""]
    .join("|")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function profileMatchesRequestedRow(row = {}, text = "") {
  const requested = String(row.cardTitle || "").trim();
  if (!requested) return false;
  const profile = extractProfileFromText(text).title;
  if (!profile) return false;
  const requestedTokens = meaningfulTitleTokens(requested);
  const profileTokens = meaningfulTitleTokens(profile);
  if (requestedTokens.length < 2 || profileTokens.length < 2) return false;
  const profileSet = new Set(profileTokens);
  const matches = requestedTokens.filter((token) => profileSet.has(token)).length;
  return matches >= Math.min(4, Math.ceil(requestedTokens.length * 0.45));
}

function meaningfulTitleTokens(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\b(?:psa|bgs|sgc|cgc|gem|mint|mt|grade|grader|pop|rookie|rc|prizm|refractor)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .split(/\s+/)
    .filter((token) => token.length >= 2)
    .slice(0, 20);
}

function pageUrlMatchesCert(certNumber, beforeUrl = "") {
  const cert = String(certNumber || "").replace(/\D/g, "");
  if (!cert) return false;
  const url = location.href || "";
  if (beforeUrl && url === beforeUrl) return false;
  return new RegExp(`(?:psa|bgs|sgc|cgc|beckett)[^0-9]{0,12}${escapeRegExp(cert)}|${escapeRegExp(cert)}`, "i").test(decodeURIComponent(url));
}

function invalidCertReasonFromText(text) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  if (!normalized) return "";
  const patterns = [
    /\bno\s+information\s+with\s+this\s+cert\b/i,
    /\bno\s+information\s+for\s+this\s+cert\b/i,
    /\binvalid\s+cert(?:ification)?\s*(?:number|#)?\b/i,
    /\bcert(?:ification)?\s*(?:number|#)?\s+not\s+found\b/i,
  ];
  const matched = patterns.find((pattern) => pattern.test(normalized));
  return matched ? "Card Ladder showed no information with this cert." : "";
}

function clearInvalidCertAlerts() {
  const candidates = invalidCertToastCandidates();
  for (const item of candidates.slice(0, 4)) {
    item.el.remove();
  }
}

function invalidCertToastReason() {
  const toast = invalidCertToastCandidates()[0];
  return toast ? "Card Ladder showed no information with this cert." : "";
}

function invalidCertToastCandidates() {
  const viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
  return [...document.querySelectorAll("body *")]
    .filter((el) => isVisible(el))
    .map((el) => ({ el, text: visibleText(el).replace(/\s+/g, " ").trim(), rect: el.getBoundingClientRect() }))
    .filter((item) => invalidCertReasonFromText(item.text))
    .filter((item) => item.text.length <= 260 && item.rect.width <= 620 && item.rect.height <= 220)
    .filter((item) => {
      const nearBottomRight = item.rect.right >= viewportWidth - 460 && item.rect.bottom >= viewportHeight - 280;
      const alertLike = /alert|toast|snackbar|notification/i.test(`${item.el.className || ""} ${item.el.getAttribute("role") || ""} ${item.el.getAttribute("aria-live") || ""}`);
      return nearBottomRight || alertLike;
    })
    .sort((a, b) => {
      const aScore = (viewportWidth - a.rect.right) + (viewportHeight - a.rect.bottom) + a.text.length;
      const bScore = (viewportWidth - b.rect.right) + (viewportHeight - b.rect.bottom) + b.text.length;
      return aScore - bScore;
    });
}

function certSearchModalVisible() {
  if (certSearchModal()) return true;
  const title = [...document.querySelectorAll("body *")]
    .filter((el) => isVisible(el))
    .map((el) => ({ el, text: visibleText(el).replace(/\s+/g, " ").trim(), rect: el.getBoundingClientRect() }))
    .filter((item) => /^SEARCH SALES BY CERT #$/i.test(item.text) || /SEARCH SALES BY CERT #/i.test(item.text))
    .sort((a, b) => a.text.length - b.text.length || a.rect.top - b.rect.top)[0];
  if (!title) return false;
  const centerBand = title.rect.left >= window.innerWidth * 0.18 && title.rect.right <= window.innerWidth * 0.85 && title.rect.top < window.innerHeight * 0.35;
  if (!centerBand) return false;
  const pageText = String(document.body.innerText || "").replace(/\s+/g, " ");
  return /Cert #/i.test(pageText) && /Grader/i.test(pageText) && /Submit/i.test(pageText);
}

function noResultsReasonFromText(text) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  if (!normalized) return "";
  const patterns = [
    /\b0\s+results?\b/i,
    /\bthere\s+are\s+no\s+results\s+for\s+your\s+query\b/i,
    /\btry\s+searching\s+for\s+something\s+else\b/i,
    /\bno\s+(?:sales\s+)?results?\s+found\b/i,
    /\bno\s+matching\s+(?:sales\s+)?results?\b/i,
    /\bno\s+sales\s+history\b/i,
    /\bwe\s+could(?:n['’]?t| not)\s+find\b/i,
    /\bno\s+matches?\b/i,
  ];
  const matched = patterns.find((pattern) => pattern.test(normalized));
  return matched ? "Card Ladder showed no matching results." : "";
}

async function clickLoginIfNeeded() {
  const login = findClickable(/^(log in|login|sign in)$/i);
  if (login) {
    login.click();
    await sleep(2500);
  }
}

async function ensureSalesHistory() {
  if (!location.pathname.includes("sales-history")) {
    location.href = "https://app.cardladder.com/sales-history";
    await sleep(3000);
  }
}

async function clickCertMode() {
  if (certSearchModal()) return;

  await resetSearchFocus();

  const searchInput = findSearchInput();
  if (searchInput) {
    const hashNode = findHashControlNearSearch(searchInput);
    if (hashNode) {
      clickLikeHuman(hashNode);
      await sleep(650);
      if (certSearchModal()) return;
      if (await clickCertMenuOptionIfShown()) return;
      if (certInputIsVisible()) return;
    }
  }

  const exactHash = [...document.querySelectorAll("button, [role='button'], a")]
    .find((el) => visibleText(el) === "#" || (el.getAttribute("aria-label") || "").match(/cert|number|#|hash/i));
  if (exactHash) {
    clickLikeHuman(exactHash);
    await sleep(500);
    await clickCertMenuOptionIfShown();
    return;
  }

  const nearSearch = [...document.querySelectorAll("button, [role='button']")]
    .find((el) => visibleText(el).includes("#"));
  if (nearSearch) {
    clickLikeHuman(nearSearch);
    await sleep(500);
    await clickCertMenuOptionIfShown();
    return;
  }

  if (searchInput) {
    const rect = searchInput.getBoundingClientRect();
    const clickPoints = [
      [rect.right - 22, rect.top + rect.height / 2],
      [rect.right - 34, rect.top + rect.height / 2],
      [rect.right - 12, rect.top + rect.height / 2],
    ];
    for (const [x, y] of clickPoints) {
      clickAtPoint(x, y);
      await sleep(500);
      if (document.activeElement === searchInput) {
        searchInput.blur();
        document.body.click();
        await sleep(300);
      }
      if (certSearchModal()) return;
      if (await clickCertMenuOptionIfShown()) return;
      if (certInputIsVisible()) return;
    }
  }

  const option = findClickable(/^(#|cert\s*#|certification\s*#|certification number|cert number)$/i);
  if (option) {
    clickLikeHuman(option);
    await sleep(500);
    return;
  }

  throw new Error("Could not find # cert search mode. Page clue: " + pageClue());
}

async function resetSearchFocus() {
  document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
  document.activeElement?.blur?.();
  await closeGlobalSearchIfOpen();

  const heading = [...document.querySelectorAll("h1, h2, [role='heading'], div, main")]
    .filter((el) => isVisible(el))
    .find((el) => /^SALES$/i.test(visibleText(el)));
  if (heading) {
    clickLikeHuman(heading);
    await sleep(300);
    return;
  }

  const searchInput = findSearchInput();
  if (searchInput) {
    const rect = searchInput.getBoundingClientRect();
    clickAtPoint(Math.max(20, rect.left - 30), Math.max(20, rect.top - 35));
    await sleep(300);
    return;
  }

  clickAtPoint(Math.min(window.innerWidth - 20, 260), Math.min(window.innerHeight - 20, 180));
  await sleep(300);
}

async function chooseGrader(grader) {
  const normalized = String(grader || "").toUpperCase();
  const optionLabel = cardLadderGraderLabel(normalized);
  const modal = certSearchModal();
  if (modal) {
    const graderControl = findGraderControlInModal(modal) || findFieldControlInModal(modal, /grader/i);
    if (graderControl) {
      if (selectedGraderMatches(graderControl, normalized)) return;
      if (await clickDropdownThenOption(modal, graderControl, normalized, optionLabel)) {
        if (await waitForSelectedGrader(modal, normalized)) return;
      }
    }
  }

  const select = [...document.querySelectorAll("select")].find((el) =>
    [...el.options].some((option) => graderLabelMatches(option.textContent, normalized))
  );
  if (select) {
    select.value = [...select.options].find((option) => graderLabelMatches(option.textContent, normalized)).value;
    select.dispatchEvent(new Event("change", { bubbles: true }));
    await sleep(300);
    if (selectedGraderMatches(select, normalized)) return;
  }

  const combobox = document.querySelector("[role='combobox'], input[placeholder*='Company' i], input[placeholder*='Grader' i]");
  if (combobox) {
    combobox.click();
    await sleep(300);
    const option = findClickable(new RegExp(`^${escapeRegExp(optionLabel)}$`, "i"));
    if (option) {
      option.click();
      await sleep(300);
      if (await waitForSelectedGrader(modal || certSearchModal(), normalized)) return;
    }
  }

  const textOption = findClickable(new RegExp(`^${escapeRegExp(optionLabel)}$`, "i"));
  if (textOption) {
    textOption.click();
    await sleep(300);
    if (await waitForSelectedGrader(modal || certSearchModal(), normalized)) return;
  }

  throw new Error(`Could not select grader ${normalized}. ${graderSelectionDebug(modal, optionLabel)}`);
}

async function selectGraderForAutomation(grader) {
  const normalized = String(grader || "").toUpperCase();
  const optionLabel = cardLadderGraderLabel(normalized);
  const modal = certSearchModal();
  if (!modal) {
    return { ok: false, version: CARDLADDER_CONTENT_VERSION, error: "Open the SEARCH SALES BY CERT # modal first." };
  }
  const control = findGraderControlInModal(modal) || findFieldControlInModal(modal, /grader/i);
  if (control && selectedGraderMatches(control, normalized)) {
    return { ok: true, version: CARDLADDER_CONTENT_VERSION, grader: normalized, selectedLabel: optionLabel, skipped: "already selected" };
  }
  await chooseGrader(normalized);
  const afterControl = findGraderControlInModal(modal) || findFieldControlInModal(modal, /grader/i);
  if (!selectedGraderMatches(afterControl, normalized)) {
    return {
      ok: false,
      version: CARDLADDER_CONTENT_VERSION,
      grader: normalized,
      selectedLabel: optionLabel,
      selectedText: selectedControlText(afterControl),
      error: `Card Ladder grader stayed on ${selectedControlText(afterControl) || "unknown"} instead of ${optionLabel}. ${graderSelectionDebug(modal, optionLabel)}`,
    };
  }
  return {
    ok: true,
    version: CARDLADDER_CONTENT_VERSION,
    grader: normalized,
    selectedLabel: optionLabel,
    selectedText: selectedControlText(afterControl),
  };
}

function graderGeometry(grader) {
  const normalized = String(grader || "").toUpperCase();
  const modal = certSearchModal();
  if (!modal) return { ok: false, version: CARDLADDER_CONTENT_VERSION, error: "Open the SEARCH SALES BY CERT # modal first." };
  const control = findGraderControlInModal(modal) || findFieldControlInModal(modal, /grader/i);
  if (!control) return { ok: false, version: CARDLADDER_CONTENT_VERSION, error: "Could not find grader control." };
  const rect = graderFieldRect(modal, control);
  return {
    ok: true,
    version: CARDLADDER_CONTENT_VERSION,
    grader: normalized,
    selectedText: selectedControlText(control),
    control: describeNode(control, rect),
    rect: plainRect(rect),
    clickPoint: {
      x: Math.round(Math.max(rect.left + 20, rect.right - 20)),
      y: Math.round(rect.top + rect.height / 2),
    },
    optionPoint: knownGraderOptionPoint(rect, normalized),
  };
}

function graderOptionGeometry(grader) {
  const normalized = String(grader || "").toUpperCase();
  const optionLabel = cardLadderGraderLabel(normalized);
  const modal = certSearchModal();
  const control = modal ? (findGraderControlInModal(modal) || findFieldControlInModal(modal, /grader/i)) : null;
  const rect = control ? graderFieldRect(modal, control) : null;
  const option = findGraderOption(optionLabel, rect) || findGraderOptionByPosition(normalized, rect);
  if (option) {
    const optionRect = option.getBoundingClientRect();
    return {
      ok: true,
      version: CARDLADDER_CONTENT_VERSION,
      grader: normalized,
      option: describeNode(option, optionRect),
      rect: plainRect(optionRect),
      clickPoint: {
        x: Math.round(optionRect.left + Math.min(38, Math.max(12, optionRect.width / 2))),
        y: Math.round(optionRect.top + optionRect.height / 2),
      },
    };
  }
  if (rect) {
    return {
      ok: true,
      version: CARDLADDER_CONTENT_VERSION,
      grader: normalized,
      fallback: "known-position",
      clickPoint: knownGraderOptionPoint(rect, normalized),
    };
  }
  return { ok: false, version: CARDLADDER_CONTENT_VERSION, error: "Could not find grader option geometry." };
}

function knownGraderOptionPoint(rect, grader) {
  const indexByGrader = {
    PSA: 0,
    BGS: 1,
    BECKETT: 1,
    SGC: 2,
    CGC: 3,
  };
  const targetIndex = indexByGrader[grader];
  if (targetIndex == null) return null;
  const optionHeight = Math.max(48, Math.min(58, rect.height));
  return {
    x: Math.round(Math.min(rect.right - 24, rect.left + 38)),
    y: Math.round(rect.bottom + optionHeight * targetIndex + optionHeight / 2),
  };
}

function plainRect(rect) {
  return {
    left: Math.round(rect.left),
    top: Math.round(rect.top),
    right: Math.round(rect.right),
    bottom: Math.round(rect.bottom),
    width: Math.round(rect.width),
    height: Math.round(rect.height),
  };
}

async function clickDropdownThenOption(modal, control, grader, optionLabel) {
  lastGraderOpenDebug = [];
  await clickGraderDropdown(modal, control);
  const rect = graderFieldRect(modal, control);
  const option = findGraderOption(optionLabel, rect) || findGraderOptionByPosition(grader, rect);
  if (option) {
    clickLikeHuman(option);
    await sleep(650);
    return true;
  }
  return clickGraderOptionByKnownPosition(modal, control, grader);
}

function cardLadderGraderLabel(grader) {
  const labels = {
    BGS: "BECKETT",
    BVG: "BECKETT",
  };
  return labels[grader] || grader;
}

async function waitForSelectedGrader(modal, grader) {
  for (let attempt = 0; attempt < 6; attempt += 1) {
    await sleep(200);
    const currentModal = modal || certSearchModal();
    const control = currentModal ? (findGraderControlInModal(currentModal) || findFieldControlInModal(currentModal, /grader/i)) : null;
    if (selectedGraderMatches(control, grader)) return true;
  }
  return false;
}

function selectedGraderMatches(control, grader) {
  return graderLabelMatches(selectedControlText(control), grader);
}

function graderLabelMatches(text, grader) {
  const selected = normalizeGraderLabel(text);
  const expected = normalizeGraderLabel(grader);
  return Boolean(selected && expected && selected === expected);
}

function normalizeGraderLabel(value) {
  const text = String(value || "")
    .toUpperCase()
    .replace(/\bGRADER\b|[:*]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return "";
  const labels = [];
  if (/\bPSA\b/.test(text)) labels.push("PSA");
  if (/\b(?:BGS|BECKETT|BVG)\b/.test(text)) labels.push("BGS");
  if (/\bSGC\b/.test(text)) labels.push("SGC");
  if (/\bCGC\b/.test(text)) labels.push("CGC");
  const unique = [...new Set(labels)];
  return unique.length === 1 ? unique[0] : "";
}

function findGraderControlInModal(modal) {
  const direct = findDirectGraderControl(modal);
  if (direct) return direct;

  const labels = [...modal.querySelectorAll("label, legend, span, div")]
    .filter((el) => isVisible(el) && /^grader$/i.test(visibleText(el).replace(/[:*]/g, "").trim()));

  for (const label of labels) {
    const labelRect = label.getBoundingClientRect();
    const controls = [...modal.querySelectorAll("select, [role='combobox'], button, input, div")]
      .filter((el) => isVisible(el) && el !== label && !label.contains(el))
      .map((el) => ({ el, rect: el.getBoundingClientRect(), text: selectedControlText(el), rawText: visibleText(el).replace(/\s+/g, " ").trim() }))
      .filter(({ rect }) => {
        const overlapsLabelRow = rect.top <= labelRect.bottom + 35 && rect.bottom >= labelRect.top - 10;
        const belowLabel = rect.top >= labelRect.bottom - 10 && rect.top <= labelRect.bottom + 95;
        return (overlapsLabelRow || belowLabel) &&
          rect.left >= labelRect.left - 16 &&
          rect.left <= labelRect.left + 620;
      })
      .filter(({ rect, text }) =>
        rect.width >= 80 &&
        rect.height >= 20 &&
        !/^cert/i.test(text) &&
        text.length <= 120 &&
        (!text || normalizeGraderLabel(text) || /expand_more|arrow_drop_down|▾|▼/i.test(text))
      )
      .sort((a, b) =>
        graderControlScore(a) - graderControlScore(b) ||
        (a.rect.top - b.rect.top) ||
        (a.rect.width * a.rect.height) - (b.rect.width * b.rect.height)
      );
    if (controls[0]) return controls[0].el;
  }

  return [...modal.querySelectorAll("[role='combobox'], select, button")]
    .filter((el) => isVisible(el))
    .find((el) => {
      const text = selectedControlText(el);
      return text.length <= 120 && /PSA|BECKETT|BGS|SGC|CGC|CSG|TAG|ISA|HGA/i.test(text);
    }) || null;
}

function findDirectGraderControl(modal) {
  const selectors = [
    "[role='combobox']",
    "[aria-haspopup='listbox']",
    "[aria-expanded]",
    "button",
    "select",
    "input",
    "div",
  ];
  return [...modal.querySelectorAll(selectors.join(","))]
    .filter((el) => isVisible(el))
    .map((el) => ({ el, rect: el.getBoundingClientRect(), text: selectedControlText(el), rawText: visibleText(el).replace(/\s+/g, " ").trim() }))
    .filter(({ rect, text, rawText }) =>
      rect.width >= 50 &&
      rect.width <= 260 &&
      rect.height >= 18 &&
      rect.height <= 72 &&
      `${text} ${rawText}`.length <= 160 &&
      /(?:^|\b)(?:PSA|BECKETT|BGS|SGC|CGC)(?:\b|$)/i.test(`${text} ${rawText}`) &&
      !/cert|submit|search sales/i.test(`${text} ${rawText}`)
    )
    .sort((a, b) =>
      graderControlScore(a) - graderControlScore(b) ||
      (a.rect.top - b.rect.top) ||
      (a.rect.left - b.rect.left)
    )[0]?.el || null;
}

function graderControlScore(item) {
  const text = String(item.text || item.rawText || "");
  let score = 0;
  if (!normalizeGraderLabel(text)) score += 8;
  if (!/expand_more|arrow_drop_down|▾|▼/i.test(text)) score += 2;
  if (item.el.getAttribute("role") === "combobox") score -= 3;
  if (item.el.matches?.("button, select, input")) score -= 2;
  if (item.rect.width > 420) score += 4;
  if (item.rect.height > 80) score += 3;
  return score;
}

async function clickGraderDropdown(modal, control) {
  await sleep(150);
  if (typeof control.focus === "function") control.focus();
  const rect = graderFieldRect(modal, control);
  rememberGraderOpenDebug("control", control, rect);
  const clickTargets = [
    () => clickAtPoint(Math.max(rect.left + 20, rect.right - 20), rect.top + rect.height / 2, "right-edge"),
    () => clickAtPoint(Math.max(rect.left + 20, rect.right - 36), rect.top + rect.height / 2, "select-face-right"),
    () => clickAtPoint(rect.left + Math.min(90, rect.width / 2), rect.top + rect.height / 2, "select-face-mid"),
    () => clickGraderExpandIcon(modal, control),
    () => clickLikeHuman(control, Math.max(rect.left + 20, rect.right - 20), rect.top + rect.height / 2, "control-direct"),
    () => openDropdownWithKeyboard(control),
  ];
  for (const open of clickTargets) {
    open();
    await sleep(550);
    const opened = findAnyGraderOptions(rect);
    rememberGraderOpenDebug(opened ? "opened" : "still-closed", document.activeElement || control);
    if (opened) return;
  }
}

function graderFieldRect(modal, control) {
  const labels = [...modal.querySelectorAll("label, legend, span, div")]
    .filter((el) => isVisible(el) && /^grader$/i.test(visibleText(el).replace(/[:*]/g, "").trim()));
  const modalRect = modal.getBoundingClientRect();
  const controlRect = control.getBoundingClientRect();
  const label = labels
    .map((el) => ({ el, rect: el.getBoundingClientRect() }))
    .sort((a, b) => a.rect.top - b.rect.top)[0];

  if (!label) return controlRect;

  const overlapsLabelRow = controlRect.top <= label.rect.bottom + 18 && controlRect.bottom >= label.rect.top - 8;
  if (overlapsLabelRow && controlRect.width >= 60 && controlRect.height >= 18) {
    return controlRect;
  }

  const top = Math.max(label.rect.bottom - 4, controlRect.top);
  const height = Math.max(42, Math.min(56, controlRect.height || 48));
  return {
    left: Math.max(modalRect.left + 18, controlRect.left || modalRect.left + 20),
    right: Math.min(modalRect.right - 18, controlRect.right || modalRect.right - 20),
    top,
    bottom: top + height,
    width: Math.min(modalRect.right - 36, controlRect.right || modalRect.right - 20) - Math.max(modalRect.left + 18, controlRect.left || modalRect.left + 20),
    height,
  };
}

function clickGraderExpandIcon(modal, control) {
  const controlRect = control.getBoundingClientRect();
  const icons = [...modal.querySelectorAll("i, svg, [class*='material-icons'], [class*='arrow'], [class*='Select-icon'], [data-testid*='ArrowDropDown'], span, div, button")]
    .filter((el) => isVisible(el))
    .map((el) => ({
      el,
      text: `${visibleText(el)} ${el.getAttribute("aria-label") || ""} ${el.getAttribute("data-testid") || ""} ${el.getAttribute("class") || ""}`.replace(/\s+/g, " ").trim(),
      rect: el.getBoundingClientRect(),
    }))
    .filter(({ text, rect }) =>
      (/expand_more|arrow_drop_down|arrowdropdown|select-icon|▾|▼/i.test(text) || rect.width <= 40) &&
      rect.top >= controlRect.top - 12 &&
      rect.bottom <= controlRect.bottom + 18 &&
      rect.left >= controlRect.left &&
      rect.right <= controlRect.right + 24
    )
    .sort((a, b) =>
      arrowIconScore(a) - arrowIconScore(b) ||
      b.rect.left - a.rect.left
    );
  const icon = icons[0];
  if (icon) {
    rememberGraderOpenDebug("icon", icon.el, icon.rect);
    clickAtPoint(icon.rect.left + icon.rect.width / 2, icon.rect.top + icon.rect.height / 2, "icon-point");
    clickLikeHuman(icon.el, icon.rect.left + icon.rect.width / 2, icon.rect.top + icon.rect.height / 2, "icon-direct");
    return true;
  }
  rememberGraderOpenDebug("icon-missing", control);
  return false;
}

function arrowIconScore(item) {
  const text = `${item.text || ""} ${item.el.tagName || ""}`.toLowerCase();
  let score = 0;
  if (!/expand_more|arrow_drop_down|arrowdropdown|select-icon|material-icons|arrow|▾|▼/i.test(text)) score += 10;
  if (/^expand_more\b/i.test(item.text)) score -= 6;
  if (item.el.matches?.("i, svg, [class*='material-icons'], [class*='arrow']")) score -= 4;
  if (item.rect.width > 48 || item.rect.height > 48) score += 5;
  return score;
}

function openDropdownWithKeyboard(control) {
  if (typeof control.focus === "function") control.focus();
  rememberGraderOpenDebug("keyboard", control);
  const keys = [
    { key: "ArrowDown", code: "ArrowDown", altKey: true },
    { key: " ", code: "Space" },
    { key: "Enter", code: "Enter" },
  ];
  for (const init of keys) {
    control.dispatchEvent(new KeyboardEvent("keydown", { ...init, bubbles: true, cancelable: true }));
    invokeFrameworkHandlers(control, "keyDown", init);
    control.dispatchEvent(new KeyboardEvent("keyup", { ...init, bubbles: true, cancelable: true }));
    invokeFrameworkHandlers(control, "keyUp", init);
  }
}

function findAnyGraderOptions(controlRect = null) {
  return [...document.querySelectorAll("[role='option'], [role='menuitem'], li, button, div, span")]
    .filter((el) => isVisible(el))
    .some((el) => {
      const text = visibleText(el).replace(/\s+/g, " ").trim();
      return /^(PSA|BECKETT|BGS|SGC|CGC)$/i.test(text) &&
        isNearGraderDropdown(el.getBoundingClientRect(), controlRect);
    });
}

function findGraderOption(grader, controlRect = null) {
  const pattern = new RegExp(`(^|\\b)${escapeRegExp(grader)}($|\\b)`, "i");
  const candidates = [...document.querySelectorAll("[role='option'], [role='menuitem'], li, button, div, span")]
    .filter((el) => isVisible(el))
    .map((el) => ({ el, text: visibleText(el).replace(/\s+/g, " ").trim(), rect: el.getBoundingClientRect() }))
    .filter(({ text, rect }) => text && text.length <= 40 && pattern.test(text) && rect.top > 80)
    .filter(({ rect }) => isNearGraderDropdown(rect, controlRect))
    .filter(({ text }) => !/grader|cert|submit|search/i.test(text));

  candidates.sort((a, b) => {
    const exactA = a.text.toUpperCase() === grader ? 0 : 1;
    const exactB = b.text.toUpperCase() === grader ? 0 : 1;
    return exactA - exactB || a.text.length - b.text.length || a.rect.top - b.rect.top;
  });

  return candidates[0]?.el || null;
}

function findGraderOptionByPosition(grader, controlRect = null) {
  const indexByGrader = {
    PSA: 0,
    BGS: 1,
    BECKETT: 1,
    SGC: 2,
    CGC: 3,
  };
  const targetIndex = indexByGrader[grader];
  if (targetIndex == null) return null;
  const options = [...document.querySelectorAll("[role='option'], [role='menuitem'], li, button, div")]
    .filter((el) => isVisible(el))
    .map((el) => ({ el, text: visibleText(el).replace(/\s+/g, " ").trim(), rect: el.getBoundingClientRect() }))
    .filter(({ text, rect }) => text && text.length <= 40 && /^(PSA|BECKETT|BGS|SGC|CGC)$/i.test(text) && rect.top > 80)
    .filter(({ rect }) => isNearGraderDropdown(rect, controlRect))
    .sort((a, b) => a.rect.top - b.rect.top);
  return options[targetIndex]?.el || null;
}

function isNearGraderDropdown(rect, controlRect) {
  if (!controlRect) return true;
  const verticallyNear = rect.top >= controlRect.bottom - 12 && rect.top <= controlRect.bottom + 360;
  const horizontallyNear = rect.right >= controlRect.left - 24 && rect.left <= controlRect.right + 80;
  return verticallyNear && horizontallyNear;
}

async function clickGraderOptionByKnownPosition(modal, control, grader) {
  const indexByGrader = {
    PSA: 0,
    BGS: 1,
    BECKETT: 1,
    SGC: 2,
    CGC: 3,
  };
  const targetIndex = indexByGrader[grader];
  if (targetIndex == null) return false;

  const rect = graderFieldRect(modal, control);
  const optionHeight = Math.max(48, Math.min(58, rect.height));
  const x = Math.min(rect.right - 24, rect.left + 38);
  const y = rect.bottom + optionHeight * targetIndex + optionHeight / 2;
  if (y >= window.innerHeight - 8) return false;
  clickAtPoint(x, y);
  await sleep(900);
  return true;
}

function selectedControlText(control) {
  if (!control) return "";
  if (control.matches?.("select")) return control.options[control.selectedIndex]?.textContent?.trim() || "";
  if (control.matches?.("input, textarea")) return control.value || control.getAttribute("placeholder") || "";
  return visibleText(control).replace(/\s+/g, " ").trim();
}

function graderSelectionDebug(modal, optionLabel) {
  const visibleOptions = [...document.querySelectorAll("[role='option'], [role='menuitem'], li, button, div, span")]
    .filter((el) => isVisible(el))
    .map((el) => visibleText(el).replace(/\s+/g, " ").trim())
    .filter((text) => text && /PSA|BECKETT|BGS|SGC|CGC|Grader/i.test(text))
    .slice(0, 12)
    .join(" | ");
  const modalText = modal ? visibleText(modal).replace(/\s+/g, " ").slice(0, 220) : "no modal";
  const openDebug = lastGraderOpenDebug.length ? `; open ${lastGraderOpenDebug.join(" || ")}` : "";
  return `[${CARDLADDER_CONTENT_VERSION}; wanted ${optionLabel}; options ${visibleOptions || "none"}${openDebug}; modal ${modalText}]`;
}

async function fillCert(certNumber) {
  const cert = String(certNumber || "").trim();
  const modal = certSearchModal();
  if (modal) {
    const certInput = findFieldControlInModal(modal, /cert/i, "input");
    if (!certInput) throw new Error("Could not find cert input in cert search modal.");
    await setCertInputValue(certInput, cert);
    return;
  }

  const inputs = [...document.querySelectorAll("input:not([type='hidden']), textarea")];
  const certInput = inputs.find((el) =>
    `${el.placeholder || ""} ${el.getAttribute("aria-label") || ""} ${el.name || ""}`.match(/cert/i)
  ) || inputs[inputs.length - 1];

  if (!certInput) throw new Error("Could not find cert input.");
  await setCertInputValue(certInput, cert);
}

async function setCertInputValue(certInput, certNumber) {
  certInput.focus();
  certInput.click?.();
  certInput.dispatchEvent(new KeyboardEvent("keydown", { key: "a", code: "KeyA", ctrlKey: true, bubbles: true }));
  certInput.dispatchEvent(new KeyboardEvent("keyup", { key: "a", code: "KeyA", ctrlKey: true, bubbles: true }));
  setNativeValue(certInput, "");
  certInput.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "deleteContentBackward", data: null }));
  certInput.dispatchEvent(new Event("change", { bubbles: true }));
  await sleep(150);

  setNativeValue(certInput, certNumber);
  certInput.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: certNumber }));
  certInput.dispatchEvent(new Event("change", { bubbles: true }));
  await sleep(550);

  const currentValue = String(certInput.value || "").trim();
  if (currentValue !== certNumber) {
    setNativeValue(certInput, certNumber);
    certInput.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: certNumber }));
    certInput.dispatchEvent(new Event("change", { bubbles: true }));
    await sleep(550);
  }

  const verifiedValue = String(certInput.value || "").trim();
  if (verifiedValue !== certNumber) {
    throw new Error(`Cert input did not accept ${certNumber}; currently ${verifiedValue || "blank"}.`);
  }
}

async function submitSearch() {
  const modal = certSearchModal();
  if (modal) {
    const submit = [...modal.querySelectorAll("button, [role='button']")]
      .find((el) => /^submit$/i.test(visibleText(el)));
    if (!submit) throw new Error("Could not find Submit button in cert search modal.");
    await sleep(300);
    clickLikeHuman(submit);
    await sleep(900);
    return;
  }

  const button = findClickable(/^(search|apply|submit)$/i) || document.querySelector("button[type='submit']");
  if (button) {
    button.click();
  } else {
    document.activeElement?.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
  }
  await sleep(900);
}

function readCardLadderValue() {
  const text = document.body.innerText;
  const normalized = text.replace(/\s+/g, " ");
  const labeled = normalized.match(/Card\s*Ladder\s*Value[\s\S]{0,80}?\$\s*([\d,]+(?:\.\d{1,2})?)/i)
    || normalized.match(/\bC\s*L\s*Value[\s\S]{0,80}?\$\s*([\d,]+(?:\.\d{1,2})?)/i);
  if (labeled) return Number(labeled[1].replace(/,/g, ""));

  const valueLabel = normalized.search(/\b(?:C\s*L|Card\s*Ladder)\s*Value\b/i);
  if (valueLabel >= 0) {
    const afterLabel = normalized.slice(valueLabel, valueLabel + 300);
    const nearbyMoney = afterLabel.match(/\$\s*([\d,]+(?:\.\d{1,2})?)/);
    if (nearbyMoney) return Number(nearbyMoney[1].replace(/,/g, ""));
  }

  const profileSummary = normalized.match(/\b\d+\s+results\s+Grade:\s*[^$]{0,260}?\$\s*([\d,]+(?:\.\d{1,2})?)/i);
  if (profileSummary) return Number(profileSummary[1].replace(/,/g, ""));

  const beforeFirstSale = normalized.split(/\bEBAY\s+-\s+/i)[0] || "";
  if (/Grade:\s*/i.test(beforeFirstSale) && /Profile:/i.test(beforeFirstSale)) {
    const summaryMoney = beforeFirstSale.match(/\$\s*([\d,]+(?:\.\d{1,2})?)/);
    if (summaryMoney) return Number(summaryMoney[1].replace(/,/g, ""));
  }

  const clNode = [...document.querySelectorAll("body *")]
    .filter((el) => isVisible(el))
    .map((el) => ({ el, text: visibleText(el), rect: el.getBoundingClientRect() }))
    .filter((item) => /\bC\s*L\s*Value\b|\bCard\s*Ladder\s*Value\b/i.test(item.text))
    .sort((a, b) => a.rect.top - b.rect.top || a.rect.left - b.rect.left)[0];

  if (clNode) {
    const localText = collectNearbyText(clNode.el);
    const localMatch = localText.replace(/\s+/g, " ").match(/\$?\s*([\d,]+(?:\.\d{1,2})?)/);
    if (localMatch) return Number(localMatch[1].replace(/,/g, ""));
  }

  const moneyValues = [...text.matchAll(/\$\s*([\d,]+(?:\.\d{1,2})?)/g)]
    .map((match) => Number(match[1].replace(/,/g, "")))
    .filter((value) => Number.isFinite(value) && value > 0);

  return moneyValues.length === 1 ? moneyValues[0] : null;
}

function extractDomResult(row = {}) {
  const text = document.body.innerText || "";
  const invalidCertReason = invalidCertToastReason()
    || invalidCertReasonFromText(text);
  if (invalidCertReason) {
    return {
      ...row,
      ok: false,
      value: null,
      status: "invalid_cert",
      error: invalidCertReason,
      ocr: { ok: false, value: null, comps: [], evidence: invalidCertReason, debugImage: "" },
      pageUrl: location.href,
      capturedAt: new Date().toISOString(),
    };
  }
  const value = readCardLadderValue();
  const profile = extractProfileFromText(text);
  const comps = extractCompsFromText(text);
  const resultCount = extractResultCount(text);
  return {
    ...row,
    ok: value != null && comps.length > 0,
    value,
    status: value != null && comps.length > 0 ? "ok" : "dom_incomplete",
    ocr: {
      ok: value != null,
      value,
      labelSeen: value != null,
      profileTitle: profile.title,
      profileGrader: profile.grader,
      profileGrade: profile.grade,
      resultCount,
      comps,
      evidence: "Extracted from Card Ladder page text.",
      debugImage: "",
    },
    pageUrl: location.href,
    capturedAt: new Date().toISOString(),
  };
}

function extractResultCount(text) {
  const normalized = String(text || "").replace(/\s+/g, " ");
  const match = normalized.match(/\b(\d{1,4})\s+results?\b/i);
  if (!match) return null;
  const count = Number(match[1]);
  return Number.isFinite(count) ? count : null;
}

function extractProfileFromText(text) {
  const normalized = String(text || "").replace(/\s+/g, " ");
  const titleStop = `(?=\\s+(?:CL\\s*Value|Card\\s*Ladder\\s*Value|Grade:|Grader:|${COMP_SOURCE_PATTERN_TEXT}|close\\s+\\$|[x×]|help[_\\s-]*outline|Date\\s+Sold|No\\s+sales|No\\s+results|There\\s+are\\s+no\\s+results|Try\\s+searching|$))`;
  const gradeGraderProfile = normalized.match(new RegExp(`Grade:\\s*([^,|]+).*?Grader:\\s*([A-Z]+).*?Profile:\\s*(.*?)${titleStop}`, "i"));
  if (gradeGraderProfile) {
    return {
      grade: String(gradeGraderProfile[1] || "").trim(),
      grader: String(gradeGraderProfile[2] || "").trim().toUpperCase(),
      title: cleanProfileTitle(String(gradeGraderProfile[3] || "")),
    };
  }
  const profileFirst = normalized.match(new RegExp(`Profile:\\s*(.*?)${titleStop}.*?Grade:\\s*([^,|]+).*?Grader:\\s*([A-Z]+)`, "i"));
  if (profileFirst) {
    return {
      grade: String(profileFirst[2] || "").trim(),
      grader: String(profileFirst[3] || "").trim().toUpperCase(),
      title: cleanProfileTitle(String(profileFirst[1] || "")),
    };
  }
  const profileOnly = normalized.match(new RegExp(`Profile:\\s*(.*?)${titleStop}`, "i"));
  if (!profileOnly) return { title: "", grader: "", grade: "" };
  return {
    grade: "",
    grader: "",
    title: cleanProfileTitle(String(profileOnly[1] || "")),
  };
}

function cleanProfileTitle(value) {
  let title = String(value || "").replace(/\s+/g, " ").trim();
  const tailPatterns = [
    /\s+\bclose\s+\$?\d[\d,]*(?:\.\d{1,2})?.*$/i,
    /\s+[x×]\s*$/i,
    /\s+\bthere\s+are\s+no\s+results\b.*$/i,
    /\s+\btry\s+searching\b.*$/i,
    /\s+\bhelp[_\s-]*outline\b.*$/i,
    /\s+\b(?:date\s+sold|type|price)\b.*$/i,
    /\s+\$\d[\d,]*(?:\.\d{1,2})?\s+\b(?:help[_\s-]*outline|ebay|fanatics|pwcc|goldin|alt|myslabs|heritage|pristine|auction)\b.*$/i,
  ];
  for (const pattern of tailPatterns) {
    title = title.replace(pattern, "");
  }
  return title.replace(/\s*\(pop\s*[^)]*\)\s*$/i, "").replace(/\s+/g, " ").trim();
}

function extractCompsFromText(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  const comps = [];
  for (let i = 0; i < lines.length && comps.length < 20; i += 1) {
    const sourceMatch = sourceLineMatch(lines[i]);
    if (!sourceMatch) continue;
    const chunk = lines.slice(i, i + 8).join(" ");
    const comp = parseCompChunk(chunk, sourceMatch);
    if (!comp) continue;
    comps.push(comp);
  }
  return dedupeComps(comps).slice(0, 5);
}

function sourceLineMatch(line) {
  const text = String(line || "").trim();
  if (!text || /\b(?:CL|Card\s*Ladder)\s*Value\b/i.test(text)) return null;
  const match = text.match(new RegExp(`^(${COMP_SOURCE_PATTERN_TEXT})(?:\\s+\\([^)]{1,80}\\)|\\s*(?:-|–|—)\\s*.{1,100}|\\s+[A-Z0-9_'&. ]{2,100})?$`, "i"));
  return match;
}

function parseCompChunk(chunk, sourceLineMatchResult = null) {
  chunk = String(chunk || "").replace(/\s+/g, " ").trim();
  const sourceMatch = sourceLineMatchResult || chunk.match(COMP_SOURCE_PATTERN);
  if (!sourceMatch) return null;
  const sourceText = sourceMatch[1] || sourceMatch[0];
  const sourceIndex = chunk.toLowerCase().indexOf(String(sourceText).toLowerCase());
  if (sourceIndex > 0) chunk = chunk.slice(sourceIndex);
  chunk = chunk.replace(/^.*?\b(?:CL|Card\s*Ladder)\s*Value\b.*?(?=\b(?:Date\s+Sold|Type|Price)\b|\$|$)/i, " ");
  const dateMatch = chunk.match(compDatePattern());
  const priceMatches = [...chunk.matchAll(/\$\s*[\d,]+(?:\.\d{1,2})?/g)];
  if (!dateMatch || !priceMatches.length) return null;
  const price = priceMatches[priceMatches.length - 1][0].replace(/\s+/g, "");
  const saleType = (chunk.match(/\b(Auction|Best Offer|Buy It Now|Fixed Price|BIN)\b/i) || [""])[0];
  let title = chunk
    .replace(sourceText, " ")
    .replace(dateMatch[0], " ")
    .replace(price, " ")
    .replace(/\b(Auction|Best Offer|Buy It Now|Fixed Price|BIN)\b/ig, " ")
    .replace(/\b(?:CL|Card\s*Ladder)\s*Value\b.*$/i, " ")
    .replace(/\b(?:Date Sold|Type|Price)\b/ig, " ")
    .replace(/\s+/g, " ")
    .trim();
  title = title.split(new RegExp(`\\s(?:${COMP_SOURCE_PATTERN_TEXT})\\s`, "i"))[0]?.trim() || title;
  return {
    source: sourceText.replace(/\s+/g, " ").toUpperCase(),
    title: cleanCompTitle(title),
    date_sold: dateMatch[0],
    sale_type: saleType,
    price,
  };
}

function compDatePattern() {
  return /\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}\b|\b\d{1,2}\/\d{1,2}\/\d{2,4}\b/i;
}

function sourceLabelToPattern(label) {
  return String(label || "")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/-/g, " ")
    .replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
    .replace(/\s+/g, "[\\s-]*");
}

function dedupeComps(comps) {
  const ordered = [];
  for (const raw of comps) {
    if (!raw) continue;
    const comp = {
      ...raw,
      source: cleanCompSource(raw.source),
      title: cleanCompTitle(raw.title),
    };
    if (isJunkCompTitle(comp.title)) continue;

    const price = String(comp.price || "").replace(/[$,\s]/g, "");
    const saleType = String(comp.sale_type || "").replace(/\s+/g, " ").trim().toLowerCase();
    const titleKey = compactCompTitle(comp.title).slice(0, 80);
    const source = cleanCompSource(comp.source).toLowerCase();
    const existingIndex = ordered.findIndex((existing) => {
      const existingPrice = String(existing.price || "").replace(/[$,\s]/g, "");
      if (existingPrice !== price) return false;
      const existingSaleType = String(existing.sale_type || "").replace(/\s+/g, " ").trim().toLowerCase();
      const existingTitleKey = compactCompTitle(existing.title).slice(0, 80);
      const sameSource = cleanCompSource(existing.source).toLowerCase() === source;
      const similarTitle = Boolean(titleKey && existingTitleKey && (titleKey.includes(existingTitleKey) || existingTitleKey.includes(titleKey)));
      const sameDate = normalizeDateText(existing.date_sold) === normalizeDateText(comp.date_sold);
      return (sameDate && (sameSource || similarTitle)) || (sameSaleTypeOrBlank(existingSaleType, saleType) && sameSource && similarTitle);
    });

    if (existingIndex === -1) {
      ordered.push(comp);
      continue;
    }
    const existing = ordered[existingIndex];
    const existingDate = parseCompDate(existing.date_sold);
    const compDate = parseCompDate(comp.date_sold);
    if (existingDate && compDate && compDate < existingDate) {
      ordered[existingIndex] = comp;
    } else if (!existingDate || existingDate?.getTime() === compDate?.getTime()) {
      if (compQuality(comp) > compQuality(existing)) ordered[existingIndex] = comp;
    }
  }
  return ordered;
}

function cleanCompSource(value) {
  return String(value || "").replace(/\s*\(confirmed paid\)\s*/ig, " ").replace(/\s+/g, " ").trim();
}

function cleanCompTitle(value) {
  let title = String(value || "").replace(/\s+/g, " ").trim();
  title = title.replace(/\b(?:close|help[_\s-]*outline|Date Sold|Type|Price)\b/ig, " ");
  title = title.replace(/^\s*[-|:]+\s*/, "").replace(/\s*[-|:]+\s*$/, "");
  return title.replace(/\s+/g, " ").trim();
}

function compactCompTitle(value) {
  return cleanCompTitle(value)
    .toLowerCase()
    .replace(/\b(psa|bgs|sgc|cgc|gem|mint|mt|pop|rookie|rc)\b/g, " ")
    .replace(/[^a-z0-9]+/g, "");
}

function isJunkCompTitle(value) {
  const title = cleanCompTitle(value);
  if (!title) return true;
  if (title.replace(/[^A-Za-z0-9]/g, "").length < 8) return true;
  return !/[A-Za-z]{3,}/.test(title);
}

function compQuality(comp) {
  const title = cleanCompTitle(comp.title);
  let score = Math.min(title.length, 160);
  if (/\b\d{4}\b/.test(title)) score += 20;
  if (/#\s*[A-Za-z0-9-]+|\b[A-Za-z]{1,5}\d{1,4}\b/.test(title)) score += 10;
  if (isJunkCompTitle(title)) score -= 200;
  return score;
}

function sameSaleTypeOrBlank(a, b) {
  return !a || !b || a === b;
}

function normalizeDateText(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function parseCompDate(value) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  const parsed = Date.parse(text);
  return Number.isNaN(parsed) ? null : new Date(parsed);
}

function collectNearbyText(node) {
  const parts = [];
  let current = node;
  for (let i = 0; i < 4 && current; i += 1) {
    parts.push(current.innerText || current.textContent || "");
    current = current.parentElement;
  }
  const rect = node.getBoundingClientRect();
  [...document.querySelectorAll("body *")]
    .filter((el) => isVisible(el))
    .map((el) => ({ el, rect: el.getBoundingClientRect(), text: visibleText(el) }))
    .filter((item) =>
      item.rect.top >= rect.top - 20 &&
      item.rect.top <= rect.bottom + 40 &&
      item.rect.left >= rect.left &&
      item.rect.left <= rect.right + 180
    )
    .forEach((item) => parts.push(item.text));
  return parts.join(" ");
}

async function waitForClValue() {
  for (let i = 0; i < 20; i += 1) {
    if (/\b(?:CL|Card\s*Ladder)\s*Value\b/i.test(document.body.innerText || "")) return;
    await sleep(500);
  }
}

function findClickable(pattern) {
  return [...document.querySelectorAll("button, [role='button'], a, [role='option'], li, div")]
    .find((el) => pattern.test(visibleText(el)));
}

function findSearchInput() {
  const inputs = [...document.querySelectorAll("input:not([type='hidden']), textarea")];
  return inputs.find((el) =>
    `${el.placeholder || ""} ${el.getAttribute("aria-label") || ""} ${el.name || ""}`.match(/search listing titles/i)
  );
}

async function closeGlobalSearchIfOpen() {
  const globalSearch = [...document.querySelectorAll("input:not([type='hidden']), textarea")]
    .find((el) =>
      isVisible(el) &&
      !`${el.placeholder || ""} ${el.getAttribute("aria-label") || ""} ${el.name || ""}`.match(/search listing titles/i) &&
      `${el.placeholder || ""} ${el.getAttribute("aria-label") || ""} ${el.name || ""}`.match(/^ ?search ?$/i)
    );
  if (!globalSearch) return;

  const rect = globalSearch.getBoundingClientRect();
  const overlayClose = [...document.querySelectorAll("button, [role='button'], span, div")]
    .filter((el) => isVisible(el))
    .map((el) => ({ el, rect: el.getBoundingClientRect(), text: visibleText(el) }))
    .filter(({ rect: r, text }) =>
      text === "×" &&
      r.top >= rect.top &&
      r.left >= rect.left &&
      r.left <= rect.right + 40
    )
    .sort((a, b) => a.rect.top - b.rect.top)[0]?.el;

  if (overlayClose) {
    clickLikeHuman(overlayClose);
    await sleep(250);
  }

  globalSearch.blur();
  document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
  await sleep(250);
  clickAtPoint(Math.min(window.innerWidth - 80, rect.right + 160), Math.min(window.innerHeight - 80, rect.bottom + 160));
  await sleep(300);
}

function findHashControlNearSearch(searchInput) {
  const inputRect = searchInput.getBoundingClientRect();
  const candidates = [...document.querySelectorAll("button, [role='button'], span, div, svg, path")]
    .filter((el) => isVisible(el))
    .map((el) => ({ el, rect: el.getBoundingClientRect(), text: (el.innerText || el.textContent || "").trim() }))
    .filter(({ rect }) =>
      rect.top >= inputRect.top - 10 &&
      rect.bottom <= inputRect.bottom + 10 &&
      rect.left >= inputRect.right - 80 &&
      rect.right <= inputRect.right + 20
    )
    .filter(({ text, el }) => text === "#" || (el.getAttribute("aria-label") || "").match(/cert|number|hash|#/i));

  candidates.sort((a, b) => {
    const aExact = a.text === "#" ? 0 : 1;
    const bExact = b.text === "#" ? 0 : 1;
    return aExact - bExact || Math.abs(a.rect.right - inputRect.right) - Math.abs(b.rect.right - inputRect.right);
  });

  return candidates[0]?.el || null;
}

function certSearchModal() {
  const candidates = [...document.querySelectorAll("[role='dialog'], .modal, div")]
    .filter((el) => isVisible(el) && /SEARCH SALES BY CERT #/i.test(visibleText(el)))
    .map((el) => {
      const rect = el.getBoundingClientRect();
      const text = visibleText(el);
      return {
        el,
        rect,
        text,
        area: rect.width * rect.height,
        roleScore: el.getAttribute("role") === "dialog" ? 0 : 1,
      };
    })
    .filter(({ rect, text }) =>
      rect.width >= 300 &&
      rect.width <= Math.min(window.innerWidth, 900) &&
      rect.height >= 180 &&
      rect.height <= Math.min(window.innerHeight, 700) &&
      /Cert #/i.test(text) &&
      /Grader/i.test(text) &&
      /Submit/i.test(text)
    );

  candidates.sort((a, b) => a.roleScore - b.roleScore || a.area - b.area);
  return candidates[0]?.el || null;
}

function findFieldControlInModal(modal, labelPattern, preferredSelector = "input, textarea, [role='combobox'], select, button, div") {
  const controls = [...modal.querySelectorAll(preferredSelector)].filter((el) => isVisible(el));
  const direct = controls.find((el) =>
    `${el.placeholder || ""} ${el.getAttribute("aria-label") || ""} ${el.name || ""} ${visibleText(el)}`.match(labelPattern)
  );
  if (direct && direct.matches("input, textarea, select, [role='combobox'], button")) return direct;

  const labels = [...modal.querySelectorAll("label, legend, span, div")]
    .filter((el) => isVisible(el) && labelPattern.test(visibleText(el)));
  for (const label of labels) {
    const labelRect = label.getBoundingClientRect();
    const below = controls
      .map((el) => ({ el, rect: el.getBoundingClientRect() }))
      .filter(({ rect }) => rect.top >= labelRect.top - 4 && rect.top <= labelRect.bottom + 44)
      .filter(({ rect }) => rect.left >= labelRect.left - 20 && rect.left <= labelRect.right + 560)
      .sort((a, b) => (a.rect.top - b.rect.top) || (a.rect.left - b.rect.left))[0];
    if (below) return below.el;
  }

  if (labelPattern.test("cert")) return controls.find((el) => el.matches("input, textarea"));
  return controls[0] || null;
}

function certInputIsVisible() {
  return [...document.querySelectorAll("input:not([type='hidden']), textarea")]
    .some((el) => isVisible(el) && `${el.placeholder || ""} ${el.getAttribute("aria-label") || ""} ${el.name || ""}`.match(/cert/i));
}

async function clickCertMenuOptionIfShown() {
  await sleep(250);
  const option = [...document.querySelectorAll("button, [role='button'], [role='option'], li, div, span")]
    .filter((el) => isVisible(el))
    .find((el) => /^(#|cert\s*#|certification\s*#|certification number|cert number)$/i.test(visibleText(el)));
  if (!option) return false;
  clickLikeHuman(option);
  await sleep(400);
  return true;
}

function clickAtPoint(x, y, reason = "point") {
  const target = document.elementFromPoint(x, y);
  if (!target) return;
  const candidates = uniqueElements([
    target,
    target.closest("[role='combobox']"),
    target.closest("[aria-haspopup='listbox']"),
    target.closest("[aria-expanded]"),
    target.closest("i, svg, [class*='material-icons'], [class*='arrow']"),
    target.closest(".select-input"),
    target.closest("[class*='select-input']"),
    target.closest("[class*='select']"),
    target.closest("[class*='dropdown']"),
    target.closest("[class*='field']"),
    target.closest(".MuiSelect-select"),
    target.closest(".MuiInputBase-root"),
    target.closest("button, [role='button']"),
    target.closest("label"),
    ...clickableAncestors(target, 5),
  ]).filter(Boolean);
  const clickNodes = candidates.length ? candidates.slice(0, 8) : [target];
  rememberGraderOpenDebug(`click-${reason}`, target, null, `${Math.round(x)},${Math.round(y)}`);
  for (const node of clickNodes) {
    clickLikeHuman(node, x, y, reason);
  }
}

function clickableAncestors(node, maxDepth = 4) {
  const ancestors = [];
  let current = node?.parentElement;
  while (current && ancestors.length < maxDepth) {
    if (isVisible(current)) ancestors.push(current);
    current = current.parentElement;
  }
  return ancestors;
}

function clickLikeHuman(node, clientX = null, clientY = null, reason = "click") {
  if (!node) return;
  const rect = node.getBoundingClientRect();
  const x = clientX ?? rect.left + rect.width / 2;
  const y = clientY ?? rect.top + rect.height / 2;
  rememberGraderOpenDebug(reason, node, rect, `${Math.round(x)},${Math.round(y)}`);
  if (typeof node.focus === "function") node.focus();
  ["pointerover", "mouseover", "pointerenter", "mouseenter", "pointermove", "mousemove", "pointerdown", "mousedown", "pointerup", "mouseup", "click"].forEach((type) => {
    const pressed = /down|move/.test(type);
    if (type.startsWith("pointer") && window.PointerEvent) {
      node.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, composed: true, view: window, clientX: x, clientY: y, button: 0, buttons: pressed ? 1 : 0, pointerId: 1, pointerType: "mouse", isPrimary: true }));
    } else {
      node.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, composed: true, view: window, clientX: x, clientY: y, button: 0, buttons: pressed ? 1 : 0, detail: 1 }));
    }
  });
  if (typeof node.click === "function") node.click();
  node.dispatchEvent(new MouseEvent("dblclick", { bubbles: true, cancelable: true, composed: true, view: window, clientX: x, clientY: y, button: 0, detail: 2 }));
  invokeFrameworkHandlers(node, "mouseDown", { clientX: x, clientY: y });
  invokeFrameworkHandlers(node, "click", { clientX: x, clientY: y });
}

function invokeFrameworkHandlers(node, eventKind, init = {}) {
  const eventNamesByKind = {
    mouseDown: ["onPointerDown", "onMouseDown"],
    click: ["onClick"],
    keyDown: ["onKeyDown"],
    keyUp: ["onKeyUp"],
  };
  const eventNames = eventNamesByKind[eventKind] || [];
  const candidates = uniqueElements([node, ...clickableAncestors(node, 6)]);
  for (const candidate of candidates) {
    const props = frameworkProps(candidate);
    if (!props) continue;
    for (const eventName of eventNames) {
      const handler = props[eventName] || props[eventName.toLowerCase?.()];
      if (typeof handler !== "function") continue;
      try {
        rememberGraderOpenDebug(`handler-${eventName}`, candidate);
        handler(fakeFrameworkEvent(candidate, eventKind, init));
      } catch (error) {
        rememberGraderOpenDebug(`handler-error-${eventName}`, candidate, null, String(error?.message || error).slice(0, 80));
      }
    }
  }
}

function frameworkProps(node) {
  if (!node) return null;
  const directKey = Object.keys(node).find((key) =>
    /^__reactProps\$|^__reactEventHandlers\$/.test(key)
  );
  if (directKey && node[directKey]) return node[directKey];
  const fiberKey = Object.keys(node).find((key) =>
    /^__reactFiber\$|^__reactInternalInstance\$/.test(key)
  );
  const fiber = fiberKey ? node[fiberKey] : null;
  return fiber?.memoizedProps || fiber?.return?.memoizedProps || null;
}

function fakeFrameworkEvent(target, eventKind, init = {}) {
  const event = {
    ...init,
    target,
    currentTarget: target,
    bubbles: true,
    cancelable: true,
    defaultPrevented: false,
    isTrusted: true,
    nativeEvent: null,
    preventDefault() {
      this.defaultPrevented = true;
    },
    stopPropagation() {},
    persist() {},
  };
  event.nativeEvent = event;
  if (eventKind === "keyDown" || eventKind === "keyUp") {
    event.key = init.key || "ArrowDown";
    event.code = init.code || event.key;
    event.altKey = Boolean(init.altKey);
  } else {
    event.button = 0;
    event.buttons = eventKind === "mouseDown" ? 1 : 0;
    event.type = eventKind === "mouseDown" ? "mousedown" : "click";
  }
  return event;
}

function uniqueElements(items) {
  const seen = new Set();
  const elements = [];
  for (const item of items) {
    if (!item || seen.has(item)) continue;
    seen.add(item);
    elements.push(item);
  }
  return elements;
}

function rememberGraderOpenDebug(label, node, rect = null, extra = "") {
  if (!node && !extra) return;
  const bits = [label];
  if (extra) bits.push(extra);
  if (node) bits.push(describeNode(node, rect));
  lastGraderOpenDebug.push(bits.filter(Boolean).join(" "));
  if (lastGraderOpenDebug.length > 18) {
    lastGraderOpenDebug = lastGraderOpenDebug.slice(-18);
  }
}

function describeNode(node, rect = null) {
  const box = rect || node.getBoundingClientRect?.();
  const classes = String(node.getAttribute?.("class") || "")
    .replace(/\s+/g, ".")
    .slice(0, 80);
  const attrs = [
    node.tagName?.toLowerCase() || "node",
    node.getAttribute?.("role") ? `role=${node.getAttribute("role")}` : "",
    node.getAttribute?.("aria-expanded") ? `expanded=${node.getAttribute("aria-expanded")}` : "",
    node.getAttribute?.("aria-haspopup") ? `popup=${node.getAttribute("aria-haspopup")}` : "",
    classes ? `class=${classes}` : "",
    box ? `rect=${Math.round(box.left)},${Math.round(box.top)},${Math.round(box.width)}x${Math.round(box.height)}` : "",
  ].filter(Boolean).join("/");
  const text = visibleText(node).replace(/\s+/g, " ").slice(0, 70);
  return `${attrs}${text ? ` text=${text}` : ""}`;
}

function setNativeValue(input, value) {
  const descriptor = Object.getOwnPropertyDescriptor(input.constructor.prototype, "value");
  if (descriptor?.set) descriptor.set.call(input, value);
  else input.value = value;
  input.dispatchEvent(new Event("input", { bubbles: true }));
}

function visibleText(el) {
  const style = getComputedStyle(el);
  const box = el.getBoundingClientRect();
  if (style.display === "none" || style.visibility === "hidden" || box.width === 0 || box.height === 0) return "";
  return (el.innerText || el.textContent || el.getAttribute("aria-label") || el.title || "").trim();
}

function isVisible(node) {
  const rect = node.getBoundingClientRect();
  const style = getComputedStyle(node);
  return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
}

function pageClue() {
  return String(document.body?.innerText || "").replace(/\s+/g, " ").slice(0, 1500);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

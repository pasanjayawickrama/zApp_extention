const output = document.getElementById("output");
const setupView = document.getElementById("setup");
const mainView = document.getElementById("main");
const settingsView = document.getElementById("settings");
const topicInput = document.getElementById("topic");
const linkInput = document.getElementById("link");
const saveButton = document.getElementById("save");
const doneButton = document.getElementById("done");
const changeButton = document.getElementById("change");
const openSettingsFromMainButton = document.getElementById("openSettingsFromMain");
const backFromSettingsButton = document.getElementById("backFromSettings");
const setupError = document.getElementById("setupError");
const savedLinksEl = document.getElementById("savedLinks");
const removeAllButton = document.getElementById("removeAll");
const refreshMinutesInput = document.getElementById("refreshMinutes");
const refreshSecondsInput = document.getElementById("refreshSeconds");
const saveRefreshRateButton = document.getElementById("saveRefreshRate");
const notifyAllEnabledInput = document.getElementById("notifyAllEnabled");
const notifyP1Input = document.getElementById("notifyP1");
const notifyP2Input = document.getElementById("notifyP2");
const notifyP3Input = document.getElementById("notifyP3");
const notifyP4Input = document.getElementById("notifyP4");
const notifyP5Input = document.getElementById("notifyP5");
const notifyUnknownInput = document.getElementById("notifyUnknown");
const saveNotifySettingsButton = document.getElementById("saveNotifySettings");
const notifySettingsStatus = document.getElementById("notifySettingsStatus");
const settingsError = document.getElementById("settingsError");
const refreshRateStatus = document.getElementById("refreshRateStatus");
const cardEl = document.querySelector(".card");
const mainStatusEl = document.getElementById("mainStatus");

const PRIORITY_BUCKETS = [
  "1 - Critical",
  "2 - High",
  "3 - Moderate",
  "4 - Low",
  "5 - Planning"
];

function makeZeroCounts() {
  const out = { Unknown: 0 };
  for (const bucket of PRIORITY_BUCKETS) out[bucket] = 0;
  return out;
}

const STORAGE_KEY = "queueConfigV1";
const REFRESH_RATE_KEY = "refreshRateV1";
const NOTIFY_SETTINGS_KEY = "notifySettingsV1";
const QUEUE_EXPAND_KEY = "queueExpandStateV1";
const DEFAULT_REFRESH_SECONDS = 30;

let currentConfig = null;
let lastNonSettingsView = "main";
let refreshRateLoadedSeconds = null;
let notifySettingsLoaded = null;
let mainAutoRefreshTimerId = null;
let mainAutoRefreshInFlight = false;
let mainAutoRefreshSeconds = null;

function setCardMode(mode) {
  if (!cardEl) return;
  cardEl.classList.remove("mode-main", "mode-setup", "mode-settings");
  if (mode === "main") cardEl.classList.add("mode-main");
  else if (mode === "setup") cardEl.classList.add("mode-setup");
  else if (mode === "settings") cardEl.classList.add("mode-settings");
}

function setRefreshRateStatus(text) {
  if (!refreshRateStatus) return;
  const t = String(text || "").trim();
  if (!t) {
    refreshRateStatus.hidden = true;
    refreshRateStatus.textContent = "";
    return;
  }
  refreshRateStatus.hidden = false;
  refreshRateStatus.textContent = t;
}

function setNotifySettingsStatus(text) {
  if (!notifySettingsStatus) return;
  const t = String(text || "").trim();
  if (!t) {
    notifySettingsStatus.hidden = true;
    notifySettingsStatus.textContent = "";
    return;
  }
  notifySettingsStatus.hidden = false;
  notifySettingsStatus.textContent = t;
}

function normalizeNotifySettings(raw) {
  const defaults = {
    enabled: true,
    priorities: {
      "1 - Critical": true,
      "2 - High": true,
      "3 - Moderate": true,
      "4 - Low": true,
      "5 - Planning": true,
      Unknown: true
    }
  };

  if (!raw || typeof raw !== "object") return defaults;
  const enabled = raw.enabled !== false;
  const p = raw.priorities && typeof raw.priorities === "object" ? raw.priorities : {};
  return {
    enabled,
    priorities: {
      "1 - Critical": p["1 - Critical"] !== false,
      "2 - High": p["2 - High"] !== false,
      "3 - Moderate": p["3 - Moderate"] !== false,
      "4 - Low": p["4 - Low"] !== false,
      "5 - Planning": p["5 - Planning"] !== false,
      Unknown: p.Unknown !== false
    }
  };
}

function readNotifySettingsFromUI() {
  const enabled = Boolean(notifyAllEnabledInput?.checked);
  return {
    enabled,
    priorities: {
      "1 - Critical": Boolean(notifyP1Input?.checked),
      "2 - High": Boolean(notifyP2Input?.checked),
      "3 - Moderate": Boolean(notifyP3Input?.checked),
      "4 - Low": Boolean(notifyP4Input?.checked),
      "5 - Planning": Boolean(notifyP5Input?.checked),
      Unknown: Boolean(notifyUnknownInput?.checked)
    }
  };
}

function applyNotifyEnabledStateToUI(masterEnabled) {
  const disabled = !masterEnabled;
  const inputs = [notifyP1Input, notifyP2Input, notifyP3Input, notifyP4Input, notifyP5Input, notifyUnknownInput];
  for (const el of inputs) {
    if (el) el.disabled = disabled;
  }
}

function updateNotifySettingsDirtyUI() {
  if (!saveNotifySettingsButton) return;
  if (!notifyAllEnabledInput) return;

  const current = readNotifySettingsFromUI();
  applyNotifyEnabledStateToUI(current.enabled);

  const loaded = notifySettingsLoaded ? normalizeNotifySettings(notifySettingsLoaded) : null;
  const normalizedCurrent = normalizeNotifySettings(current);
  const dirty = loaded ? JSON.stringify(loaded) !== JSON.stringify(normalizedCurrent) : false;

  saveNotifySettingsButton.disabled = !dirty;
  saveNotifySettingsButton.classList.toggle("buttonDirty", dirty);
  setNotifySettingsStatus(dirty ? "Unsaved changes" : "");
}

async function loadNotifySettingsToUI() {
  if (!notifyAllEnabledInput) return;
  const stored = await storageGet(NOTIFY_SETTINGS_KEY);
  const s = normalizeNotifySettings(stored);
  notifySettingsLoaded = s;

  notifyAllEnabledInput.checked = s.enabled;
  if (notifyP1Input) notifyP1Input.checked = s.priorities["1 - Critical"] !== false;
  if (notifyP2Input) notifyP2Input.checked = s.priorities["2 - High"] !== false;
  if (notifyP3Input) notifyP3Input.checked = s.priorities["3 - Moderate"] !== false;
  if (notifyP4Input) notifyP4Input.checked = s.priorities["4 - Low"] !== false;
  if (notifyP5Input) notifyP5Input.checked = s.priorities["5 - Planning"] !== false;
  if (notifyUnknownInput) notifyUnknownInput.checked = s.priorities.Unknown !== false;

  applyNotifyEnabledStateToUI(s.enabled);
  if (saveNotifySettingsButton) saveNotifySettingsButton.disabled = true;
  if (saveNotifySettingsButton) saveNotifySettingsButton.classList.remove("buttonDirty");
  setNotifySettingsStatus("");
}

function getRefreshRateSecondsFromInputs() {
  const minutes = clampInt(refreshMinutesInput?.value, { min: 0, max: 1440 });
  const seconds = clampInt(refreshSecondsInput?.value, { min: 0, max: 1_000_000 });
  const totalSeconds = minutes * 60 + seconds;
  if (!totalSeconds) return { totalSeconds: 0, clampedSeconds: 0 };
  const clampedSeconds = Math.max(5, Math.min(3600, totalSeconds));
  return { totalSeconds, clampedSeconds };
}

function updateRefreshRateDirtyUI() {
  if (!saveRefreshRateButton) return;

  const loaded = typeof refreshRateLoadedSeconds === "number" ? refreshRateLoadedSeconds : null;
  const { totalSeconds, clampedSeconds } = getRefreshRateSecondsFromInputs();
  const dirty = loaded !== null && clampedSeconds !== 0 && clampedSeconds !== loaded;
  const canSave = dirty && totalSeconds > 0;

  saveRefreshRateButton.disabled = !canSave;
  saveRefreshRateButton.classList.toggle("buttonDirty", dirty);
  setRefreshRateStatus(dirty ? "Unsaved changes" : "");
}

function showSettings() {
  setCardMode("settings");
  setupView.hidden = true;
  mainView.hidden = true;
  if (settingsView) settingsView.hidden = false;
  stopMainAutoRefresh();
}

function showSettingsError(errorText) {
  if (!settingsError) return;
  if (typeof errorText === "string" && errorText.trim()) {
    settingsError.hidden = false;
    settingsError.textContent = errorText;
  } else {
    settingsError.hidden = true;
    settingsError.textContent = "";
  }
}

function clampInt(value, { min = 0, max = 1_000_000 } = {}) {
  const n = Number.parseInt(String(value ?? "0"), 10);
  if (!Number.isFinite(n)) return min;
  return Math.max(min, Math.min(max, n));
}

function normalizeRefreshRateInputs() {
  if (!refreshMinutesInput || !refreshSecondsInput) return;

  const minutesRaw = clampInt(refreshMinutesInput.value, { min: 0, max: 1440 });
  const secondsRaw = clampInt(refreshSecondsInput.value, { min: 0, max: 1_000_000 });

  if (secondsRaw < 60) return;

  const extraMinutes = Math.floor(secondsRaw / 60);
  const seconds = secondsRaw % 60;
  const minutes = Math.min(1440, minutesRaw + extraMinutes);

  const nextMinutes = String(minutes);
  const nextSeconds = String(seconds);
  if (refreshMinutesInput.value !== nextMinutes) refreshMinutesInput.value = nextMinutes;
  if (refreshSecondsInput.value !== nextSeconds) refreshSecondsInput.value = nextSeconds;
}

async function loadRefreshRateToUI() {
  if (!refreshMinutesInput || !refreshSecondsInput) return;
  const stored = await storageGet(REFRESH_RATE_KEY);
  const totalSecondsRaw = typeof stored === "number" ? stored : Number.parseInt(String(stored ?? ""), 10);
  const totalSeconds =
    Number.isFinite(totalSecondsRaw) && totalSecondsRaw > 0
      ? Math.max(5, Math.min(3600, totalSecondsRaw))
      : 30;
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  refreshMinutesInput.value = String(minutes);
  refreshSecondsInput.value = String(seconds);
  refreshRateLoadedSeconds = totalSeconds;

  if (saveRefreshRateButton) saveRefreshRateButton.disabled = true;
  if (saveRefreshRateButton) saveRefreshRateButton.classList.remove("buttonDirty");
  setRefreshRateStatus("");
}

function sendSetRefreshRate(totalSeconds) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "SET_REFRESH_RATE", totalSeconds }, (response) => {
      const err = chrome.runtime.lastError;
      if (err) {
        resolve({ success: false, error: err.message || String(err) });
        return;
      }
      resolve(response);
    });
  });
}

function showSetup(errorText) {
  lastNonSettingsView = "setup";
  stopMainAutoRefresh();
  setCardMode("setup");
  setupView.hidden = false;
  mainView.hidden = true;
  if (settingsView) settingsView.hidden = true;
  if (typeof errorText === "string" && errorText.trim()) {
    setupError.hidden = false;
    setupError.textContent = errorText;
  } else {
    setupError.hidden = true;
    setupError.textContent = "";
  }
}

function showMain(topic) {
  lastNonSettingsView = "main";
  setCardMode("main");
  setupView.hidden = true;
  mainView.hidden = false;
  if (settingsView) settingsView.hidden = true;
  startMainAutoRefresh();
}

function setMainStatus(text) {
  if (!mainStatusEl) return;
  mainStatusEl.textContent = String(text || "");
}

function setMainLoading(isLoading, { statusText } = {}) {
  if (output) output.classList.toggle("isLoading", Boolean(isLoading));
  if (isLoading) {
    // Keep header clean; show loading in the main content area.
    setMainStatus("");
    if (output) output.textContent = statusText || "Loading counts…";
    return;
  }

  if (typeof statusText === "string") {
    setMainStatus(statusText);
    return;
  }

  try {
    const time = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
    setMainStatus(`Updated ${time}`);
  } catch {
    setMainStatus("Updated");
  }
}

function showLoginRequiredMessage() {
  const msg = [
    "Login required.",
    "",
    "1) Open https://support.ifs.com in a tab",
    "2) Sign in",
    "3) Re-open this popup"
  ].join("\n");
  if (output) output.textContent = msg;
  setMainLoading(false, { statusText: "Login required" });
}

function clampSeconds(seconds) {
  const n = typeof seconds === "number" ? seconds : Number.parseInt(String(seconds ?? ""), 10);
  if (!Number.isFinite(n) || n <= 0) return DEFAULT_REFRESH_SECONDS;
  return Math.max(5, Math.min(3600, n));
}

async function getRefreshRateSeconds() {
  const stored = await storageGet(REFRESH_RATE_KEY);
  return clampSeconds(stored);
}

function stopMainAutoRefresh() {
  if (mainAutoRefreshTimerId) {
    clearInterval(mainAutoRefreshTimerId);
    mainAutoRefreshTimerId = null;
  }
  mainAutoRefreshSeconds = null;
}

async function startMainAutoRefresh() {
  if (!mainView || mainView.hidden) return;
  const seconds = await getRefreshRateSeconds();
  if (mainAutoRefreshTimerId && mainAutoRefreshSeconds === seconds) return;

  stopMainAutoRefresh();
  mainAutoRefreshSeconds = seconds;

  mainAutoRefreshTimerId = setInterval(async () => {
    if (!mainView || mainView.hidden) return;
    if (mainAutoRefreshInFlight) return;
    mainAutoRefreshInFlight = true;
    try {
      await run({ skipNotify: true });
    } finally {
      mainAutoRefreshInFlight = false;
    }
  }, seconds * 1000);
}

chrome.storage?.onChanged?.addListener?.((changes, areaName) => {
  if (areaName !== "sync") return;
  if (!changes || !Object.prototype.hasOwnProperty.call(changes, REFRESH_RATE_KEY)) return;
  if (!mainView || mainView.hidden) return;
  startMainAutoRefresh();
});

function computeTotalCount(counts) {
  if (!counts || typeof counts !== "object") return 0;
  let total = 0;
  for (const k of Object.keys(counts)) {
    const v = counts[k];
    const n = typeof v === "number" ? v : Number.parseInt(String(v ?? "0"), 10);
    if (Number.isFinite(n)) total += n;
  }
  return total;
}

function storageGet(key) {
  return new Promise((resolve) => {
    chrome.storage.sync.get([key], (res) => resolve(res?.[key] ?? null));
  });
}

function storageSet(obj) {
  return new Promise((resolve) => {
    chrome.storage.sync.set(obj, () => resolve());
  });
}

// Debounced writers to reduce storage.sync write frequency
let saveConfigPending = null;
let saveConfigTimerId = null;
function scheduleSaveConfig(cfg, delayMs = 600) {
  saveConfigPending = cfg;
  if (saveConfigTimerId) clearTimeout(saveConfigTimerId);
  saveConfigTimerId = setTimeout(async () => {
    const toSave = saveConfigPending;
    saveConfigPending = null;
    saveConfigTimerId = null;
    if (toSave) await saveConfig(toSave);
  }, delayMs);
}

async function flushScheduledConfig() {
  if (saveConfigTimerId) {
    clearTimeout(saveConfigTimerId);
    saveConfigTimerId = null;
  }
  const toSave = saveConfigPending;
  saveConfigPending = null;
  if (toSave) await saveConfig(toSave);
}

let saveExpandPending = null;
let saveExpandTimerId = null;
function scheduleSaveExpandState(state, delayMs = 600) {
  saveExpandPending = state;
  if (saveExpandTimerId) clearTimeout(saveExpandTimerId);
  saveExpandTimerId = setTimeout(async () => {
    const toSave = saveExpandPending;
    saveExpandPending = null;
    saveExpandTimerId = null;
    if (toSave) await saveExpandState(toSave);
  }, delayMs);
}

async function flushScheduledExpandState() {
  if (saveExpandTimerId) {
    clearTimeout(saveExpandTimerId);
    saveExpandTimerId = null;
  }
  const toSave = saveExpandPending;
  saveExpandPending = null;
  if (toSave) await saveExpandState(toSave);
}

async function loadExpandState() {
  const raw = await storageGet(QUEUE_EXPAND_KEY);
  return raw && typeof raw === "object" ? raw : {};
}

async function saveExpandState(state) {
  await storageSet({ [QUEUE_EXPAND_KEY]: state });
}

function makeEmptyConfig() {
  return { links: [] };
}

function normalizeConfig(raw) {
  if (!raw || typeof raw !== "object") return makeEmptyConfig();

  // New format
  if (Array.isArray(raw.links)) {
    return {
      links: raw.links
        .filter((x) => x && typeof x === "object")
        .map((x) => ({
          id: String(x.id || "").trim() || String(Date.now()),
          topic: String(x.topic || "").trim(),
          link: String(x.link || "").trim(),
          notifyEnabled: x.notifyEnabled !== false,
          pinned: x.pinned === true
        }))
        .filter((x) => x.topic && x.link)
    };
  }

  // Legacy format: { topic, link }
  if (raw.topic && raw.link) {
    return {
      links: [
        {
          id: "legacy",
          topic: String(raw.topic).trim(),
          link: String(raw.link).trim(),
          notifyEnabled: true,
          pinned: false
        }
      ]
    };
  }

  return makeEmptyConfig();
}

async function loadConfig() {
  const raw = await storageGet(STORAGE_KEY);
  const cfg = normalizeConfig(raw);
  // If we upgraded legacy -> new format, persist it.
  if (raw && raw.topic && raw.link && !raw.links) {
    await storageSet({ [STORAGE_KEY]: cfg });
  }
  return cfg;
}

async function saveConfig(cfg) {
  currentConfig = cfg;
  await storageSet({ [STORAGE_KEY]: cfg });
}

function renderSavedLinks(cfg) {
  if (!savedLinksEl) return;
  const links = cfg?.links || [];
  if (!links.length) {
    savedLinksEl.innerHTML = `<div class="hint" style="margin-top:0;">No links added yet.</div>`;
    return;
  }

  const esc = (s) => String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;");

  savedLinksEl.innerHTML = links
    .map(
      (l) => {
        const notifyOn = l?.notifyEnabled !== false;
        const bell = notifyOn ? "\uD83D\uDD14" : "\uD83D\uDD15";
        const bellTitle = notifyOn ? "Notifications on" : "Notifications off";
        return `
        <div class="savedLinkRow ${notifyOn ? "" : "savedLinkRowDisabled"}" data-id="${esc(l.id)}">
          <div class="savedLinkTopic">${esc(l.topic)}</div>
          <div class="savedLinkUrl">${esc(l.link)}</div>
          <div class="savedLinkActions">
            <button class="moveButton" type="button" data-move-up="${esc(l.id)}" aria-label="Move up" title="Move up">▲</button>
            <button class="moveButton" type="button" data-move-down="${esc(l.id)}" aria-label="Move down" title="Move down">▼</button>
            <button
              class="bellButton"
              type="button"
              data-toggle-notify="${esc(l.id)}"
              data-state="${notifyOn ? "on" : "off"}"
              aria-label="Toggle notifications"
              title="${esc(bellTitle)}"
            >${bell}</button>
            <button class="removeButton" type="button" data-remove="${esc(l.id)}">Remove</button>
          </div>
        </div>
      `;
      }
    )
    .join("\n");
}

function safeJsonParse(value) {
  if (typeof value !== "string") return null;
  try {
    return JSON.parse(value);
  } catch {
    return null;
  }
}

function escapeHtml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;");
}

function escapeAttr(text) {
  // For simple attribute contexts like href/title.
  return escapeHtml(text).replace(/'/g, "&#39;");
}

function normalizeHostname(hostname) {
  return String(hostname || "").trim().toLowerCase();
}

function extractTableFromPath(pathname) {
  const name = String(pathname || "").split("/").pop() || "";
  // sn_customerservice_case_list.do
  const m = /^([a-z0-9_]+)_list\.do$/i.exec(name);
  return m ? m[1] : null;
}

function extractListIdTinyFromAgentListPath(pathname) {
  // Example:
  // /now/cwf/agent/list/params/list-id/<LIST_ID>/tiny-id/<TINY_ID>
  const p = String(pathname || "");
  const m = /\/now\/cwf\/agent\/list\/params\/list-id\/([^\/]+)\/tiny-id\/([^\/]+)/i.exec(p);
  if (!m) return null;
  return { listId: m[1], tinyId: m[2] };
}

function getTypeBadgeForTable(tableName) {
  const t = String(tableName || "").trim().toLowerCase();
  if (!t) return null;
  if (t === "sn_customerservice_case") return { letter: "C", cls: "typeBadgeCase", label: "Case" };
  if (t === "sn_customerservice_task") return { letter: "T", cls: "typeBadgeTask", label: "Task" };
  return null;
}

function tryParseUrl(raw) {
  try {
    return new URL(raw);
  } catch {
    return null;
  }
}

function decodeMaybeEncoded(value, maxPasses = 3) {
  let current = String(value ?? "");
  for (let i = 0; i < maxPasses; i += 1) {
    try {
      const next = decodeURIComponent(current);
      if (next === current) break;
      current = next;
    } catch {
      break;
    }
  }
  return current;
}

function getSysparmQueryFromUrl(url) {
  // Sometimes links use sysparm_fixed_query instead of sysparm_query.
  return url.searchParams.get("sysparm_query") || url.searchParams.get("sysparm_fixed_query");
}

function getSysparmQueryFromHash(url) {
  const hash = String(url.hash || "").replace(/^#/, "");
  if (!hash) return null;

  // Some SPA routes put params after '#', sometimes including an embedded '?'.
  const maybeQueryString = hash.includes("?") ? hash.split("?").slice(1).join("?") : hash;
  const params = new URLSearchParams(maybeQueryString);
  return params.get("sysparm_query") || params.get("sysparm_fixed_query");
}

function parseServiceNowListLink(rawLink) {
  const raw = String(rawLink || "").trim();
  if (!raw) return { error: "Please paste a link." };

  const url = tryParseUrl(raw);
  if (!url) return { error: "That doesn't look like a valid URL." };

  const host = normalizeHostname(url.hostname);
  if (host !== "support.ifs.com") {
    return { error: "Link must be on https://support.ifs.com" };
  }

  // Case 0: modern "agent list" link containing list-id/tiny-id (no sysparm_query in URL)
  const agentList = extractListIdTinyFromAgentListPath(url.pathname);
  if (agentList?.listId && agentList?.tinyId) {
    return { listId: agentList.listId, tinyId: agentList.tinyId };
  }

  // Case A: direct list link like .../sn_customerservice_case_list.do?sysparm_query=...
  const directTable = extractTableFromPath(url.pathname);
  const directQuery = getSysparmQueryFromUrl(url) || getSysparmQueryFromHash(url);
  if (directTable && directQuery) {
    return { table: directTable, query: directQuery };
  }

  // Case B: classic nav wrapper. e.g. .../now/nav/ui/classic/params/target/<encoded>
  // or target=<encoded> in search params.
  const targetFromParams = url.searchParams.get("target");
  const targetFromPath = (() => {
    const marker = "/now/nav/ui/classic/params/target/";
    const idx = url.pathname.indexOf(marker);
    if (idx === -1) return null;
    return url.pathname.slice(idx + marker.length);
  })();

  const maybeEncodedTarget = targetFromParams || targetFromPath;
  if (maybeEncodedTarget) {
    const decoded = decodeMaybeEncoded(maybeEncodedTarget, 3);
    const inner = tryParseUrl(
      decoded.startsWith("http") ? decoded : `https://support.ifs.com/${decoded.replace(/^\/+/, "")}`
    );
    if (inner) {
      const innerTable = extractTableFromPath(inner.pathname);
      const innerQuery = getSysparmQueryFromUrl(inner) || getSysparmQueryFromHash(inner);
      if (innerTable && innerQuery) {
        return { table: innerTable, query: innerQuery };
      }
    }
  }

  // Case C: already a table API link
  const apiMatch = /^\/api\/now\/table\/([^\/]+)\/?$/i.exec(url.pathname);
  if (apiMatch) {
    const apiTable = apiMatch[1];
    const apiQuery = getSysparmQueryFromUrl(url) || getSysparmQueryFromHash(url);
    if (apiQuery) return { table: apiTable, query: apiQuery };
  }

  return {
    error:
      "Couldn't find a sysparm_query in that link.\n\nTip: some ServiceNow 'now' links don't include the filter in the URL. Open the list, apply the filter, and copy a URL that includes something like '..._list.do?sysparm_query=...'."
  };
}

function computePriorityCounts(graphqlResponseData) {
  const counts = {
    "1 - Critical": 0,
    "2 - High": 0,
    "3 - Moderate": 0,
    "4 - Low": 0,
    "5 - Planning": 0,
    Unknown: 0
  };

  const results =
    graphqlResponseData?.data?.GlideRecord_Query?.ui_notification_inbox?._results || [];

  for (const item of results) {
    const payloadValue = item?.payload?.value;
    const payloadObj = safeJsonParse(payloadValue);

    const recordNumber = payloadObj?.displayValue;
    const priorityDisplay =
      typeof payloadObj?.changes?.priority?.displayValue === "string"
        ? payloadObj.changes.priority.displayValue.trim()
        : payloadObj?.changes?.priority?.displayValue;

    // We don't render case numbers (per requirement), but we do a quick sanity check
    // so malformed payloads don't get counted under a real bucket.
    const looksLikeRecord =
      typeof recordNumber === "string" && /^(CS\d+|CSTASK\d+)/i.test(recordNumber.trim());

    if (!looksLikeRecord || typeof priorityDisplay !== "string") {
      counts.Unknown += 1;
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(counts, priorityDisplay)) {
      counts[priorityDisplay] += 1;
    } else {
      counts.Unknown += 1;
    }
  }

  return counts;
}

function computePriorityCountsFromListLayout(graphqlResponseData) {
  const counts = {
    "1 - Critical": 0,
    "2 - High": 0,
    "3 - Moderate": 0,
    "4 - Low": 0,
    "5 - Planning": 0,
    Unknown: 0
  };

  const rows =
    graphqlResponseData?.data?.GlideListLayout_Query?.getTinyListLayout?.layoutQuery?.queryRows ||
    graphqlResponseData?.data?.GlideListLayout_Query?.getListLayout?.layoutQuery?.queryRows ||
    [];

  for (const row of rows) {
    // Only handle non-grouped rows
    const rowData = row?.rowData;
    if (!Array.isArray(rowData)) {
      continue;
    }

    const priorityCell = rowData.find((c) => c?.columnName === "priority");
    const priorityDisplayRaw = priorityCell?.columnData?.displayValue;
    const priorityDisplay = typeof priorityDisplayRaw === "string" ? priorityDisplayRaw.trim() : null;
    if (!priorityDisplay) {
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(counts, priorityDisplay)) {
      counts[priorityDisplay] += 1;
    } else {
      counts.Unknown += 1;
    }
  }

  return counts;
}

function formatCountsBlock(title, counts, linkUrl, typeBadge, options = {}) {
  const total = computeTotalCount(counts);
  const safeTitle = escapeHtml(title);

  const safeLink = (() => {
    const url = tryParseUrl(String(linkUrl || "").trim());
    if (!url) return null;
    return url.toString();
  })();

  const badgeHtml = typeBadge
    ? `<span class="typeBadge ${escapeAttr(typeBadge.cls)}" title="${escapeAttr(typeBadge.label)}" aria-label="${escapeAttr(typeBadge.label)}">${escapeHtml(typeBadge.letter)}</span>`
    : "";
  const headerInner = `${badgeHtml}<span class="topicHeaderText">${safeTitle}</span><span class="totalPill" aria-label="Total">${Number.isFinite(total) ? total : 0}</span>`;
  const header = safeLink
    ? `<strong><a class="topicLink" href="${escapeAttr(safeLink)}" target="_blank" rel="noreferrer">${headerInner}</a></strong>`
    : `<strong>${headerInner}</strong>`;

  const queueId = String(options?.queueId || "");
  const expanded = options?.expanded === true;
  const caretLabel = expanded ? "Collapse" : "Expand";

  const lines = [
    `<div class="queueBlock ${expanded ? "expanded" : "collapsed"}" ${queueId ? `data-qid="${escapeAttr(queueId)}"` : ""}>`,
    `<div class="topicHeader">` +
      `<button class="toggleButton" type="button" aria-label="${escapeAttr(caretLabel)} queue" aria-expanded="${expanded ? "true" : "false"}" title="${escapeAttr(caretLabel)}">` +
        `<svg class="toggleIcon ${expanded ? "isExpanded" : ""}" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">` +
          `<path d="M9 6l6 6-6 6"/>` +
        `</svg>` +
      `</button>` +
      `${header}` +
    `</div>`,
    `<div class="queueBody">`
  ];
  for (const bucket of PRIORITY_BUCKETS) {
    const nRaw = Number(counts?.[bucket] ?? 0);
    const n = Number.isFinite(nRaw) ? nRaw : 0;
    const isActive = n > 0;
    const dot = isActive ? "●" : "○";
    const name = bucket.replace(/^\d+\s*-\s*/g, "");
    const prioClass =
      bucket === "1 - Critical" ? "prioCritical" : bucket === "2 - High" ? "prioHigh" : "prioPurple";
    const numHtml = isActive ? String(n) : ".";

    lines.push(
      `<div class="prioLine ${prioClass} ${isActive ? "isActive" : ""}">` +
        `<span class="prioDot" aria-hidden="true">${dot}</span>` +
        `<span class="prioName">${escapeHtml(name)}</span>` +
        `<span class="prioNum ${isActive ? "" : "isZero"}">${numHtml}</span>` +
      `</div>`
    );
  }

  const unknownRaw = Number(counts?.Unknown ?? 0);
  const unknown = Number.isFinite(unknownRaw) ? unknownRaw : 0;
  if (unknown) {
    lines.push(
      `<div class="prioLine prioPurple isActive">` +
        `<span class="prioDot" aria-hidden="true">●</span>` +
        `<span class="prioName">Unknown</span>` +
        `<span class="prioNum">${unknown}</span>` +
      `</div>`
    );
  }

  // Avoid inserting literal newlines inside the <pre> container.
  // Newlines can dominate perceived spacing (line-height) and hide small margin tweaks.
  lines.push(`</div>`); // queueBody
  lines.push(`</div>`); // queueBlock
  return lines.join("");
}

function sendGraphql(body) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "FETCH_QUEUE_JSON", body }, (response) => {
      const err = chrome.runtime.lastError;
      if (err) {
        resolve({ success: false, error: err.message || String(err) });
        return;
      }
      resolve(response);
    });
  });
}

function sendCaseCountsRequest(payload) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "FETCH_CASE_COUNTS", payload }, (response) => {
      const err = chrome.runtime.lastError;
      if (err) {
        resolve({ success: false, error: err.message || String(err) });
        return;
      }
      resolve(response);
    });
  });
}

function sendCompareAndNotify(updates) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "COMPARE_AND_NOTIFY", updates }, (response) => {
      const err = chrome.runtime.lastError;
      if (err) {
        // Notification compare is best-effort; don't fail UI.
        resolve({ success: false, error: err.message || String(err) });
        return;
      }
      resolve(response);
    });
  });
}

function buildLiveOpenCasesGraphqlBody() {
  // Mirroring the list page from the attached HAR (list-id 0e66..., tiny-id QF9k...)
  return {
    operationName: "nowRecordListConnected_min",
    variables: {
      table: "sn_customerservice_case",
      view: "",
      columns: "number,priority",
      fixedQuery: "",
      query:
        "active=true^assigned_toISEMPTY^stateNOT IN6,3,7^assignment_group.nameSTARTSWITHManu - Manu^ORassignment_group.nameSTARTSWITHManu - Pro",
      limit: 100,
      offset: 0,
      queryCategory: "list",
      maxColumns: 50,
      listId: "0e66f167977c4a184b77ff21f053af1f",
      listTitle: "Cases%20%20Manu%20-%20Manu%20All",
      runHighlightedValuesQuery: false,
      menuSelection: "sys_ux_my_list",
      ignoreTotalRecordCount: false,
      columnPreferenceKey: "",
      tiny: "QF9ktNm6N38E3iy4FZfKtZ5tr1DvDUPp"
    },
    query: `
      query nowRecordListConnected_min(
        $columns:String
        $listId:String
        $maxColumns:Int
        $limit:Int
        $offset:Int
        $query:String
        $fixedQuery:String
        $table:String!
        $view:String
        $runHighlightedValuesQuery:Boolean!
        $tiny:String
        $queryCategory:String
        $listTitle:String
        $menuSelection:String
        $ignoreTotalRecordCount:Boolean
        $columnPreferenceKey:String
      ){
        GlideListLayout_Query{
          getListLayout(
            columns:$columns
            listId:$listId
            maxColumns:$maxColumns
            limit:$limit
            offset:$offset
            query:$query
            fixedQuery:$fixedQuery
            table:$table
            view:$view
            runHighlightedValuesQuery:$runHighlightedValuesQuery
            tiny:$tiny
            queryCategory:$queryCategory
            listTitle:$listTitle
            menuSelection:$menuSelection
            ignoreTotalRecordCount:$ignoreTotalRecordCount
            columnPreferenceKey:$columnPreferenceKey
          ){
            layoutQuery{
              count
              queryRows{
                ... on GlideListLayout_QueryRowType{
                  rowData{columnName columnData{displayValue value}}
                }
              }
            }
          }
        }
      }
    `
  };
}

function buildAgentListGraphqlBody(listId, tinyId, limit = 100, offset = 0) {
  // For agent list links (list-id + tiny-id). The reliable way to resolve these
  // (including task queues) is via getTinyListLayout, which returns the list's
  // underlying table/query/rows.
  return {
    operationName: "nowRecordListConnected_min",
    variables: {
      tiny: tinyId,
      runHighlightedValuesQuery: false,
      limit,
      offset,
      ignoreTotalRecordCount: false,
      columnPreferenceKey: ""
    },
    query: `
      query nowRecordListConnected_min(
        $tiny:String!
        $runHighlightedValuesQuery:Boolean!
        $limit:Int
        $offset:Int
        $ignoreTotalRecordCount:Boolean
        $columnPreferenceKey:String
      ){
        GlideListLayout_Query{
          getTinyListLayout(
            tiny:$tiny
            runHighlightedValuesQuery:$runHighlightedValuesQuery
            limit:$limit
            offset:$offset
            ignoreTotalRecordCount:$ignoreTotalRecordCount
            columnPreferenceKey:$columnPreferenceKey
          ){
            tableLabel
            table
            listTitle
            listId
            layoutQuery{
              table
              count
              queryRows{
                ... on GlideListLayout_QueryRowType{
                  rowData{columnName columnData{displayValue value}}
                }
              }
            }
          }
        }
      }
    `
  };
}

async function fetchAgentListCounts(listId, tinyId) {
  const counts = {
    "1 - Critical": 0,
    "2 - High": 0,
    "3 - Moderate": 0,
    "4 - Low": 0,
    "5 - Planning": 0,
    Unknown: 0
  };

  const pageSize = 100;
  const maxRecords = 20000;
  let offset = 0;
  let total = null;
  let tableLabel = "";
  let table = "";

  while (offset < maxRecords) {
    // eslint-disable-next-line no-await-in-loop
    const resp = await sendGraphql(buildAgentListGraphqlBody(listId, tinyId, pageSize, offset));
    if (!resp) {
      return { success: false, error: "No response from background (agent list)" };
    }
    if (resp.needsLogin) {
      return { success: false, needsLogin: true, http: resp.http };
    }
    if (!resp.success) {
      return { success: false, error: "Agent list GraphQL failed", data: resp };
    }

    const data = resp.data;
    const tinyLayout = data?.data?.GlideListLayout_Query?.getTinyListLayout;
    const layout = tinyLayout?.layoutQuery;
    const rows = layout?.queryRows || [];
    if (typeof layout?.count === "number") total = layout.count;

    // keep metadata from the first successful page
    if (!tableLabel && typeof tinyLayout?.tableLabel === "string") tableLabel = tinyLayout.tableLabel.trim();
    if (!table) {
      if (typeof layout?.table === "string") table = layout.table.trim();
      else if (typeof tinyLayout?.table === "string") table = tinyLayout.table.trim();
    }

    // Count this page
    const pageCounts = computePriorityCountsFromListLayout(data);
    for (const k of Object.keys(counts)) {
      counts[k] += pageCounts[k] || 0;
    }

    if (!Array.isArray(rows) || rows.length === 0) break;
    if (rows.length < pageSize) break;
    offset += pageSize;
    if (typeof total === "number" && offset >= total) break;
  }

  const truncated = offset >= maxRecords;
  return { success: true, counts, truncated, tableLabel, table };
}

function buildGraphqlBody(condition) {
  return {
    operationName: "snPolarisLayout",
    variables: {
      condition
    },
    query: `
      query snPolarisLayout($condition:String){
        GlideRecord_Query{
          ui_notification_inbox(
            queryConditions:$condition
          ){
            _rowCount
            _results{
              sys_id{value}
              sys_created_on{value}
              payload{value}
              target_table{value}
              target{value}
              status{value}
              route{value}
              triggered_by{value displayValue}
            }
          }
        }
      }
    `
  };
}

async function run(options = {}) {
  const { skipNotify = false } = options || {};
  const shouldMarkLoading = true;
  setMainLoading(true);

  try {
    currentConfig = await loadConfig();

    // Always render setup list if we are on setup view.
    renderSavedLinks(currentConfig);

    // If user has no links yet, start on setup.
    if (!currentConfig.links.length) {
      showSetup();
      return;
    }

    showMain();

  const items = [];
  const notifyUpdates = [];
  const expandState = await loadExpandState();

  const orderedLinks = [...(currentConfig.links || [])];

  for (const entry of orderedLinks) {
    const parsed = parseServiceNowListLink(entry.link);
    if (parsed.error) {
      const opts = { queueId: entry.id, expanded: false };
      const html = formatCountsBlock(entry.topic, makeZeroCounts(), entry.link, null, opts);
      items.push({ total: 0, html, errorHtml: parsed.error });
      continue;
    }

    let customCounts;
    let displayTopic = entry.topic;
    let typeBadge = null;
    if (parsed.listId && parsed.tinyId) {
      // eslint-disable-next-line no-await-in-loop
      const agentCountsResp = await fetchAgentListCounts(parsed.listId, parsed.tinyId);
      if (agentCountsResp.needsLogin) {
        showLoginRequiredMessage();
        return;
      }
      if (!agentCountsResp.success) {
        const opts = { queueId: entry.id, expanded: false };
        const html = formatCountsBlock(entry.topic, makeZeroCounts(), entry.link, null, opts);
        items.push({ total: 0, html, errorHtml: JSON.stringify(agentCountsResp, null, 2) });
        continue;
      }
      customCounts = agentCountsResp.counts;
      typeBadge = getTypeBadgeForTable(agentCountsResp.table);
    } else {
      // eslint-disable-next-line no-await-in-loop
      const countsResponse = await sendCaseCountsRequest({ table: parsed.table, query: parsed.query });
      if (!countsResponse) {
        blocks.push("", `<strong>${entry.topic} - 0</strong>`, "No response from background (record counts)");
        continue;
      }

      if (countsResponse.needsLogin) {
        showLoginRequiredMessage();
        return;
      }

      if (!countsResponse.success) {
        const opts = { queueId: entry.id, expanded: false };
        const html = formatCountsBlock(entry.topic, makeZeroCounts(), entry.link, null, opts);
        items.push({ total: 0, html, errorHtml: JSON.stringify(countsResponse, null, 2) });
        continue;
      }

      customCounts = countsResponse.counts;
      typeBadge = getTypeBadgeForTable(parsed.table);
    }

    const total = computeTotalCount(customCounts);
    const savedExpanded = expandState?.[entry.id];
    const expanded = total > 0 ? true : (typeof savedExpanded === "boolean" ? savedExpanded : false);
    const opts = { queueId: entry.id, expanded };
    const html = formatCountsBlock(displayTopic, customCounts, entry.link, typeBadge, opts);
    items.push({ total, html });
    if (entry.notifyEnabled !== false) {
      notifyUpdates.push({ queueId: entry.id, topic: displayTopic, counts: customCounts });
    }
  }

  const active = items.filter((it) => it.total > 0);
  const inactive = items.filter((it) => !(it.total > 0));
  const ordered = [...active, ...inactive];
  output.innerHTML = ordered.map((it) => (it.html + (it.errorHtml ? String(it.errorHtml) : ""))).join("");

  // Toggle handler: remember last expanded state
  output?.addEventListener?.("click", async (e) => {
    const btn = e?.target?.closest?.(".toggleButton");
    if (!btn) return;
    const container = btn.closest(".queueBlock");
    if (!container) return;
    const qid = container.getAttribute("data-qid") || "";
    if (!qid) return;
    const isCollapsed = container.classList.contains("collapsed");
    const nextExpanded = isCollapsed;
    container.classList.toggle("collapsed", !nextExpanded);
    container.classList.toggle("expanded", nextExpanded);
    const caretLabel = nextExpanded ? "Collapse" : "Expand";
    const icon = btn.querySelector(".toggleIcon");
    if (icon) {
      icon.classList.toggle("isExpanded", nextExpanded);
    }
    btn.setAttribute("aria-expanded", nextExpanded ? "true" : "false");
    btn.setAttribute("aria-label", `${caretLabel} queue`);
    btn.setAttribute("title", caretLabel);
    const current = await loadExpandState();
    const next = { ...(current || {}) };
    next[qid] = nextExpanded;
    scheduleSaveExpandState(next);
  });

  // Notify in the background (storage.local) without affecting UI.
  // First-run is handled in the service worker: it stores but does not notify.
  // When the popup is kept open, we auto-refresh counts but avoid driving notifications from the popup.
    if (!skipNotify) {
      sendCompareAndNotify(notifyUpdates);
    }

    if (shouldMarkLoading) setMainLoading(false);
  } finally {
    // If we returned early, ensure we don't leave the UI looking "stuck".
    if (shouldMarkLoading && output?.classList?.contains?.("isLoading")) {
      output.classList.remove("isLoading");
    }
  }
}

saveButton?.addEventListener("click", async () => {
  currentConfig = currentConfig || (await loadConfig());

  const topic = String(topicInput?.value || "").trim();
  const link = String(linkInput?.value || "").trim();

  const parsed = parseServiceNowListLink(link);
  if (parsed.error) {
    showSetup(parsed.error);
    renderSavedLinks(currentConfig);
    return;
  }
  if (!topic) {
    showSetup("Please enter a topic.");
    renderSavedLinks(currentConfig);
    return;
  }

  const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const next = {
    links: [...(currentConfig.links || []), { id, topic, link, notifyEnabled: true }]
  };

  await saveConfig(next);
  if (topicInput) topicInput.value = "";
  if (linkInput) linkInput.value = "";
  showSetup();
  renderSavedLinks(next);
});

doneButton?.addEventListener("click", async () => {
  currentConfig = currentConfig || (await loadConfig());
  if (!currentConfig.links.length) {
    showSetup("Add at least one link, or click Change later.");
    renderSavedLinks(currentConfig);
    return;
  }
  // Ensure any pending order changes are saved before switching views
  await flushScheduledConfig();
  showMain();
  run();
});

changeButton?.addEventListener("click", async () => {
  currentConfig = await loadConfig();
  showSetup();
  renderSavedLinks(currentConfig);
});

openSettingsFromMainButton?.addEventListener("click", () => {
  showSettings();
  showSettingsError("");
  loadRefreshRateToUI();
  loadNotifySettingsToUI();
});

refreshMinutesInput?.addEventListener("input", () => {
  normalizeRefreshRateInputs();
  updateRefreshRateDirtyUI();
});

refreshSecondsInput?.addEventListener("input", () => {
  normalizeRefreshRateInputs();
  updateRefreshRateDirtyUI();
});

backFromSettingsButton?.addEventListener("click", async () => {
  if (lastNonSettingsView === "setup") {
    currentConfig = currentConfig || (await loadConfig());
    showSetup();
    renderSavedLinks(currentConfig);
    return;
  }

  // Ensure pending saves are flushed when returning to main
  await flushScheduledConfig();
  await flushScheduledExpandState();
  showMain();
  run();
});

saveRefreshRateButton?.addEventListener("click", async () => {
  showSettingsError("");

  const { totalSeconds, clampedSeconds } = getRefreshRateSecondsFromInputs();

  if (!totalSeconds) {
    showSettingsError("Please set a refresh rate greater than 0 seconds.");
    return;
  }

  await storageSet({ [REFRESH_RATE_KEY]: clampedSeconds });

  const resp = await sendSetRefreshRate(clampedSeconds);
  if (!resp?.success) {
    showSettingsError(resp?.error || "Failed to apply refresh rate.");
    return;
  }

  refreshRateLoadedSeconds = clampedSeconds;
  updateRefreshRateDirtyUI();
  setRefreshRateStatus("Saved");
  setTimeout(() => {
    // Only clear if nothing changed since save.
    const { clampedSeconds: currentClamped } = getRefreshRateSecondsFromInputs();
    if (typeof refreshRateLoadedSeconds === "number" && currentClamped === refreshRateLoadedSeconds) {
      setRefreshRateStatus("");
    }
  }, 1200);
});

const notifyInputs = [
  notifyAllEnabledInput,
  notifyP1Input,
  notifyP2Input,
  notifyP3Input,
  notifyP4Input,
  notifyP5Input,
  notifyUnknownInput
].filter(Boolean);

for (const el of notifyInputs) {
  el.addEventListener("change", () => {
    updateNotifySettingsDirtyUI();
  });
}

saveNotifySettingsButton?.addEventListener("click", async () => {
  showSettingsError("");
  const s = normalizeNotifySettings(readNotifySettingsFromUI());
  await storageSet({ [NOTIFY_SETTINGS_KEY]: s });
  notifySettingsLoaded = s;
  updateNotifySettingsDirtyUI();
  setNotifySettingsStatus("Saved");
  setTimeout(() => {
    const current = normalizeNotifySettings(readNotifySettingsFromUI());
    const loaded = notifySettingsLoaded ? normalizeNotifySettings(notifySettingsLoaded) : null;
    if (loaded && JSON.stringify(loaded) === JSON.stringify(current)) setNotifySettingsStatus("");
  }, 1200);
});

savedLinksEl?.addEventListener("click", async (e) => {
  const btn = e?.target;

  const toggleId = btn?.getAttribute?.("data-toggle-notify");
  if (toggleId) {
    currentConfig = currentConfig || (await loadConfig());
    const next = {
      links: (currentConfig.links || []).map((l) =>
        l.id === toggleId ? { ...l, notifyEnabled: l.notifyEnabled === false } : l
      )
    };
    scheduleSaveConfig(next);
    showSetup();
    renderSavedLinks(next);
    return;
  }

  const moveUpId = btn?.getAttribute?.("data-move-up");
  if (moveUpId) {
    // Persist order
    currentConfig = currentConfig || (await loadConfig());
    const arr = [...(currentConfig.links || [])];
    const idx = arr.findIndex((l) => l.id === moveUpId);
    if (idx > 0) {
      const tmp = arr[idx - 1];
      arr[idx - 1] = arr[idx];
      arr[idx] = tmp;
      await saveConfig({ links: arr });

      // Reorder DOM in place to keep pointer on the same element
      const row = btn.closest('.savedLinkRow');
      const parent = row?.parentElement;
      const prev = row?.previousElementSibling;
      if (row && parent && prev) {
        const prevTop = row.offsetTop;
        parent.insertBefore(row, prev);
        const newTop = row.offsetTop;
        const container = savedLinksEl || parent;
        if (container && typeof container.scrollTop === 'number') {
          container.scrollTop += (newTop - prevTop);
        }
        btn.focus();
      }
    }
    return;
  }

  const moveDownId = btn?.getAttribute?.("data-move-down");
  if (moveDownId) {
    // Persist order
    currentConfig = currentConfig || (await loadConfig());
    const arr = [...(currentConfig.links || [])];
    const idx = arr.findIndex((l) => l.id === moveDownId);
    if (idx !== -1 && idx < arr.length - 1) {
      const tmp = arr[idx + 1];
      arr[idx + 1] = arr[idx];
      arr[idx] = tmp;
      await saveConfig({ links: arr });

      // Reorder DOM in place to keep pointer on the same element
      const row = btn.closest('.savedLinkRow');
      const parent = row?.parentElement;
      const nextRow = row?.nextElementSibling;
      if (row && parent && nextRow) {
        const prevTop = row.offsetTop;
        parent.insertBefore(row, nextRow.nextSibling);
        const newTop = row.offsetTop;
        const container = savedLinksEl || parent;
        if (container && typeof container.scrollTop === 'number') {
          container.scrollTop += (newTop - prevTop);
        }
        btn.focus();
      }
    }
    return;
  }

  const id = btn?.getAttribute?.("data-remove");
  if (!id) return;
  currentConfig = currentConfig || (await loadConfig());
  const next = { links: (currentConfig.links || []).filter((l) => l.id !== id) };
  scheduleSaveConfig(next);
  showSetup();
  renderSavedLinks(next);
});

removeAllButton?.addEventListener("click", async () => {
  currentConfig = currentConfig || (await loadConfig());
  const next = { links: [] };
  scheduleSaveConfig(next);
  showSetup();
  renderSavedLinks(next);
});

// Flush pending saves when popup becomes hidden or is about to unload
document.addEventListener("visibilitychange", async () => {
  if (document.hidden) {
    await flushScheduledConfig();
    await flushScheduledExpandState();
  }
});

window.addEventListener("beforeunload", async () => {
  await flushScheduledConfig();
  await flushScheduledExpandState();
});

run();

// -----------------------------
// Automatic refresh + notifications
// -----------------------------

// Stores previous counts per queueId in chrome.storage.local.
const PREV_COUNTS_KEY = "prevCountsV1";
// Stores user-configured queues in chrome.storage.sync.
const QUEUE_CONFIG_KEY = "queueConfigV1";
// Stores refresh rate in seconds in chrome.storage.sync.
const REFRESH_RATE_KEY = "refreshRateV1";
// Stores notification settings (global enable + per-priority filters) in chrome.storage.sync.
const NOTIFY_SETTINGS_KEY = "notifySettingsV1";
// Alarm name for periodic refresh.
const REFRESH_ALARM_NAME = "ifsQueueRefreshV1";
// Default: 30 seconds. Note browsers may clamp alarms to >= ~1 minute.
const DEFAULT_REFRESH_SECONDS = 30;

const PRIORITY_BUCKETS = [
  "1 - Critical",
  "2 - High",
  "3 - Moderate",
  "4 - Low",
  "5 - Planning"
];

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

function loadNotifySettings() {
  return new Promise((resolve) => {
    chrome.storage.sync.get([NOTIFY_SETTINGS_KEY], (res) => {
      resolve(normalizeNotifySettings(res?.[NOTIFY_SETTINGS_KEY]));
    });
  });
}

function parseJsonSafe(text) {
  try {
    return text ? JSON.parse(text) : null;
  } catch {
    return { raw: text };
  }
}

function looksUnauthenticated(res, text) {
  const looksLikeHtml = /<html[\s>]/i.test(text);
  const redirectedToLogin = res.redirected && /login\.|login\.do|auth|sso/i.test(res.url);
  const loginHtml = looksLikeHtml && /login\.|login\.do|sso|sign in/i.test(text);
  return res.status === 401 || res.status === 403 || redirectedToLogin || loginHtml;
}

function describeFetchFailure(err, url) {
  const msg = err?.name === "AbortError" ? "Request timed out" : (err?.message || String(err));
  return [
    `Network error while fetching ${url}: ${msg}`,
    "If this persists: ensure you are logged in to support.ifs.com in a normal tab, and the extension has Site access to https://support.ifs.com/."
  ].join(" ");
}

async function fetchTextWithTimeout(url, options, timeoutMs = 25000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url, { ...(options || {}), signal: controller.signal });
    const text = await res.text();
    return { ok: true, res, text };
  } catch (err) {
    return { ok: false, error: err };
  } finally {
    clearTimeout(timer);
  }
}

function normalizeCounts(counts) {
  const out = {};
  for (const bucket of PRIORITY_BUCKETS) out[bucket] = 0;
  out.Unknown = 0;

  if (!counts || typeof counts !== "object") return out;
  for (const [k, v] of Object.entries(counts)) {
    const n = typeof v === "number" ? v : Number.parseInt(String(v ?? "0"), 10);
    if (!Number.isFinite(n)) continue;
    const key = typeof k === "string" ? k.trim() : k;
    if (Object.prototype.hasOwnProperty.call(out, key)) out[key] = n;
    else out.Unknown += n;
  }
  return out;
}

function buildIncreaseLines(queueTopic, prev, next, { enabledBuckets } = {}) {
  const lines = [];
  for (const bucket of [...PRIORITY_BUCKETS, "Unknown"]) {
    if (enabledBuckets && !enabledBuckets.has(bucket)) continue;
    const oldVal = Number(prev?.[bucket] ?? 0);
    const newVal = Number(next?.[bucket] ?? 0);
    if (Number.isFinite(oldVal) && Number.isFinite(newVal) && newVal > oldVal) {
      lines.push(`${queueTopic}: ${bucket} ${oldVal} → ${newVal}`);
    }
  }
  return lines;
}

async function createGroupedNotification(lines) {
  if (!lines.length) return;

  const MAX_LINES = 10;
  const shown = lines.slice(0, MAX_LINES);
  const remaining = lines.length - shown.length;
  const message = remaining > 0 ? `${shown.join("\n")}\n(+${remaining} more)` : shown.join("\n");

  chrome.notifications.create({
    type: "basic",
    iconUrl: chrome.runtime.getURL("icon128.png"),
    title: "IFS Queue Monitor",
    message
  });
}

function getPrevCounts() {
  return new Promise((resolve) => {
    chrome.storage.local.get([PREV_COUNTS_KEY], (res) => resolve(res?.[PREV_COUNTS_KEY] || {}));
  });
}

function setPrevCounts(value) {
  return new Promise((resolve) => {
    chrome.storage.local.set({ [PREV_COUNTS_KEY]: value }, () => resolve());
  });
}

async function compareAndNotify(updates) {
  const notifySettings = await loadNotifySettings();
  const enabledBuckets = new Set();
  if (notifySettings?.enabled) {
    for (const bucket of PRIORITY_BUCKETS) {
      if (notifySettings?.priorities?.[bucket] !== false) enabledBuckets.add(bucket);
    }
    if (notifySettings?.priorities?.Unknown !== false) enabledBuckets.add("Unknown");
  }

  const prevMap = await getPrevCounts();
  const nextMap = { ...(prevMap || {}) };
  const allIncreaseLines = [];

  for (const u of updates) {
    const queueId = String(u?.queueId || "").trim();
    const topic = String(u?.topic || "").trim() || "Queue";
    const nextCounts = normalizeCounts(u?.counts);

    if (!queueId) continue;

    const prevEntry = prevMap?.[queueId];
    const prevCounts = prevEntry?.counts ? normalizeCounts(prevEntry.counts) : null;

    // First-run for this queueId: store only, no notifications.
    if (!prevCounts) {
      nextMap[queueId] = { topic, counts: nextCounts, updatedAt: Date.now() };
      continue;
    }

    const lines = buildIncreaseLines(topic, prevCounts, nextCounts, { enabledBuckets });
    if (lines.length) allIncreaseLines.push(...lines);
    nextMap[queueId] = { topic, counts: nextCounts, updatedAt: Date.now() };
  }

  // Update storage first to avoid repeated notifications on rapid refreshes.
  await setPrevCounts(nextMap);

  if (allIncreaseLines.length && notifySettings?.enabled && enabledBuckets.size) {
    await createGroupedNotification(allIncreaseLines);
    return { notified: true, increases: allIncreaseLines.length };
  }
  return { notified: false, increases: 0 };
}

function normalizeQueueConfig(raw) {
  if (!raw || typeof raw !== "object") return { links: [] };
  if (Array.isArray(raw.links)) {
    return {
      links: raw.links
        .filter((x) => x && typeof x === "object")
        .map((x) => ({
          id: String(x.id || "").trim(),
          topic: String(x.topic || "").trim(),
          link: String(x.link || "").trim(),
          notifyEnabled: x.notifyEnabled !== false
        }))
        .filter((x) => x.id && x.topic && x.link)
    };
  }
  // Legacy {topic, link}
  if (raw.topic && raw.link) {
    return {
      links: [
        {
          id: "legacy",
          topic: String(raw.topic).trim(),
          link: String(raw.link).trim(),
          notifyEnabled: true
        }
      ]
    };
  }
  return { links: [] };
}

function loadQueueConfig() {
  return new Promise((resolve) => {
    chrome.storage.sync.get([QUEUE_CONFIG_KEY], (res) => {
      resolve(normalizeQueueConfig(res?.[QUEUE_CONFIG_KEY]));
    });
  });
}

function loadRefreshRateSeconds() {
  return new Promise((resolve) => {
    chrome.storage.sync.get([REFRESH_RATE_KEY], (res) => {
      const raw = res?.[REFRESH_RATE_KEY];
      const n = typeof raw === "number" ? raw : Number.parseInt(String(raw ?? ""), 10);
      if (!Number.isFinite(n) || n <= 0) {
        resolve(DEFAULT_REFRESH_SECONDS);
        return;
      }
      resolve(Math.max(5, Math.min(3600, n)));
    });
  });
}

function normalizeHostname(hostname) {
  return String(hostname || "").trim().toLowerCase();
}

function tryParseUrl(raw) {
  try {
    return new URL(raw);
  } catch {
    return null;
  }
}

function extractListIdTinyFromAgentListPath(pathname) {
  const p = String(pathname || "");
  const m = /\/now\/cwf\/agent\/list\/params\/list-id\/([^\/]+)\/tiny-id\/([^\/]+)/i.exec(p);
  if (!m) return null;
  return { listId: m[1], tinyId: m[2] };
}

function extractTableFromPath(pathname) {
  const name = String(pathname || "").split("/").pop() || "";
  const m = /^([a-z0-9_]+)_list\.do$/i.exec(name);
  return m ? m[1] : null;
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
  return url.searchParams.get("sysparm_query") || url.searchParams.get("sysparm_fixed_query");
}

function getSysparmQueryFromHash(url) {
  const hash = String(url.hash || "").replace(/^#/, "");
  if (!hash) return null;
  const maybeQueryString = hash.includes("?") ? hash.split("?").slice(1).join("?") : hash;
  const params = new URLSearchParams(maybeQueryString);
  return params.get("sysparm_query") || params.get("sysparm_fixed_query");
}

function parseServiceNowListLink(rawLink) {
  const raw = String(rawLink || "").trim();
  if (!raw) return { error: "empty" };
  const url = tryParseUrl(raw);
  if (!url) return { error: "invalid-url" };
  if (normalizeHostname(url.hostname) !== "support.ifs.com") return { error: "wrong-host" };

  const agentList = extractListIdTinyFromAgentListPath(url.pathname);
  if (agentList?.listId && agentList?.tinyId) return { listId: agentList.listId, tinyId: agentList.tinyId };

  const directTable = extractTableFromPath(url.pathname);
  const directQuery = getSysparmQueryFromUrl(url) || getSysparmQueryFromHash(url);
  if (directTable && directQuery) return { table: directTable, query: directQuery };

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
      if (innerTable && innerQuery) return { table: innerTable, query: innerQuery };
    }
  }

  const apiMatch = /^\/api\/now\/table\/([^\/]+)\/?$/i.exec(url.pathname);
  if (apiMatch) {
    const apiTable = apiMatch[1];
    const apiQuery = getSysparmQueryFromUrl(url) || getSysparmQueryFromHash(url);
    if (apiQuery) return { table: apiTable, query: apiQuery };
  }

  return { error: "unrecognized" };
}

function emptyCounts() {
  return {
    "1 - Critical": 0,
    "2 - High": 0,
    "3 - Moderate": 0,
    "4 - Low": 0,
    "5 - Planning": 0,
    Unknown: 0
  };
}

function computePriorityCountsFromListLayoutResponse(graphqlJson) {
  const counts = emptyCounts();
  const rows =
    graphqlJson?.data?.GlideListLayout_Query?.getTinyListLayout?.layoutQuery?.queryRows ||
    graphqlJson?.data?.GlideListLayout_Query?.getListLayout?.layoutQuery?.queryRows ||
    [];

  for (const row of rows) {
    const rowData = row?.rowData;
    if (!Array.isArray(rowData)) continue;
    const priorityCell = rowData.find((c) => c?.columnName === "priority");
    const raw = priorityCell?.columnData?.displayValue;
    const priorityDisplay = typeof raw === "string" ? raw.trim() : null;
    if (!priorityDisplay) continue;
    if (Object.prototype.hasOwnProperty.call(counts, priorityDisplay)) counts[priorityDisplay] += 1;
    else counts.Unknown += 1;
  }

  return counts;
}

function buildAgentListGraphqlBody(listId, tinyId, limit = 100, offset = 0) {
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

async function fetchGraphql(body) {
  const url = "https://support.ifs.com/api/now/graphql";
  const r = await fetchTextWithTimeout(url, {
    method: "POST",
    credentials: "include",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  if (!r.ok) return { success: false, error: describeFetchFailure(r.error, url) };

  const { res, text } = r;
  const contentType = res.headers.get("content-type") || "";
  if (looksUnauthenticated(res, text)) {
    return { success: false, needsLogin: true, http: { status: res.status, statusText: res.statusText, url: res.url, contentType } };
  }
  const data = parseJsonSafe(text);
  if (!res.ok) return { success: false, error: `HTTP ${res.status} ${res.statusText}`, data };
  if (data && data.errors && data.errors.length) return { success: false, error: "GraphQL returned errors", data };
  return { success: true, data };
}

async function fetchCountsForAgentList(listId, tinyId) {
  const pageSize = 100;
  const maxRecords = 20000;
  let offset = 0;
  let total = null;
  const totals = emptyCounts();

  while (offset < maxRecords) {
    // eslint-disable-next-line no-await-in-loop
    const resp = await fetchGraphql(buildAgentListGraphqlBody(listId, tinyId, pageSize, offset));
    if (!resp.success) return resp;

    const graphqlJson = resp.data;
    const tinyLayout = graphqlJson?.data?.GlideListLayout_Query?.getTinyListLayout;
    const layout = tinyLayout?.layoutQuery;
    const rows = layout?.queryRows || [];
    if (typeof layout?.count === "number") total = layout.count;

    const pageCounts = computePriorityCountsFromListLayoutResponse(graphqlJson);
    for (const k of Object.keys(totals)) totals[k] += pageCounts[k] || 0;

    if (!Array.isArray(rows) || rows.length === 0) break;
    if (rows.length < pageSize) break;
    offset += pageSize;
    if (typeof total === "number" && offset >= total) break;
  }

  return { success: true, counts: totals };
}

async function fetchCountsForTableQuery(table, query) {
  const statsUrl = new URL(`https://support.ifs.com/api/now/stats/${encodeURIComponent(table)}`);
  statsUrl.searchParams.set("sysparm_query", query);
  statsUrl.searchParams.set("sysparm_count", "true");
  statsUrl.searchParams.set("sysparm_group_by", "priority");
  statsUrl.searchParams.set("sysparm_display_value", "true");

  const statsUrlStr = statsUrl.toString();
  const r = await fetchTextWithTimeout(statsUrlStr, {
    method: "GET",
    credentials: "include",
    headers: { Accept: "application/json" }
  });
  if (!r.ok) return { success: false, error: describeFetchFailure(r.error, statsUrlStr) };

  const { res, text } = r;
  const contentType = res.headers.get("content-type") || "";
  if (looksUnauthenticated(res, text)) {
    return { success: false, needsLogin: true, http: { status: res.status, statusText: res.statusText, url: res.url, contentType } };
  }
  const data = parseJsonSafe(text);
  if (!res.ok) return { success: false, error: `HTTP ${res.status} ${res.statusText}`, data };
  if (!/application\/json/i.test(contentType) && /<html[\s>]/i.test(text)) {
    return { success: false, needsLogin: true, http: { status: res.status, statusText: res.statusText, url: res.url, contentType } };
  }

  const counts = emptyCounts();
  const result = data?.result;
  if (Array.isArray(result)) {
    for (const row of result) {
      const maybePriority =
        row?.groupby_fields?.priority ??
        row?.group_by_fields?.priority ??
        row?.priority ??
        row?.group_by?.priority;
      const priorityDisplay =
        typeof maybePriority === "string"
          ? maybePriority.trim()
          : (maybePriority?.display_value || maybePriority?.displayValue || maybePriority?.value || "").trim();
      const maybeCount = row?.stats?.count ?? row?.stats?.COUNT ?? row?.count ?? row?.stats?.total;
      const count = Number.parseInt(String(maybeCount ?? "0"), 10);
      if (priorityDisplay && Number.isFinite(count)) {
        if (Object.prototype.hasOwnProperty.call(counts, priorityDisplay)) counts[priorityDisplay] += count;
        else counts.Unknown += count;
      }
    }
    return { success: true, counts, source: "stats" };
  }

  // Fallback: table paging.
  const tableUrlBase = new URL(`https://support.ifs.com/api/now/table/${encodeURIComponent(table)}`);
  tableUrlBase.searchParams.set("sysparm_query", query);
  tableUrlBase.searchParams.set("sysparm_fields", "priority");
  tableUrlBase.searchParams.set("sysparm_display_value", "true");
  tableUrlBase.searchParams.set("sysparm_exclude_reference_link", "true");

  const limit = 1000;
  const maxRecords = 20000;
  let offset = 0;
  const counts2 = emptyCounts();

  while (offset < maxRecords) {
    const pageUrl = new URL(tableUrlBase);
    pageUrl.searchParams.set("sysparm_limit", String(limit));
    pageUrl.searchParams.set("sysparm_offset", String(offset));

    const pageUrlStr = pageUrl.toString();

    // eslint-disable-next-line no-await-in-loop
    const pageR = await fetchTextWithTimeout(pageUrlStr, {
      method: "GET",
      credentials: "include",
      headers: { Accept: "application/json" }
    });
    if (!pageR.ok) return { success: false, error: describeFetchFailure(pageR.error, pageUrlStr) };

    const pageRes = pageR.res;
    const pageText = pageR.text;
    if (looksUnauthenticated(pageRes, pageText)) {
      const ct = pageRes.headers.get("content-type") || "";
      return { success: false, needsLogin: true, http: { status: pageRes.status, statusText: pageRes.statusText, url: pageRes.url, contentType: ct } };
    }
    const pageData = parseJsonSafe(pageText);
    if (!pageRes.ok) return { success: false, error: `HTTP ${pageRes.status} ${pageRes.statusText}`, data: pageData };
    const records = pageData?.result;
    if (!Array.isArray(records) || records.length === 0) break;
    for (const rec of records) {
      const p = rec?.priority;
      const priorityDisplay = typeof p === "string" ? p.trim() : (p?.display_value || p?.displayValue || p?.value || "").trim();
      if (priorityDisplay && Object.prototype.hasOwnProperty.call(counts2, priorityDisplay)) counts2[priorityDisplay] += 1;
      else counts2.Unknown += 1;
    }
    if (records.length < limit) break;
    offset += limit;
  }

  return { success: true, counts: counts2, source: "table" };
}

async function refreshAndNotifyAllQueues() {
  const cfg = await loadQueueConfig();
  if (!cfg.links.length) return;

  const updates = [];
  for (const entry of cfg.links) {
    if (entry?.notifyEnabled === false) continue;
    const parsed = parseServiceNowListLink(entry.link);
    if (parsed.listId && parsed.tinyId) {
      // eslint-disable-next-line no-await-in-loop
      const resp = await fetchCountsForAgentList(parsed.listId, parsed.tinyId);
      if (resp.success) updates.push({ queueId: entry.id, topic: entry.topic, counts: resp.counts });
      continue;
    }
    if (parsed.table && typeof parsed.query === "string") {
      // eslint-disable-next-line no-await-in-loop
      const resp = await fetchCountsForTableQuery(parsed.table, parsed.query);
      if (resp.success) updates.push({ queueId: entry.id, topic: entry.topic, counts: resp.counts });
    }
  }

  if (updates.length) await compareAndNotify(updates);
}

async function ensureRefreshAlarm() {
  const seconds = await loadRefreshRateSeconds();
  chrome.alarms.create(REFRESH_ALARM_NAME, { periodInMinutes: seconds / 60 });
}

chrome.runtime.onInstalled.addListener(() => {
  ensureRefreshAlarm();
});

chrome.runtime.onStartup?.addListener?.(() => {
  ensureRefreshAlarm();
});

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm?.name !== REFRESH_ALARM_NAME) return;
  (async () => {
    await refreshAndNotifyAllQueues();
  })().catch((err) => console.warn("IFS Queue Monitor refresh failed:", err));
});

// -----------------------------
// Existing message-based API for the popup UI
// -----------------------------

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  const respondNeedsLogin = (res, text) => {
    const contentType = res.headers.get("content-type") || "";
    const looksLikeHtml = /<html[\s>]/i.test(text);
    sendResponse({
      success: false,
      needsLogin: true,
      error: "Not logged in to support.ifs.com (or session expired)",
      http: { status: res.status, statusText: res.statusText, url: res.url, contentType },
      data: looksLikeHtml ? { htmlSnippet: text.slice(0, 500) } : { raw: text.slice(0, 500) }
    });
  };

  if (msg.type === "FETCH_QUEUE_JSON") {
    fetch("https://support.ifs.com/api/now/graphql", {
      method: "POST",
      credentials: "include",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(msg.body)
    })
      .then(async (res) => {
        const text = await res.text();

        const contentType = res.headers.get("content-type") || "";

        if (looksUnauthenticated(res, text)) {
          respondNeedsLogin(res, text);
          return;
        }

        const data = parseJsonSafe(text);

        if (!res.ok) {
          sendResponse({
            success: false,
            error: `HTTP ${res.status} ${res.statusText}`,
            data
          });
          return;
        }

        // Defensive: some instances return 200 with HTML when unauthenticated.
        if (!/application\/json/i.test(contentType) && /<html[\s>]/i.test(text)) {
          sendResponse({
            success: false,
            needsLogin: true,
            error: "Not logged in to support.ifs.com (HTML response)",
            http: { status: res.status, statusText: res.statusText, url: res.url, contentType },
            data: { htmlSnippet: text.slice(0, 500) }
          });
          return;
        }

        if (data && data.errors && data.errors.length) {
          sendResponse({
            success: false,
            error: "GraphQL returned errors",
            data
          });
          return;
        }

        sendResponse({ success: true, data });
      })
      .catch(err => sendResponse({ success: false, error: err.toString() }));

    return true; // keep channel open
  }

  if (msg.type === "COMPARE_AND_NOTIFY") {
    (async () => {
      const updates = Array.isArray(msg?.updates) ? msg.updates : [];
      if (!updates.length) {
        sendResponse({ success: true, notified: false, reason: "no-updates" });
        return;
      }

      const result = await compareAndNotify(updates);
      if (result.notified) sendResponse({ success: true, notified: true, increases: result.increases });
      else sendResponse({ success: true, notified: false, reason: "no-increases" });
    })().catch((err) => sendResponse({ success: false, error: err?.toString?.() || String(err) }));

    return true;
  }

  if (msg.type === "SET_REFRESH_RATE") {
    (async () => {
      const raw = msg?.totalSeconds;
      const n = typeof raw === "number" ? raw : Number.parseInt(String(raw ?? ""), 10);
      if (!Number.isFinite(n) || n <= 0) {
        sendResponse({ success: false, error: "Invalid refresh rate" });
        return;
      }
      const clamped = Math.max(5, Math.min(3600, n));
      await new Promise((resolve) => {
        chrome.storage.sync.set({ [REFRESH_RATE_KEY]: clamped }, () => resolve());
      });
      await ensureRefreshAlarm();
      sendResponse({ success: true, seconds: clamped });
    })();
    return true;
  }

  if (msg.type === "FETCH_CASE_COUNTS") {
    const table = msg?.payload?.table;
    const query = msg?.payload?.query;

    if (typeof table !== "string" || !table.trim() || typeof query !== "string") {
      sendResponse({ success: false, error: "Missing or invalid table/query" });
      return;
    }

    const emptyCounts = () => ({
      "1 - Critical": 0,
      "2 - High": 0,
      "3 - Moderate": 0,
      "4 - Low": 0,
      "5 - Planning": 0,
      Unknown: 0
    });

    const statsUrl = new URL(`https://support.ifs.com/api/now/stats/${encodeURIComponent(table)}`);
    statsUrl.searchParams.set("sysparm_query", query);
    statsUrl.searchParams.set("sysparm_count", "true");
    statsUrl.searchParams.set("sysparm_group_by", "priority");
    statsUrl.searchParams.set("sysparm_display_value", "true");

    fetch(statsUrl.toString(), {
      method: "GET",
      credentials: "include",
      headers: {
        Accept: "application/json"
      }
    })
      .then(async (res) => {
        const text = await res.text();
        const contentType = res.headers.get("content-type") || "";

        if (looksUnauthenticated(res, text)) {
          respondNeedsLogin(res, text);
          return;
        }

        const data = parseJsonSafe(text);
        if (!res.ok) {
          sendResponse({ success: false, error: `HTTP ${res.status} ${res.statusText}`, data });
          return;
        }

        // Some instances may return HTML even with 200.
        if (!/application\/json/i.test(contentType) && /<html[\s>]/i.test(text)) {
          respondNeedsLogin(res, text);
          return;
        }

        const counts = emptyCounts();

        const result = data?.result;
        if (Array.isArray(result)) {
          for (const row of result) {
            // Try a few possible shapes.
            const maybePriority =
              row?.groupby_fields?.priority ??
              row?.group_by_fields?.priority ??
              row?.priority ??
              row?.group_by?.priority;

            const priorityDisplay =
              typeof maybePriority === "string"
                ? maybePriority
                : maybePriority?.display_value || maybePriority?.displayValue || maybePriority?.value;

            const maybeCount = row?.stats?.count ?? row?.stats?.COUNT ?? row?.count ?? row?.stats?.total;
            const count = Number.parseInt(String(maybeCount ?? "0"), 10);

            if (typeof priorityDisplay === "string" && Number.isFinite(count)) {
              if (Object.prototype.hasOwnProperty.call(counts, priorityDisplay)) {
                counts[priorityDisplay] += count;
              } else {
                counts.Unknown += count;
              }
            } else {
              counts.Unknown += 1;
            }
          }

          sendResponse({ success: true, counts, source: "stats" });
          return;
        }

        // Fallback: table API paging (slower but resilient).
        const tableUrlBase = new URL(`https://support.ifs.com/api/now/table/${encodeURIComponent(table)}`);
        tableUrlBase.searchParams.set("sysparm_query", query);
        tableUrlBase.searchParams.set("sysparm_fields", "priority");
        tableUrlBase.searchParams.set("sysparm_display_value", "true");
        tableUrlBase.searchParams.set("sysparm_exclude_reference_link", "true");

        const limit = 1000;
        const maxRecords = 20000;
        let offset = 0;
        const counts2 = emptyCounts();

        while (offset < maxRecords) {
          const pageUrl = new URL(tableUrlBase);
          pageUrl.searchParams.set("sysparm_limit", String(limit));
          pageUrl.searchParams.set("sysparm_offset", String(offset));

          // eslint-disable-next-line no-await-in-loop
          const pageRes = await fetch(pageUrl.toString(), {
            method: "GET",
            credentials: "include",
            headers: { Accept: "application/json" }
          });

          // eslint-disable-next-line no-await-in-loop
          const pageText = await pageRes.text();
          if (looksUnauthenticated(pageRes, pageText)) {
            respondNeedsLogin(pageRes, pageText);
            return;
          }

          const pageData = parseJsonSafe(pageText);
          if (!pageRes.ok) {
            sendResponse({ success: false, error: `HTTP ${pageRes.status} ${pageRes.statusText}`, data: pageData });
            return;
          }

          const records = pageData?.result;
          if (!Array.isArray(records) || records.length === 0) break;

          for (const rec of records) {
            const p = rec?.priority;
            const priorityDisplay = typeof p === "string" ? p : p?.display_value || p?.displayValue || p?.value;
            if (typeof priorityDisplay === "string" && Object.prototype.hasOwnProperty.call(counts2, priorityDisplay)) {
              counts2[priorityDisplay] += 1;
            } else {
              counts2.Unknown += 1;
            }
          }

          if (records.length < limit) break;
          offset += limit;
        }

        const truncated = offset >= maxRecords;
        sendResponse({ success: true, counts: counts2, source: "table", truncated });
      })
      .catch((err) => sendResponse({ success: false, error: err.toString() }));

    return true;
  }
});

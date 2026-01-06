const output = document.getElementById("output");
const setupView = document.getElementById("setup");
const mainView = document.getElementById("main");
const topicInput = document.getElementById("topic");
const linkInput = document.getElementById("link");
const saveButton = document.getElementById("save");
const doneButton = document.getElementById("done");
const changeButton = document.getElementById("change");
const setupError = document.getElementById("setupError");
const savedLinksEl = document.getElementById("savedLinks");
const removeAllButton = document.getElementById("removeAll");

const PRIORITY_BUCKETS = [
  "1 - Critical",
  "2 - High",
  "3 - Moderate",
  "4 - Low",
  "5 - Planning"
];

const STORAGE_KEY = "queueConfigV1";

let currentConfig = null;

function showSetup(errorText) {
  setupView.hidden = false;
  mainView.hidden = true;
  if (typeof errorText === "string" && errorText.trim()) {
    setupError.hidden = false;
    setupError.textContent = errorText;
  } else {
    setupError.hidden = true;
    setupError.textContent = "";
  }
}

function showMain(topic) {
  setupView.hidden = true;
  mainView.hidden = false;
}

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
          link: String(x.link || "").trim()
        }))
        .filter((x) => x.topic && x.link)
    };
  }

  // Legacy format: { topic, link }
  if (raw.topic && raw.link) {
    return {
      links: [{ id: "legacy", topic: String(raw.topic).trim(), link: String(raw.link).trim() }]
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
      (l) => `
        <div class="savedLinkRow" data-id="${esc(l.id)}">
          <div class="savedLinkText">
            <div class="savedLinkTopic">${esc(l.topic)}</div>
            <div class="savedLinkUrl">${esc(l.link)}</div>
          </div>
          <button class="removeButton" type="button" data-remove="${esc(l.id)}">Remove</button>
        </div>
      `
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

function formatCountsBlock(title, counts) {
  const lines = [`<strong>${title} - ${computeTotalCount(counts)}</strong>`];
  for (const bucket of PRIORITY_BUCKETS) {
    lines.push(`${bucket} : ${counts[bucket]}`);
  }
  if (counts.Unknown) lines.push(`Unknown : ${counts.Unknown}`);
  return lines.join("\n");
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

async function run() {
  currentConfig = await loadConfig();

  // Always render setup list if we are on setup view.
  renderSavedLinks(currentConfig);

  // If user has no links yet, start on setup.
  if (!currentConfig.links.length) {
    showSetup();
    return;
  }

  showMain();
  output.textContent = "Loading...";

  const blocks = [];
  const notifyUpdates = [];

  for (const entry of currentConfig.links) {
    const parsed = parseServiceNowListLink(entry.link);
    if (parsed.error) {
      blocks.push("", `<strong>${entry.topic} - 0</strong>`, parsed.error);
      continue;
    }

    let customCounts;
    let displayTopic = entry.topic;
    if (parsed.listId && parsed.tinyId) {
      // eslint-disable-next-line no-await-in-loop
      const agentCountsResp = await fetchAgentListCounts(parsed.listId, parsed.tinyId);
      if (agentCountsResp.needsLogin) {
        output.textContent = JSON.stringify(
          {
            error: "Please log in to https://support.ifs.com",
            hint: "Open support.ifs.com in a tab, sign in, then reopen this popup.",
            details: agentCountsResp.http || undefined
          },
          null,
          2
        );
        return;
      }
      if (!agentCountsResp.success) {
        blocks.push("", `<strong>${entry.topic} - 0</strong>`, JSON.stringify(agentCountsResp, null, 2));
        continue;
      }
      customCounts = agentCountsResp.counts;
      if (agentCountsResp.tableLabel) {
        displayTopic = `${entry.topic} (${agentCountsResp.tableLabel})`;
      }
    } else {
      // eslint-disable-next-line no-await-in-loop
      const countsResponse = await sendCaseCountsRequest({ table: parsed.table, query: parsed.query });
      if (!countsResponse) {
        blocks.push("", `<strong>${entry.topic} - 0</strong>`, "No response from background (record counts)");
        continue;
      }

      if (countsResponse.needsLogin) {
        output.textContent = JSON.stringify(
          {
            error: "Please log in to https://support.ifs.com",
            hint: "Open support.ifs.com in a tab, sign in, then reopen this popup.",
            details: countsResponse.http || undefined
          },
          null,
          2
        );
        return;
      }

      if (!countsResponse.success) {
        blocks.push("", `<strong>${entry.topic} - 0</strong>`, JSON.stringify(countsResponse, null, 2));
        continue;
      }

      customCounts = countsResponse.counts;
    }

    blocks.push("", formatCountsBlock(displayTopic, customCounts));
    notifyUpdates.push({ queueId: entry.id, topic: displayTopic, counts: customCounts });
  }

  output.innerHTML = blocks.join("\n");

  // Notify in the background (storage.local) without affecting UI.
  // First-run is handled in the service worker: it stores but does not notify.
  sendCompareAndNotify(notifyUpdates);
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
    links: [...(currentConfig.links || []), { id, topic, link }]
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
  showMain();
  output.textContent = "Loading...";
  run();
});

changeButton?.addEventListener("click", async () => {
  currentConfig = await loadConfig();
  showSetup();
  renderSavedLinks(currentConfig);
});

savedLinksEl?.addEventListener("click", async (e) => {
  const btn = e?.target;
  const id = btn?.getAttribute?.("data-remove");
  if (!id) return;
  currentConfig = currentConfig || (await loadConfig());
  const next = { links: (currentConfig.links || []).filter((l) => l.id !== id) };
  await saveConfig(next);
  showSetup();
  renderSavedLinks(next);
});

removeAllButton?.addEventListener("click", async () => {
  currentConfig = currentConfig || (await loadConfig());
  const next = { links: [] };
  await saveConfig(next);
  showSetup();
  renderSavedLinks(next);
});

run();

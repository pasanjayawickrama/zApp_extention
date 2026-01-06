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

  const parseJsonSafe = (text) => {
    try {
      return text ? JSON.parse(text) : null;
    } catch {
      return { raw: text };
    }
  };

  const looksUnauthenticated = (res, text) => {
    const looksLikeHtml = /<html[\s>]/i.test(text);
    const redirectedToLogin = res.redirected && /login\.|login\.do|auth|sso/i.test(res.url);
    const loginHtml = looksLikeHtml && /login\.|login\.do|sso|sign in/i.test(text);
    return res.status === 401 || res.status === 403 || redirectedToLogin || loginHtml;
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

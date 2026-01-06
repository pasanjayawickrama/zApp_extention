chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
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
        const looksLikeHtml = /<html[\s>]/i.test(text);
        const redirectedToLogin = res.redirected && /login\.|login\.do|auth|sso/i.test(res.url);
        const loginHtml = looksLikeHtml && /login\.|login\.do|sso|sign in/i.test(text);

        if (res.status === 401 || res.status === 403 || redirectedToLogin || loginHtml) {
          sendResponse({
            success: false,
            needsLogin: true,
            error: "Not logged in to support.ifs.com (or session expired)",
            http: { status: res.status, statusText: res.statusText, url: res.url, contentType },
            data: looksLikeHtml ? { htmlSnippet: text.slice(0, 500) } : { raw: text.slice(0, 500) }
          });
          return;
        }

        let data;
        try {
          data = text ? JSON.parse(text) : null;
        } catch {
          data = { raw: text };
        }

        if (!res.ok) {
          sendResponse({
            success: false,
            error: `HTTP ${res.status} ${res.statusText}`,
            data
          });
          return;
        }

        // Defensive: some instances return 200 with HTML when unauthenticated.
        if (!/application\/json/i.test(contentType) && looksLikeHtml) {
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
});

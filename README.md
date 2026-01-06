# IFS Support Fetch Observer (Step 1) — Self-test

This extension is intentionally minimal and has **no UI**, **no storage**, and **does not handle login/credentials**.

## What it captures

- Only runs on: `https://support.ifs.com/*`
- Only observes **page `fetch()`** responses (not XHR, not WebSocket).
- Only forwards responses whose URL path contains `/api/` or `/now/`.
- Only forwards bodies that successfully parse as JSON.
- Uses `response.clone()` so it does **not** consume the page’s original response body.

## Load (Chrome / Edge)

1. Open `chrome://extensions` (or `edge://extensions`).
2. Enable **Developer mode**.
3. Click **Load unpacked** and select this folder.

## Verify it works

1. Open a new tab to `https://support.ifs.com/` (or any path under it).
2. Open DevTools for the extension **service worker**:
   - Go back to `chrome://extensions`
   - Find the extension
   - Click **Service worker** (or **Inspect views** → service worker)
3. In another tab (on `https://support.ifs.com/*`), do something that triggers ServiceNow API calls (for example navigating a page that loads tickets/queues).

When a matching JSON fetch completes, the service worker console should show:

```
===== IFS QUEUE RAW JSON =====
{ ...pretty JSON... }
===== END =====
```

## Notes / limitations

- This step only intercepts **`window.fetch`**. If the site uses `XMLHttpRequest`, those responses will not be captured.
- Some ServiceNow endpoints may return non-JSON or error pages; those are ignored.
- The extension does not modify or block any requests; it only observes and forwards parsed JSON.

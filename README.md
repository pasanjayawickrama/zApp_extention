# IFS Queue Monitor
## Branch Strategy

- **dev**: Active development happens here. All new commits land on `dev`.
- **main**: Receives updates only via merges from `dev` (no direct commits).

### Enforced via pre-push hook

This repo includes a pre-push hook (tracked in `.githooks/pre-push`) that blocks direct pushes to `main` unless the commit is a merge commit (two parents), which is what `--no-ff` merges produce.

To enable the hook locally:

```powershell
git config core.hooksPath .githooks
```

Release to `main`:

```powershell
git checkout main
git pull
git merge --no-ff -m "Release from dev" dev
git push
```

This ensures `main` only advances through merges from `dev`.

This is a Chrome/Edge MV3 extension that shows **record counts by priority** (Cases and Tasks) for queues on `https://support.ifs.com/*`.

It uses your existing login session (it does **not** store credentials). If you're not logged in, the popup will ask you to sign in.

## Load (Chrome / Edge)

1. Open `chrome://extensions` (or `edge://extensions`).
2. Enable **Developer mode**.
3. Click **Load unpacked** and select this folder.

## Verify it works

1. Open a tab to `https://support.ifs.com/` and sign in.
2. Click the extension icon to open the popup.
3. On first use, the popup asks for:
   - **Topic**: any label you want (e.g. the queue name)
   - **Queue link**: paste a ServiceNow list link (either a classic `..._list.do?sysparm_query=...` link or an agent-list `list-id/tiny-id` link)

The popup will then show counts by priority.

### What link should I paste?

Paste a list URL that looks like one of these:

- `https://support.ifs.com/sn_customerservice_case_list.do?sysparm_query=...`
- `https://support.ifs.com/sn_customerservice_task_list.do?sysparm_query=...`
- `https://support.ifs.com/now/cwf/agent/list/params/list-id/<LIST_ID>/tiny-id/<TINY_ID>`
- Classic nav wrapper URLs are also supported as long as they ultimately include a `..._list.do?sysparm_query=...` target.

## Notes / limitations

- The popup stores your topic/link in `chrome.storage.sync` for your browser profile.
- The extension does not modify or block any requests; it only reads counts.

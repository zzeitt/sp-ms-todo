# Microsoft To Do ↔ Super Productivity sync plugin

An [issueProvider](https://github.com/super-productivity/super-productivity) plugin that
**two-way syncs** Microsoft To Do tasks with Super Productivity (SP) projects.

- **Cross-platform**: reads and writes both go over SP's `http` channel (Angular HttpClient →
  XHR), so the plugin runs on desktop, mobile, and web. No `nodeExecution` permission is used,
  so it is **not** rejected on mobile.
- **Desktop writes**: Graph write requests (create / update / delete) can still be blocked in
  the desktop renderer by SP's CORS preflight handling (`status=0 Unknown Error`) — a host
  limitation the plugin can't bypass without `nodeExecution` (which would break mobile).
- **Auth**: OAuth 2.0 device code flow (Entra ID) — no redirect URI to register; refresh
  tokens renew automatically.
- **Sync**: driven by SP's two-way-sync engine; the plugin only declares field mappings.
- **Language**: the config dialog and info panel follow SP's UI language (English / 简体中文).

> ⚠️ On the desktop renderer, Graph **write requests** (create / update / delete) can fail
> with `status=0 Unknown Error` (SP's Electron CORS-preflight handling). The plugin
> deliberately does **not** use `nodeExecution` (which would fix desktop writes but make SP
> reject the plugin on mobile / web), so **mobile works fully**; desktop auto-create is
> affected by this host-level limitation.

> 中文说明见 [README-zh.md](README-zh.md)。

## Requirements

- Super Productivity **18.20+** (desktop, mobile, or web)
- A Microsoft Entra ID app registration (for the `clientId`) — a public client, no redirect
  URI needed

## 1. Register an Entra ID app (get a clientId)

1. Open the [Azure portal → App registrations → New registration](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade).
2. Name it whatever you like (e.g. `sp-ms-todo`); for account type choose "Accounts in any
   organizational directory and personal Microsoft accounts" (personal MSA needs this type for
   To Do).
3. After registering, note the "Application (client) ID".

4. No redirect URI is required — the plugin signs in with the OAuth 2.0 **device code flow**,
   which never redirects back to the app. Leave `replyUrlsWithType` empty.

5. Enable public client flows: set the Manifest's top-level `"allowPublicClient"` to `true`
   (equivalent to "Authentication → Advanced settings → Allow public client flows").

6. Under "API permissions" confirm (and grant) the `Tasks.ReadWrite` delegated permission;
   `offline_access` is requested via the `scope` parameter (no need to add it manually).

## 2. Configure the clientId (either, or both)

The clientId is **not hardcoded** in the source; there are two ways to set it:

**Option A — enter it in the SP UI (recommended)**: the plugin's config dialog has a
**Client ID** input (mapping to `pluginConfig.clientId`). Fill it in before connecting; it
overrides the build-time default.

**Option B — inject at build time as the default**: set `MS_TODO_CLIENT_ID` before building;
esbuild bakes it into `dist/plugin.js` as the default clientId:

```bash
# Windows PowerShell
$env:MS_TODO_CLIENT_ID = 'your Entra app clientId'
# macOS / Linux
export MS_TODO_CLIENT_ID='your Entra app clientId'
```

> The `clientId` is a public client, **not a secret** (RFC 8252); device code flow needs no
> client secret or redirect URI — the user approves the device code at the verification URI.

Then build:

```bash
cd ms-todo-plugin
npm install
npm run build          # produces dist/ (plugin.js + manifest.json + icon.svg + index.html + ms-todo-provider.zip)
npm run typecheck      # optional: type check
```

## 3. Install into Super Productivity

1. Open SP → Settings → Plugins → "Load from disk".
2. Pick the built `dist/ms-todo-provider.zip`. SP **only accepts `.zip`**, and
   `manifest.json` / `plugin.js` / `icon.svg` must be at the **root** of the archive (not
   inside a subfolder).
3. Go to the task view → create/edit an **issue provider**, type "Microsoft To Do".

## 4. Usage

1. In the provider config click "Test connection" (it doubles as the sign-in button).
2. A dialog shows a code — open the printed URL in any browser, enter the code, and approve.
3. Back in SP the "To Do list" dropdown auto-loads; pick a list.
4. That list's tasks flow into the project backlog; once added as local tasks they two-way
   sync with To Do:
   - completion status, title, notes, due date, and due time sync both ways.

> The provider pulls tasks in pages and, by default, stops at **200** tasks per list to
> avoid lag on huge lists. Change the **"Max tasks to load"** dropdown
> (50 / 100 / 200 / 500 / 1000 / All) to raise or remove the cap.

## 5. Multiple lists → multiple projects

One issue provider instance is bound to one SP project (an SP plugin API limitation: imported
tasks always land in the provider's own project, and there's no per-issue project routing).
To import **different To Do lists** into **different SP projects**, create **one provider
instance per list** and put it in the target project:

1. Open the target project → new issue provider, type "Microsoft To Do".
2. Pick that list in the "To Do list" dropdown (sign-in is shared — every instance of the
   plugin stores its token under the same `pluginId`, so you sign in once).
3. Repeat for each (list → project) pair.

> Tokens are keyed by `pluginId` (via `setSecret`), so multiple instances don't need repeat
> logins; but each instance's `listId` is independent — make sure you pick the right list per instance.

## 6. Reverse sync (SP → To Do)

Tasks created in SP can flow back to To Do in two ways:

- **Auto-create new tasks**: because the plugin implements `createIssue`, SP's provider
  config gains an **"Auto-create issues for new tasks"** checkbox (in the two-way-sync
  section). It is disabled until you set a **default project** for the provider. Once
  enabled, every new top-level task you add to that default project is automatically pushed
  to the configured To Do list (and stays linked for two-way sync).
- **Manually link / push an existing task**: add an existing SP task from this provider
  (the issue panel's "add as issue" flow); from then on its completion status, title, notes,
  and due date push to To Do via the sync engine (`updateIssue`).

Both directions use the same field mappings (see below).

### Completion push safety net (0.1.4+)

SP only pushes on task-update actions and does not retry transient failures (network /
desktop CORS `status=0`), and it skips a push when the remote changed since the last
synced baseline. From 0.1.4 the plugin adds a `taskUpdate` hook safety net: 4 s after a
local completion it checks whether SP already wrote
`issueLastSyncedValues.state = completed`; if not, the plugin fetches the issue, PATCHes
`status=completed` itself (idempotent, now retried) and writes the baseline back. Failures
surface in a snackbar and the plugin log, and the **Reconcile** tab repairs them later.

Pulling is batched too (0.1.8+): the plugin keeps an in-process snapshot per list — one
full fetch on cold start (`$top=1000`, usually a single request), then a single
`lastModifiedDateTime gt …` incremental per 15 s. `getByIds` reads only the snapshot and
never probes missing ids one by one. A 60 s poll therefore costs ~1 request per list, which
avoids Graph 429 / `ERR_HTTP2_PROTOCOL_ERROR`. In practice those errors made the host's
`getByIds` catch drop the **whole** provider batch, so a completion made in To Do never
reached SP. Provider instances bound to the same list share one snapshot.

Limitation: MS→SP is still driven by the host's 60 s poll (measured 0–60 s end-to-end).
Tasks moved to an unconfigured list are recognised by the cross-list probe and are never
mistaken for deletions.

### Delete sync (0.1.9+)

**SP → To Do** is host-built-in: deleting a linked task waits out the undo window, then
calls the plugin's `deleteIssue` (Graph DELETE, 404 treated as already deleted, idempotent).
This covers single and bulk deletes. Host caveat: quitting the app inside the undo window
drops the remote delete (an orphaned To Do task is left behind).

**To Do → SP** is new in 0.1.9. Each list snapshot is rebuilt authoritatively every
**5 minutes**, so remotely deleted tasks disappear from it. `getByIds` then confirms a
missing id (at most 5 per cycle): a successful cross-list `getById` means the task was only
moved and is returned normally; a confirmed 404 reports a `state='deleted'` placeholder
issue, and the host deletes the local SP task (asking first when tracked time exists).
429/network errors are never treated as deletions.

Config **"Keep tasks deleted in To Do"** (unchecked by default): check it to keep the local
task and only keep the link; unlink manually from **Link health**.

## Force sync & dead-link detection (the plugin's standalone view)

To manually (re)push a task that failed to auto-create, and to chase down recurring
`404 ErrorItemNotFound` errors, the plugin provides a **standalone view** (`index.html`, i.e.
SP's "column" page) that gathers every action and its **progress output** in one place —
no more "select a task + top-toolbar button".

You can trigger the sync two ways:

- **One-click sync (top toolbar)**: click the **"Sync Microsoft To Do"** button (the
  `check_box` icon) in the top toolbar. It pushes every unsynced task in the current scope
  (the current project, or all projects when not in a project view) to its bound To Do list,
  with a start/end snackbar — no need to open the view. (The header button can't show a
  per-task progress log; for that, use the standalone view below.)
- **Standalone view (full detail + progress)**: open the **"Microsoft To Do"** entry from the
  left menu (desktop) or the left sidenav drawer (mobile). The view groups every action —
  push, link scan, cleanup — with inline progress output.

> The top toolbar used to carry a second, duplicate entry (the side-panel button, which the
> host renders as a colored plugin icon and auto-opens the panel before any `onClick`). It was
> removed because it duplicated the standalone view and rendered as a blank panel on mobile
> (host issue #8394).

> View-entry registration is wrapped in `PluginAPI.onReady`, fired by the host after the
> plugin finishes loading — matching the `Goals Vision Board` plugin (which works on both
> mobile and desktop) so the entries don't get lost on mobile.

> 📱 **Device-local install**: SP stores uploaded plugin **code** (`plugin.js` / `index.html`)
> in each device's own IndexedDB cache — it does **not** sync across devices (only the
> plugin's enabled/disabled flag syncs). So after updating the build you must **install the
> new `ms-todo-provider.zip` on each device again** (desktop *and* mobile). To confirm a
> device is running the current build, open the view: the status bar ends with
> **`插件 v0.x.y`**, and the plugin log (Settings → Plugins) shows
> `[ms-todo 0.x.y] plugin.js executing` / `onReady fired` / `registerHeaderButton ok` /
> `registerMenuEntry ok`. If the version is old or those lines are missing, re-install
> the ZIP on that device.

The view is split into five functional tabs:

- **Push**: lists **unsynced** local tasks (auto-create failures with no `issueId`), each
  showing title / project / tags / notes / due date & time. You can **multi-select** and press
  "Push selected" for a batch push, or push a single row. Pushing POSTs the task (title /
  notes / due date & time / completion status) to the configured To Do list and writes back
  the `issueId` to complete the link — **no need to select an SP task first** (fixing the old
  "please select…" false report).
- **Linked**: the plugin's currently-synced tasks; the `MS id` is truncated with an ellipsis
  and shown in full on hover.
- **Dead links**: click "Scan" to GET each locally-linked task in the current project and list
  the ones that no longer exist in To Do. Delete those local tasks to stop the recurring 404.
- **Reconcile**: scan every linked task for a **completion divergence** (local `isDone` vs the
  last synced baseline `issueLastSyncedValues.state`) and fix each one: remote completed →
  tick locally; local completed → push `completed`; both completed → rewrite the baseline only.
  Each row shows the action taken and any error, so missed pushes and stale baselines can be
  repaired without waiting for the next poll.
- **Cleanup**: browse To Do tasks per list with inline actions — unlink already-linked tasks
  (back to Push), or delete the remote task for unlinked ones (with a confirmation step).
- The status bar shows the **human-readable list name / project name / task count**; when not
  connected or no list is chosen it shows setup guidance; the bottom **log** prints the result
  of every action inline.

> Each tab targets the **current project** (i.e. the **currently configured list** for this
> provider instance). If you sync multiple lists, open the view once per project and run them
> there.

> "Push" / "Scan" still go through `PluginAPI.request` (the manifest declares `http` +
> `allowedHosts`, backed by the same host channel as the issueProvider-injected `http`, with
> no `nodeExecution`), so they work on desktop, mobile, and web. On desktop they are subject
> to the same `status=0` CORS-preflight limitation as auto-create (see Known limitations).

## Field mappings

| SP task | To Do task | Direction |
|---|---|---|
| `isDone` | `status` (`completed` / `notStarted`) | both |
| `title` | `title` | both |
| `notes` | `body.content` (`text`) | both |
| `dueDay` | `dueDateTime` (date-only) | both |
| `dueWithTime` | `dueDateTime` (date + time) | both |

## Known limitations

- **Desktop write requests**: SP's Electron renderer CORS-preflight handling can block Graph
  writes (`status=0 Unknown Error`), so "Auto-create issues for new tasks" and the view's
  "Push" action may fail on desktop. This is a host-level limitation the plugin can't bypass without
  the `nodeExecution` permission — which would make SP reject the plugin on mobile — so mobile
  compatibility is kept instead.
- **Web/PWA sign-in**: the device-code token endpoint has no CORS, so sign-in works on desktop
  and native mobile only (same as every other SP integration — SP shows a cross-origin warning
  on web).
- **Personal Microsoft account (MSA)**: device-code flow support for personal accounts is not
  guaranteed by Microsoft — verify with your account type; work/school accounts are fine.
- **Refresh token expiry**: the Graph refresh token rolls for 90 days. As long as the plugin
  syncs at least once every 90 days it renews automatically; after 90 days idle you sign in again.
- **Timezones**: a To Do reminder that only has a *date* (no time) is stored by Microsoft as
  that day's local midnight, so the plugin distinguishes date-only vs. date+time by whether the
  time is `00:00` — a timed reminder set exactly at midnight is treated as date-only. Times are
  parsed in the reminder's own timezone (`UTC` is parsed as UTC; anything else as the device's
  local wall-clock), and `dueWithTime` is written back to Graph as UTC; if your To Do timezone
  differs from your device timezone, times may shift.
- **Subtasks**: SP subtasks ↔ To Do `checklistItems` is not implemented (a separate collection
  that field mappings can't express).
- **Tags / time tracking**: To Do has no matching fields, so they don't sync.
- **Task deep links**: To Do has no stable public task URL; `getIssueLink` points to the web
  app's home page.

## Development

```bash
npm install
npm run typecheck        # type check src/
npm run build            # bundle → dist/plugin.js + ms-todo-provider.zip
npm test                 # all tests (integration auto-skips without credentials)
npm run test:unit        # pure-logic unit tests only
npm run test:integration # real Graph round-trip (needs MS_TODO_ACCESS_TOKEN + MS_TODO_LIST_ID)
```

### Testing

- **Unit tests** (`test/unit/`) cover config parsing, task↔issue mapping, date
  conversion, the field-mapping value transforms, and pagination — no network needed.
- **Integration test** (`test/integration/`) does a real HTTPS round-trip against
  Microsoft Graph through the plugin's HTTP channel (Node's built-in `fetch`). It
  self-skips unless you provide a token + list id:

  ```bash
  # PowerShell
  $env:MS_TODO_ACCESS_TOKEN='...'; $env:MS_TODO_LIST_ID='...'
  # bash
  export MS_TODO_ACCESS_TOKEN='...' MS_TODO_LIST_ID='...'
  ```

### Releasing

Version is single-sourced in `package.json` (`src/manifest.json` is kept in lockstep
by the bump script). The release flow mirrors zot-tool:

1. Bump the version in a PR:
   ```bash
   npm run bump:patch   # or bump:minor / bump:major, or node scripts/bump-version.mjs 1.2.3
   ```
2. Merge the PR into `main`. The `auto-release` workflow reads `package.json`,
   creates a `vX.Y.Z` tag, and publishes a GitHub Release with the built
   `ms-todo-provider.zip` attached as a download asset.

## Directory layout

```
ms-todo-plugin/
├── .github/workflows/      # CI: test.yml / auto-release.yml / release.yml
├── package.json
├── tsconfig.json
├── vitest.config.mts       # vitest config (node env)
├── scripts/
│   ├── build.js            # esbuild bundles plugin.ts → dist/plugin.js (IIFE)
│   └── bump-version.mjs    # bump package.json + src/manifest.json together
├── src/
│   ├── manifest.json
│   ├── types.ts            # self-contained plugin-api types
│   ├── i18n.ts             # zh/en runtime localization helper
│   ├── graph.ts            # Graph client (PluginHttp, desktop + mobile + web)
│   ├── auth.ts             # device-code flow: sign-in + token refresh (setSecret)
│   ├── domain.ts           # pure logic: parseConfig / mapping / fieldMappings
│   ├── force-sync.ts       # pure logic: force-sync + dead-link detection (task selection)
│   ├── plugin.ts           # registerIssueProvider main logic + view RPC bridge
│   ├── index.html          # standalone view (force sync / dead-link scan / progress log)
│   └── icon.svg
├── test/
│   ├── unit/               # pure-logic unit tests
│   └── integration/        # gated real-Graph integration test
├── README.md               # English
└── README-zh.md            # 中文
```

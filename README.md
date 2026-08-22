# Microsoft To Do ↔ Super Productivity sync plugin

An [issueProvider](https://github.com/super-productivity/super-productivity) plugin that
**two-way syncs** Microsoft To Do tasks with Super Productivity (SP) projects.

- **Desktop + mobile**: the Graph API returns full CORS headers, so the plugin reaches
  `graph.microsoft.com` directly over SP's `http` channel (a plain browser request) on desktop
  (Electron) and native mobile.
- **Auth**: OAuth 2.0 device code flow (Entra ID) — no redirect URI to register; refresh
  tokens renew automatically. Sign-in works on desktop and native mobile.
- **Sync**: driven by SP's two-way-sync engine; the plugin only declares field mappings.
- **Language**: the config dialog and info panel follow SP's UI language (English / 简体中文).

> 中文说明见 [README-zh.md](README-zh.md)。

## Requirements

- Super Productivity **18.20+** (desktop, or native mobile)
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

> The `clientId` is a public client, **not a secret** (RFC 8252); security comes from
> PKCE + the registered redirect URI.

Then build:

```bash
cd ms-todo-plugin
npm install
npm run build          # produces dist/ (plugin.js + manifest.json + icon.svg + ms-todo-provider.zip)
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
   - completion status, title, notes, and due date sync both ways.

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

## Field mappings

| SP task | To Do task | Direction |
|---|---|---|
| `isDone` | `status` (`completed` / `notStarted`) | both |
| `title` | `title` | both |
| `notes` | `body.content` (`text`) | both |
| `dueDay` | `dueDateTime` | both |

## Known limitations

- **Web/PWA sign-in**: the device-code token endpoint has no CORS, so sign-in works on desktop
  and native mobile only (same as every other SP integration — SP shows a cross-origin warning
  on web).
- **Personal Microsoft account (MSA)**: device-code flow support for personal accounts is not
  guaranteed by Microsoft — verify with your account type; work/school accounts are fine.
- **Refresh token expiry**: the Graph refresh token rolls for 90 days. As long as the plugin
  syncs at least once every 90 days it renews automatically; after 90 days idle you sign in again.
- **Timezones**: `dueDay` is written back to `dueDateTime` at UTC midnight; non-UTC users may
  see the date shift in To Do (switch to `dueWithTime` for precision — planned for v2).
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
   creates a `vX.Y.Z` tag, and publishes a GitHub Release.

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
│   ├── plugin.ts           # registerIssueProvider main logic
│   └── icon.svg
├── test/
│   ├── unit/               # pure-logic unit tests
│   └── integration/        # gated real-Graph integration test
├── README.md               # English
└── README-zh.md            # 中文
```

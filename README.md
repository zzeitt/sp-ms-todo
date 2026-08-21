# Microsoft To Do ↔ Super Productivity sync plugin

An [issueProvider](https://github.com/super-productivity/super-productivity) plugin that
**two-way syncs** Microsoft To Do tasks with Super Productivity (SP) projects.

- **Desktop (Electron) only**: the Graph main API has no CORS, so the plugin reaches
  `graph.microsoft.com` through SP's `nodeExecution` (a Node child process); Web/PWA
  won't work.
- **Auth**: SP built-in OAuth (Entra ID, Authorization Code + PKCE); refresh tokens renew
  automatically.
- **Sync**: driven by SP's two-way-sync engine; the plugin only declares field mappings.
- **Language**: the config dialog and info panel follow SP's UI language (English / 简体中文).

> 中文说明见 [README-zh.md](README-zh.md)。

## Requirements

- Super Productivity **18.20+** (desktop)
- A Microsoft Entra ID app registration (for the OAuth `clientId`)

## 1. Register an Entra ID app (get a clientId)

1. Open the [Azure portal → App registrations → New registration](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade).
2. Name it whatever you like (e.g. `sp-ms-todo`); for account type choose "Accounts in this
   organizational directory only" or "Accounts in any organizational directory and personal
   Microsoft accounts" (personal Microsoft accounts need the personal-account type for To Do).
3. After registering, note the "Application (client) ID".

4. Register the redirect URI (**must be done via the Manifest**; the portal forces https for
   anything that isn't `localhost`, so it won't accept `http://127.0.0.1`): open the
   **Manifest**, and add one entry to `replyUrlsWithType` whose `type` is `InstalledClient`
   (**not** `Web`):

   ```json
   {
     "url": "http://127.0.0.1:53682",
     "type": "InstalledClient"
   }
   ```

   SP's Electron loopback requires `127.0.0.1` (not `localhost`) and requires an exact match
   (Microsoft does not treat the port on `127.0.0.1` as interchangeable), so the port is
   fixed at 53682.

5. In the same manifest set the top-level `"allowPublicClient"` to `true` (equivalent to
   "Authentication → Advanced settings → Allow public client flows"). SP uses Authorization
   Code + PKCE without a client secret, so this must be on.

6. Under "API permissions" confirm (and grant) the `Tasks.ReadWrite` delegated permission;
   `offline_access` is requested via the `scope` parameter (no need to add it manually).

## 2. Configure the clientId (either, or both)

The clientId is **not hardcoded** in the source; there are two ways to set it:

**Option A — enter it in the SP UI (recommended)**: the plugin's config dialog has a
**Client ID** input (mapping to `pluginConfig.oauthOverrides.clientId`). Fill it in before
connecting; it overrides the build-time default. If you don't use Option B, this field is
required.

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
3. On first load SP shows a **Node.js execution** consent dialog (the plugin uses it to reach
   Graph from a Node child process) — choose "Allow".
4. Go to the task view → create/edit an **issue provider**, type "Microsoft To Do".

## 4. Usage

1. In the provider config click "Connect Microsoft account" and sign in.
2. Once connected, the "To Do list" dropdown auto-loads; pick a list.
3. That list's tasks flow into the project backlog; once added as local tasks they two-way
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
2. Pick that list in the "To Do list" dropdown (OAuth is a single sign-in — every instance
   shares the same Microsoft account and token).
3. Repeat for each (list → project) pair.

> OAuth tokens are keyed by `pluginId`, so multiple instances don't need repeat logins; but
> each instance's `listId` is independent — make sure you pick the right list per instance.

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

- **Desktop only**: Web/PWA has no `nodeExecution`, so it can't reach Graph.
- **Timezones**: `dueDay` is written back to `dueDateTime` at UTC midnight; non-UTC users may
  see the date shift in To Do (switch to `dueWithTime` for precision — planned for v2).
- **Subtasks**: SP subtasks ↔ To Do `checklistItems` is not implemented (a separate collection
  that field mappings can't express).
- **Tags / time tracking**: To Do has no matching fields, so they don't sync.
- **Task deep links**: To Do has no stable public task URL; `getIssueLink` points to the web
  app's home page.
- **Proxy**: the Node child process env is cleared (no `http_proxy`), so system-level proxies
  won't apply.

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
  Microsoft Graph through a Node child process. It self-skips unless you provide a
  token + list id:

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
│   ├── graph.ts            # executeNodeScript + require('https') Graph client
│   ├── domain.ts           # pure logic: parseConfig / mapping / fieldMappings
│   ├── plugin.ts           # registerIssueProvider main logic
│   └── icon.svg
├── test/
│   ├── unit/               # pure-logic unit tests
│   └── integration/        # gated real-Graph integration test
├── README.md               # English
└── README-zh.md            # 中文
```

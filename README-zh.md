# Microsoft To Do ↔ Super Productivity 同步插件

将 Microsoft To Do 的任务与 Super Productivity（SP）项目**双向同步**的
[issueProvider](https://github.com/super-productivity/super-productivity) 插件。

- **仅桌面端（Electron）**：Graph 主接口无 CORS，插件通过 SP 的 `nodeExecution`
  （Node 子进程）访问 `graph.microsoft.com`；Web/PWA 不可用。
- **认证**：SP 内置 OAuth（Entra ID，Authorization Code + PKCE），刷新令牌自动续期。
- **同步**：由 SP 的 two-way-sync 引擎驱动，插件只声明字段映射。
- **语言**：配置对话框与信息面板跟随 SP 客户端语言（英文 / 简体中文）。

## 需求

- Super Productivity **18.20+**（桌面版）
- 一个 Microsoft Entra ID 应用注册（用于 OAuth 的 `clientId`）

## 1. 注册 Entra ID 应用（获取 clientId）

1. 打开 [Azure 门户 → 应用注册 → 新注册](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade)。
2. 名称随意（如 `sp-ms-todo`）；账户类型选「仅此组织目录中的账户」或「任何组织目录 + 个人 Microsoft 账户」（个人 MSA 用 To Do 需包含个人账户类型）。
3. 注册完成后，记下「应用程序(客户端) ID」。
4. 登记重定向 URI（**必须走清单 Manifest**；门户对非 `localhost` 强制 https，不接受
   `http://127.0.0.1`）：进入「清单(Manifest)」，在 `replyUrlsWithType` 里新增一条，
   `type` 必须是 `InstalledClient`（不能是 `Web`）：
   ```json
   {
     "url": "http://127.0.0.1:53682",
     "type": "InstalledClient"
   }
   ```
   SP 的 Electron loopback 强制 `127.0.0.1`（不能 `localhost`），且要求精确匹配
   （微软不把 `127.0.0.1` 的端口视为可互换），故端口固定为 53682。
5. 同清单里把顶层 `"allowPublicClient"` 设为 `true`（对应「身份验证 → 高级设置 →
   允许公共客户端流」）。SP 走 Authorization Code + PKCE、不带 client secret，必须开启。
6. 「API 权限」确认已含（并授予）`Tasks.ReadWrite` 委托权限；`offline_access` 由
   `scope` 参数请求（无需手动加）。

## 2. 配置 clientId（二选一，或两者都要）

clientId **不硬编码在源码里**，有两种配置方式：

**方式 A —— 在 SP 界面里填写（推荐）**：插件配置对话框里有一个 **Client ID** 输入框
（对应 `pluginConfig.oauthOverrides.clientId`），连接前填上即可；它会覆盖构建时的默认值。
未用方式 B 注入时，这是**必填项**。

**方式 B —— 构建时从环境变量注入（作为默认值）**：构建前设置 `MS_TODO_CLIENT_ID`，
它会被 esbuild 打进 `dist/plugin.js` 作为默认 clientId：

```bash
# Windows PowerShell
$env:MS_TODO_CLIENT_ID = '你的 Entra 应用 clientId'
# macOS / Linux
export MS_TODO_CLIENT_ID='你的 Entra 应用 clientId'
```

> `clientId` 是 public client，**不是机密**（RFC 8252）；真正的安全由 PKCE + 注册的重定向 URI 保证。

然后构建：

```bash
cd ms-todo-plugin
npm install
npm run build          # 产出 dist/（plugin.js + manifest.json + icon.svg + ms-todo-provider.zip）
npm run typecheck      # 可选：类型检查
```

## 3. 安装到 Super Productivity

1. 打开 SP → 设置 → 插件 → 「从本地加载」。
2. 选择构建产出的 `dist/ms-todo-provider.zip`。SP **只接受 `.zip`**，且
   `manifest.json` / `plugin.js` / `icon.svg` 必须在压缩包**根目录**（不能套一层文件夹）。
3. 首次加载会弹出 **Node.js 执行授权**对话框（本插件用它在 Node 子进程里访问
   Graph），选择「Allow」。
4. 进入任务视图 → 新建/编辑一个 **issue provider**，类型选「Microsoft To Do」。

## 4. 使用

1. 在 provider 配置里点击「连接 Microsoft 账户」，完成登录。
2. 连接成功后「To Do 列表」下拉会自动加载，选择一个列表。
3. 该列表的任务会进入项目 backlog；添加为本地任务后即与 To Do 双向同步：
   - 完成状态、标题、备注、截止日期互相同步。

> 插件按分页拉取任务，默认每个列表最多拉 **200** 条，避免超大列表卡顿。可在
> 「**加载任务上限**」下拉里调整（50 / 100 / 200 / 500 / 1000 / 全部）以放宽或取消上限。

## 5. 多个列表 → 多个项目

一个 issue provider 实例只绑定一个 SP 项目（SP 插件 API 限制：导入任务固定落到该
provider 所属项目，无法按 issue 自动跨项目路由）。要把**不同的 To Do 列表**导入到
**不同的 SP 项目**，做法是**为每个列表建一个 provider 实例**并放进目标项目：

1. 打开目标项目 → 新建 issue provider，类型选「Microsoft To Do」。
2. 在「To Do 列表」下拉里选该列表（OAuth 只登录一次，所有实例共享同一 Microsoft
   账户与令牌）。
3. 重复以上两步，为每个 (列表 → 项目) 组合各建一个实例。

> OAuth 令牌按 `pluginId` 共享，多实例无需重复登录；但各实例的 `listId` 相互独立，
> 务必为每个实例选对列表。

## 6. 反向同步（SP → To Do）

在 SP 中创建的任务可以通过两种方式回流到 To Do：

- **自动创建新任务**：因为插件实现了 `createIssue`，provider 配置的「双向同步」区会
  多出一个 **「自动为新任务创建问题」** 复选框；在为该 provider 设置**默认项目**之前它
  处于禁用状态。启用后，在该默认项目里新建的每个顶级任务都会自动推送到所配置的
  To Do 列表，并保持双向同步链接。
- **手动链接 / 推送已有任务**：把已有 SP 任务从该 provider 添加为 issue（信息面板的
  「添加为 issue」流程），之后其完成状态、标题、备注、截止日期即通过同步引擎
  （`updateIssue`）推送到 To Do。

两个方向共用同一套字段映射（见下文）。

## 字段映射

| SP 任务 | To Do 任务 | 方向 |
|---|---|---|
| `isDone` | `status`（`completed` / `notStarted`） | 双向 |
| `title` | `title` | 双向 |
| `notes` | `body.content`（`text`） | 双向 |
| `dueDay` | `dueDateTime` | 双向 |

## 已知限制

- **仅桌面端**：Web/PWA 无 `nodeExecution`，无法访问 Graph。
- **时区**：`dueDay` 以 UTC 午夜写回 `dueDateTime`，非 UTC 用户可能在 To Do 里看到
  日期偏移（如需精确，改用 `dueWithTime`，属 v2）。
- **子任务**：SP 子任务 ↔ To Do `checklistItems` 未实现（独立集合，字段映射无法表达）。
- **标签 / 时间追踪**：To Do 无对应字段，不参与同步。
- **任务深链**：To Do 无稳定公开任务 URL，`getIssueLink` 指向 Web 版应用首页。
- **代理**：Node 子进程的 env 被清空（无 `http_proxy`），系统级代理不会生效。

## 开发

```bash
npm install
npm run typecheck        # 类型检查 src/
npm run build            # 打包 → dist/plugin.js + ms-todo-provider.zip
npm test                 # 全部测试（无凭据时集成测试自动跳过）
npm run test:unit        # 仅纯逻辑单元测试
npm run test:integration # 真实 Graph 往返（需 MS_TODO_ACCESS_TOKEN + MS_TODO_LIST_ID）
```

### 测试

- **单元测试**（`test/unit/`）覆盖配置解析、任务↔issue 映射、日期转换、字段映射
  值变换、分页 —— 无需网络。
- **集成测试**（`test/integration/`）通过 Node 子进程对 Microsoft Graph 做真实
  HTTPS 往返。未提供令牌 + 列表 ID 时自动跳过：

  ```bash
  # PowerShell
  $env:MS_TODO_ACCESS_TOKEN='...'; $env:MS_TODO_LIST_ID='...'
  # bash
  export MS_TODO_ACCESS_TOKEN='...' MS_TODO_LIST_ID='...'
  ```

### 发版

版本单一来源是 `package.json`（`src/manifest.json` 由 bump 脚本保持同步）。发版流程
沿用 zot-tool：

1. 在 PR 里 bump 版本：
   ```bash
   npm run bump:patch   # 或 bump:minor / bump:major，或 node scripts/bump-version.mjs 1.2.3
   ```
2. 把 PR 合并到 `main`。`auto-release` workflow 会读 `package.json`，打 `vX.Y.Z`
   标签并发布 GitHub Release。

## 目录结构

```
ms-todo-plugin/
├── .github/workflows/      # CI：test.yml / auto-release.yml / release.yml
├── package.json
├── tsconfig.json
├── vitest.config.mts       # vitest 配置（node 环境）
├── scripts/
│   ├── build.js            # esbuild 打包 plugin.ts → dist/plugin.js (IIFE)
│   └── bump-version.mjs    # 同时 bump package.json 与 src/manifest.json
├── src/
│   ├── manifest.json
│   ├── types.ts            # 自包含的 plugin-api 类型
│   ├── i18n.ts             # zh/en 运行时多语言辅助
│   ├── graph.ts            # executeNodeScript + require('https') 的 Graph 客户端
│   ├── domain.ts           # 纯逻辑：parseConfig / 映射 / fieldMappings
│   ├── plugin.ts           # registerIssueProvider 主逻辑
│   └── icon.svg
├── test/
│   ├── unit/               # 纯逻辑单元测试
│   └── integration/        # 按凭据开关的真实 Graph 集成测试
├── README.md               # 英文
└── README-zh.md            # 中文
```

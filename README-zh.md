# Microsoft To Do ↔ Super Productivity 同步插件

将 Microsoft To Do 的任务与 Super Productivity（SP）项目**双向同步**的
[issueProvider](https://github.com/super-productivity/super-productivity) 插件。

- **跨平台**：读 / 写均走 SP 的 `http` 通道（Angular HttpClient → XHR），桌面端、
  移动端、Web 端三端可用；不使用 `nodeExecution` 权限，因此在移动端**不会被拒载**。
- **桌面端写请求**：Graph 的写请求（创建 / 更新 / 删除）仍可能被桌面版渲染进程的
  CORS 预检处理拦掉（`status=0 Unknown Error`）—— 这是宿主层面的限制，插件侧无法在
  不牺牲移动端的前提下绕过（要绕过就得用 `nodeExecution`，但那会让移动端拒载）。
- **认证**：OAuth 2.0 设备码流（Entra ID）—— 无需登记重定向 URI，刷新令牌自动续期。
- **同步**：由 SP 的 two-way-sync 引擎驱动，插件只声明字段映射。
- **语言**：配置对话框与信息面板跟随 SP 客户端语言（英文 / 简体中文）。

> ⚠️ 桌面版渲染进程里 Graph 的**写请求**（创建 / 更新 / 删除）可能因宿主 Electron 的
> CORS 预检处理失败为 `status=0 Unknown Error`。本插件**刻意不使用** `nodeExecution`
> （那虽然能修桌面端写回，但会令 SP 在移动端 / Web 端拒载本插件），因此**移动端完全
> 可用**；桌面端的自动创建仍受此宿主层面限制影响。

## 需求

- Super Productivity **18.20+**（桌面端 / 移动端 / Web 端均可）
- 一个 Microsoft Entra ID 应用注册（用于 `clientId`）—— public client，无需登记重定向 URI

## 1. 注册 Entra ID 应用（获取 clientId）

1. 打开 [Azure 门户 → 应用注册 → 新注册](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade)。
2. 名称随意（如 `sp-ms-todo`）；账户类型选「任何组织目录 + 个人 Microsoft 账户」（个人 MSA 用 To Do 需包含个人账户类型）。
3. 注册完成后，记下「应用程序(客户端) ID」。
4. 无需登记重定向 URI —— 插件用 OAuth 2.0 **设备码流**登录，浏览器永不跳回 App，
   `replyUrlsWithType` 留空即可。
5. 开启公共客户端流：把清单顶层 `"allowPublicClient"` 设为 `true`（对应「身份验证 →
   高级设置 → 允许公共客户端流」）。
6. 「API 权限」确认已含（并授予）`Tasks.ReadWrite` 委托权限；`offline_access` 由
   `scope` 参数请求（无需手动加）。

## 2. 配置 clientId（二选一，或两者都要）

clientId **不硬编码在源码里**，有两种配置方式：

**方式 A —— 在 SP 界面里填写（推荐）**：插件配置对话框里有一个 **Client ID** 输入框
（对应 `pluginConfig.clientId`），连接前填上即可；它会覆盖构建时的默认值。

**方式 B —— 构建时从环境变量注入（作为默认值）**：构建前设置 `MS_TODO_CLIENT_ID`，
它会被 esbuild 打进 `dist/plugin.js` 作为默认 clientId：

```bash
# Windows PowerShell
$env:MS_TODO_CLIENT_ID = '你的 Entra 应用 clientId'
# macOS / Linux
export MS_TODO_CLIENT_ID='你的 Entra 应用 clientId'
```

> `clientId` 是 public client，**不是机密**（RFC 8252）；设备码流不需要 client secret，也没有重定向 URI —— 用户在验证页确认设备码即完成授权。

然后构建：

```bash
cd ms-todo-plugin
npm install
npm run build          # 产出 dist/（plugin.js + manifest.json + icon.svg + index.html + ms-todo-provider.zip）
npm run typecheck      # 可选：类型检查
```

## 3. 安装到 Super Productivity

1. 打开 SP → 设置 → 插件 → 「从本地加载」。
2. 选择构建产出的 `dist/ms-todo-provider.zip`。SP **只接受 `.zip`**，且
   `manifest.json` / `plugin.js` / `icon.svg` 必须在压缩包**根目录**（不能套一层文件夹）。
3. 进入任务视图 → 新建/编辑一个 **issue provider**，类型选「Microsoft To Do」。

## 4. 使用

1. 在 provider 配置里点击「测试连接」（它兼任登录按钮）。
2. 弹窗会显示一个验证码 —— 在任意浏览器打开提示的网址、输入验证码并授权。
3. 回到 SP，「To Do 列表」下拉会自动加载，选择一个列表。
4. 该列表的任务会进入项目 backlog；添加为本地任务后即与 To Do 双向同步：
   - 完成状态、标题、备注、截止日期、截止时刻互相同步。

> 插件按分页拉取任务，默认每个列表最多拉 **200** 条，避免超大列表卡顿。可在
> 「**加载任务上限**」下拉里调整（50 / 100 / 200 / 500 / 1000 / 全部）以放宽或取消上限。

## 5. 多个列表 → 多个项目

一个 issue provider 实例只绑定一个 SP 项目（SP 插件 API 限制：导入任务固定落到该
provider 所属项目，无法按 issue 自动跨项目路由）。要把**不同的 To Do 列表**导入到
**不同的 SP 项目**，做法是**为每个列表建一个 provider 实例**并放进目标项目：

1. 打开目标项目 → 新建 issue provider，类型选「Microsoft To Do」。
2. 在「To Do 列表」下拉里选该列表（登录只做一次，同一插件的所有实例共用同一
   `pluginId` 下的令牌）。
3. 重复以上两步，为每个 (列表 → 项目) 组合各建一个实例。

> 令牌按 `pluginId`（经 `setSecret`）共享，多实例无需重复登录；但各实例的 `listId`
> 相互独立，务必为每个实例选对列表。

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

### 完成状态兜底推送（0.1.4+）

宿主只在任务更新动作发生时推送；瞬时失败（网络 / 桌面 CORS `status=0`）不会重试，
`provider-changed` 跳过后也不会补。0.1.4 起插件在 `taskUpdate` hook 上补了一层兜底：
本地给已链接任务打勾后延迟 4 秒检查基线，若宿主已成功推送（`issueLastSyncedValues.state
= completed`）则什么都不做；否则插件自己 GET 远端 → PATCH `status=completed` → 回写
基线（PATCH 幂等，已加入重试）。失败会在 snackbar 与插件日志中可见，并可在「对账」页补修。

拉取方向同样加固（0.1.8+）：插件为每张列表维护进程内任务快照——冷启动整表拉一次
（`$top=1000`，通常 1 次请求），之后每 15 秒只发一条 `lastModifiedDateTime gt …` 增量；
`getByIds` 只读快照，缺失的 id 不再逐个 404 探测。于是 60 秒一轮的轮询通常只有
「每列表 1 次请求」，避免触发 Graph 429 / `ERR_HTTP2_PROTOCOL_ERROR`——实测这些错误会
让宿主的 `getByIds` catch 丢掉**整批**远端更新，表现正是「MS 打勾了、SP 永远收不到」。
同列表的多个 provider 实例共享同一份快照。

限制：MS→SP 仍由宿主的 60 秒轮询驱动，实测远端改动后 **0–60 秒**内落到 SP；移到未配置
列表的任务由跨列表探测识别，不会误删。

### 删除同步（0.1.9+）

**SP → To Do**：宿主内置。删除已链接任务后会等一个「撤销窗口」，窗口结束才调用插件的
`deleteIssue`（Graph DELETE，404 视为已删除，幂等）；单条删除与批量删除都覆盖。注意这是
宿主行为：如果在撤销窗口内退出 App，这次远端删除会被丢弃（远端留下孤儿任务，可在
To Do 里手动清理）。

**To Do → SP**：0.1.9 起支持。每张列表的快照每 **5 分钟**做一次权威整表刷新，远端已删的
条目会从快照消失；`getByIds` 对缺失的 id 做**有界确认**（每轮最多 5 个）：先排除“被移到
别的列表”（跨列表 `getById` 成功则按正常任务返回），确认 404 后才向宿主上报
`state='deleted'` 占位 issue；宿主据此删除本地 SP 任务（若该任务已记录时间，会先弹确认）。
任何 429/网络错误都不会被当作删除。

配置项 **「保留远端已删除的任务」**（默认不勾选）：勾选后不再自动删本地任务，只保留链接，
可在「链接体检」里手动解除。

## 强制同步 & 检测失效链接（插件独立视图）

为手动补救（重新推送）自动创建失败的任务、并排查 `404 ErrorItemNotFound` 反复报错，
插件提供一个**独立视图页面**（`index.html`，即 SP 的「column」页），把所有操作和
**进度打印**集中在一起，不再依赖「选中任务 + 顶部按钮」。

触发同步有两种方式：

- **一键同步（顶部工具栏）**：点击顶部工具栏的 **「同步 Microsoft To Do」** 按钮（
  `check_box` 图标）。它把当前范围内（项目上下文内则该项目，否则全部）的所有未同步任务
  一次性推到各自绑定的 To Do 列表，用「开始 / 结束」两条 snackbar 反馈结果 —— 无需打开
  视图。（头部按钮无法滚动打印逐任务进度，需要详细进度请用下面的独立视图。）
- **独立视图（完整操作 + 进度打印）**：从左侧菜单（桌面端）或左侧菜单抽屉（移动端）打开
  **「Microsoft To Do」** 入口。视图把推送、链接体检、清理等所有操作和进度打印集中在一起。

> 顶部工具栏原先还有一个重复入口（side-panel 按钮，宿主渲染成彩色的插件图标，点击时会
> 先自动打开面板再回调 `onClick`）。它既重复了独立视图、在移动端又只会渲染成空白面板
> （宿主 #8394），因此已移除。

> 视图入口的注册放在 `PluginAPI.onReady` 里、由宿主在插件加载完成后触发——与
> 移动端/桌面端都正常的 `Goals Vision Board` 插件一致，避免顶层注册在移动端丢失。

> 📱 **设备本地安装**：SP 把上传插件的**代码**（`plugin.js` / `index.html`）存在每台设备
> 自己的 IndexedDB 缓存里，**不会随同步在多设备间传输**（同步的只有插件的启用/禁用标记）。
> 所以每次更新构建后，都要**在每台设备上重新安装新的 `ms-todo-provider.zip`**（桌面端和
> 移动端各装一次）。要确认某台设备跑的是否是最新构建：打开视图，状态栏末尾会显示
> **`插件 v0.x.y`**；插件日志（设置 → 插件）里也会打印
> `[ms-todo 0.x.y] plugin.js executing` / `onReady fired` / `registerHeaderButton ok` /
> `registerMenuEntry ok`。若版本偏旧或看不到这些日志，就说明该设备还在跑旧包，需重新安装。

视图按功能分为五个标签页：

- **推送 · Push**：列出**未同步**（自动创建失败、没有 `issueId`）的本地任务，每条显示
  标题 / 所属项目 / 标签 / 备注 / 截止日期（带时刻）。可**勾选多选**后「推送选中」批量
  推送，也可点单行「推送」。推送时把任务（标题 / 备注 / 截止日期（含时刻） / 完成状态）
  POST 到所配置的 To Do 列表并回写 `issueId` 完成链接 —— **不需要先选中 SP 任务**
  （修掉了旧版「请先选中…」的误报）。
- **已链接 · Linked**：当前已同步的本插件任务，`MS id` 以省略号截断、悬停可看完整 id。
- **失效链接 · Dead links**：点「开始检测」逐个 GET 当前项目里的已链接任务，把在 To Do
  里已不存在的任务列出来。删掉这些本地任务即可停止反复 404。
- **对账 · Reconcile**：扫描所有已链接任务的**完成状态分叉**（本地 `isDone` 与上次同步
  基线 `issueLastSyncedValues.state` 不一致），逐条读远端并修复：远端已完成 → 本地打勾；
  本地已完成 → 推送 `completed`；两边一致 → 只回写基线。同时扫描**日期分叉**（本地
  `dueDay`/`dueWithTime` 与基线不一致）：远端日期未变时把本地日期推上去，远端已变则跳过
  并等待轮询拉取。**「本地只剩日期、MS 有提醒时刻」的历史缺口（0.1.10 之前把定时截止读成
  纯日期）按「时间以 MS 为准」处理：把远端时刻补回本地，绝不反向清掉 MS 的提醒**（动作
  `due-pull`）。真正的「用户清掉了时刻」（基线日期与提醒时刻不在同一天）才推空远端提醒
  （动作 `due-push`）。每条结果显示动作与错误，用于修复历史漏推（例如宿主推送瞬时失败、
  来自另一台设备的修改、或拉取批次饿死导致基线过期）。
- **清理 · Cleanup**：按列表查看 To Do 任务并做行内操作——已链接的任务可「解除链接」
  （回到推送页），未链接的可「删除」远端任务（二次确认后从 To Do 删除）。
- 顶部状态栏显示**可读的列表名 / 项目名 / 任务数**；未连接或未选列表时给出配置指引；
  底部 **日志**即时打印每个操作的结果。

> 各标签页都针对**当前项目**（对应**当前 provider 实例所绑定的列表**）。如果你同步了
> 多个列表，请在每个项目里各打开一次视图、各跑一次。

> 视图里点「推送 / 检测」等请求仍走 `PluginAPI.request`（manifest 已声明 `http` +
> `allowedHosts`，与 issueProvider 注入的 `http` 是同一个宿主通道，无
> `nodeExecution`），桌面端 / 移动端 / Web 端一致可用。桌面端仍会受与自动创建相同的
> `status=0` CORS 预检限制（见「已知限制」）。

## 字段映射

| SP 任务 | To Do 任务 | 方向 |
|---|---|---|
| `isDone` | `status`（`completed` / `notStarted`） | 双向 |
| `title` | `title` | 双向 |
| `notes` | `body.content`（`text`） | 双向 |
| `dueDay` | `dueDateTime`（仅日期） | 双向 |
| `dueWithTime` | `reminderDateTime`（+ `isReminderOn`） | 双向 |

## 已知限制

- **桌面端写请求**：宿主的 Electron 渲染进程 CORS 预检处理可能拦掉 Graph 写请求
  （`status=0 Unknown Error`），因此「自动为新任务创建问题」与视图里的「推送」在桌面端
  可能失败。这是宿主层面的限制；插件侧无法在不牺牲移动端的前提下绕过（要用
  `nodeExecution` 才行，但那会令 SP 在移动端拒载本插件），故选择保留移动端兼容。
- **Web/PWA 无法登录**：设备码的 token 端点无 CORS，登录仅在桌面与原生移动端可用
  （与 SP 其它集成一致，SP 会在 Web 端弹跨域警告）。
- **个人微软账户（MSA）**：微软对个人账户的设备码流支持不保证，需按你的账户类型实测；
  工作/学校账户没问题。
- **刷新令牌 90 天滚动**：只要 90 天内同步过至少一次即自动续期；若 90 天完全不使用则需重新登录。
- **日期与时刻（0.1.10+）**：To Do 的数据模型里 **`dueDateTime` 只承载日期**（实测在
  该字段写时刻会被归一化掉），**时刻在 `reminderDateTime`（`isReminderOn`）**。
  - 读取：`dueDateTime` → `dueDay`；`reminderDateTime` → `dueWithTime`。旧版把时刻写进
    due 的历史值仍兼容读取，不会让历史任务被误判为“远端已变”。
  - 写入：计划日 → `<日期>T00:00:00` + **设备时区名**（Graph 接受 IANA，如 `Asia/Shanghai`），
    MS 归一化为当地午夜并显示为纯日期；定时 → `reminderDateTime = 绝对时刻` +
    `isReminderOn=true`，同时把 `dueDateTime` 对齐到该时刻的本地日期；取消定时 → 清掉提醒。
  - 局限：To Do 没有 all-day 标志；设备时区与 To Do 时区不同时，带时刻的提醒可能偏移
    （时区名按设备本地墙钟解析）。
- **子任务**：SP 子任务 ↔ To Do `checklistItems` 未实现（独立集合，字段映射无法表达）。
- **标签 / 时间追踪**：To Do 无对应字段，不参与同步。
- **任务深链**：To Do 无稳定公开任务 URL，`getIssueLink` 指向 Web 版应用首页。

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
- **集成测试**（`test/integration/`）通过插件的 HTTP 通道（Node 内置 `fetch`）对
  Microsoft Graph 做真实 HTTPS 往返。未提供令牌 + 列表 ID 时自动跳过：

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
   标签并发布 GitHub Release（自动构建并附上 `ms-todo-provider.zip` 插件包）。

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
│   ├── graph.ts            # Graph 客户端（PluginHttp，桌面 + 移动 + Web）
│   ├── auth.ts             # 设备码流：登录 + 令牌刷新（setSecret）
│   ├── domain.ts           # 纯逻辑：parseConfig / 映射 / fieldMappings
│   ├── force-sync.ts       # 纯逻辑：强制同步 + 失效链接检测（任务选择）
│   ├── plugin.ts           # registerIssueProvider 主逻辑 + 视图 RPC 桥
│   ├── index.html          # 独立视图（强制同步 / 检测失效链接 / 进度日志）
│   └── icon.svg
├── test/
│   ├── unit/               # 纯逻辑单元测试
│   └── integration/        # 按凭据开关的真实 Graph 集成测试
├── README.md               # 英文
└── README-zh.md            # 中文
```

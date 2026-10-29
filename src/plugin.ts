// Microsoft To Do issue provider for Super Productivity（桌面端 + 移动端）。
//
// 通过 SP 的 issueProvider 一等接口实现，双向同步由宿主 two-way-sync 引擎驱动；
// 本插件只需：设备码流登录、Graph 读写、字段映射。
//
// 数据路径：Graph 请求经 graph.ts 全程走宿主 PluginHttp（`http` 权限，普通浏览器
// 请求，桌面 / 移动 / Web 三端一致，Graph 返回 CORS）。认证走 auth.ts 的设备码流
// （OAuth 2.0 Device Authorization Grant，不依赖 redirect URI，令牌存
// PluginAPI.setSecret/getSecret 本地密钥存储，用 refresh token 自动续期）。
// 纯逻辑（配置解析、映射、字段映射）在 domain.ts，便于单元测试。

import type {
  IssueProviderPluginDefinition,
  PluginHttp,
  PluginIssue,
  PluginSearchResult,
  PluginAPI as PluginApiType,
  SpProject,
  SpTag,
  SpTask,
} from './types';
import {
  buildDueDateTime,
  buildDueDateTimeFromTs,
  extractSyncValues,
  FIELD_MAPPINGS,
  getClientId,
  localDayOf,
  mapTaskToIssue,
  mapTaskToSearchResult,
  parseConfig,
  TODO_WEB_URL,
} from './domain';
import type { TodoTask } from './domain';
import { graphRequest, graphRequestAll } from './graph';
import { TodoTaskCache } from './list-cache';
import {
  buildCreateTaskBody,
  buildProjectToListMap,
  classifyLinkedTasks,
  COMPLETED_STATE,
  decideCompletionSync,
  DELETED_STATE,
  isDeleteSyncEnabled,
  filterIssuesByIds,
  isLegacyTimedBaselineGap,
  findIssueProviderId,
  isMsTodoTask,
  isPushableTask,
  PLUGIN_ISSUE_TYPE,
  resolvePushListId,
  selectCompletionDivergences,
  selectDueDivergences,
  selectLinkedMsTasks,
  selectTasksToSync,
  shouldPushCompletionOnUpdate,
  withCompletedBaseline,
} from './force-sync';
import {
  createDeviceCodeDialog,
  createSecretStore,
  getAccessToken,
  runDeviceCodeLogin,
  SECRET_REFRESH_TOKEN,
  withHostDbRetry,
} from './auth';
import { t } from './i18n';
import { createRequestHttp, hasRequestApi } from './request-http';

declare const PluginAPI: PluginApiType;

// 供 esbuild `define` 在构建期注入环境变量（编译期替换，非运行时读取）。
declare const process: { env: Record<string, string | undefined> };

// 默认 clientId：构建时从环境变量 MS_TODO_CLIENT_ID 注入；未设置则为空串。
// 运行时用户可在插件配置里用「Client ID」字段覆盖。
const CLIENT_ID = process.env.MS_TODO_CLIENT_ID || '';

// 构建版本（构建时由 scripts/build.js 从 package.json 注入）。独立视图会展示它，
// 便于在移动端确认设备上跑的是不是最新构建（插件代码是设备本地缓存，不随 SP 同步）。
const PLUGIN_VERSION = process.env.PLUGIN_VERSION || 'dev';

// 密钥存储（本地、每设备、不进 sync/备份），模块级单例即可。
const secrets = createSecretStore();

// 最近一次 provider 调用收到的 config 的持久副本（不含令牌），见 cacheContext。
const CONFIG_CACHE_SECRET = 'ms-todo-cached-config';

// 生命周期日志（写入宿主的插件日志，桌面 / 移动端都能在「设置 → 插件」里查看）。
// 任何一步都不能抛错：这里只用于在移动端排查「视图入口为何没出现」时定位卡在哪一步。
function lifecycle(msg: string): void {
  // 同时写 console：桌面端可用 CDP 抓取，移动端仍以宿主插件日志为准。
  try {
    const g = globalThis as { console?: { log?: (...args: unknown[]) => void } };
    g.console?.log?.(`[ms-todo ${PLUGIN_VERSION}] ${msg}`);
  } catch {
    // 忽略 console 失败。
  }
  try {
    PluginAPI.log?.info?.(`[ms-todo ${PLUGIN_VERSION}] ${msg}`);
  } catch {
    // 忽略日志本身的失败。
  }
}

lifecycle('plugin.js executing');

// 缓存最近一次 issueProvider 调用收到的 config / http。宿主只在 issueProvider 方法
// 里注入这两个参数，而侧边栏按钮的回调拿不到它们，所以只能从这里取。轮询（60s）
// 或任何一次自动创建（即便失败）都会刷新缓存。
let cachedConfig: Record<string, unknown> | undefined;
let cachedHttp: PluginHttp | undefined;

function cacheContext(config: Record<string, unknown>, http: PluginHttp): void {
  cachedConfig = config;
  cachedHttp = http;
  // 再持久化一份：宿主重启后、第一次轮询调用 provider 之前，插件拿不到注入的
  // config，兜底对账（每 5 分钟）会整轮跳过。配置里没有令牌（令牌单独存），
  // 放进设备本地密钥存储没有额外暴露面；内容没变就不重复写。
  try {
    const json = JSON.stringify(config);
    if (json !== lastPersistedConfigJson) {
      lastPersistedConfigJson = json;
      void secrets.set(CONFIG_CACHE_SECRET, json).catch(() => {
        lastPersistedConfigJson = '';
      });
    }
  } catch {
    // 配置不可序列化时忽略，绝不影响 provider 调用。
  }
}

let lastPersistedConfigJson = '';

// MS 任务 id → 所在列表 id 的小缓存（有界）。用途：
//   - getByIds 批量拉取时记住每个任务的列表；
//   - 之后 getById / updateIssue / deleteIssue 先按已知列表请求，避免每次都靠 404
//     探测真实列表（findTaskList 虽只列一次列表，但仍是额外往返）。
// 任务被挪走后缓存会过期，404 自愈仍会纠正并更新缓存。
const issueListCache = new Map<string, string>();
const ISSUE_LIST_CACHE_MAX = 3000;

// 404 负缓存：确认所有列表都找不到的 issueId，10 分钟内不再重复「列全部列表 +
// 每列表 probe」。视图对账/失效链接扫描可能反复碰到已删除任务，没有这个缓存
// 会再次制造 Graph 请求风暴（2026-10-26 实测每轮数十个 404）。
const missingIssueCache = new Map<string, number>();
const MISSING_ISSUE_TTL_MS = 10 * 60_000;
const MISSING_ISSUE_CACHE_MAX = 3000;

function isRecentlyMissing(issueId: string): boolean {
  const at = missingIssueCache.get(issueId);
  if (at === undefined) return false;
  if (Date.now() - at < MISSING_ISSUE_TTL_MS) return true;
  missingIssueCache.delete(issueId);
  return false;
}

function rememberMissing(issueId: string): void {
  missingIssueCache.delete(issueId);
  missingIssueCache.set(issueId, Date.now());
  if (missingIssueCache.size > MISSING_ISSUE_CACHE_MAX) {
    const oldest = missingIssueCache.keys().next().value;
    if (oldest) missingIssueCache.delete(oldest);
  }
}

function rememberIssueList(issueId: string, listId: string | undefined): void {
  if (!issueId || !listId) return;
  if (issueListCache.has(issueId)) issueListCache.delete(issueId);
  issueListCache.set(issueId, listId);
  if (issueListCache.size > ISSUE_LIST_CACHE_MAX) {
    const oldest = issueListCache.keys().next().value;
    if (oldest) issueListCache.delete(oldest);
  }
}

function knownIssueList(issueId: string): string | undefined {
  return issueListCache.get(issueId);
}

// 每个列表的任务快照（进程内、设备本地）：把轮询从「每轮整表分页 + 缺失 id
// 逐个 404 探测」降为「冷启动一次整表，之后每 TTL 一条 lastModified 增量」。
// 2026-10-26 实测：旧实现每分钟数百次 Graph 请求 → 429/HTTP2 错误 →
// 宿主 getByIds 的 catch 丢弃整批更新（MS 完成状态因此长期拉不回来）。
const taskCache = new TodoTaskCache({
  onError: (message, error) => {
    lifecycle(message + ': ' + (error instanceof Error ? error.message : String(error)));
  },
  onRefresh: (info) => {
    lifecycle(
      `cache ${info.mode} list=…${info.listId.slice(-8)}` +
        (info.since ? ` since=${info.since}` : '') +
        (info.changed !== undefined ? ` changed=${info.changed}` : '') +
        ` cached=${info.cached}` +
        (info.error ? ` error=${info.error}` : ''),
    );
  },
});

// 顶部按钮拿不到 issueProvider 注入的 http，优先用最近一次缓存；没有就退到
// PluginAPI.request（manifest 已声明 http + allowedHosts）。两者底层同一个
// PluginHttpService，桌面 / 移动行为一致。
let requestHttp: PluginHttp | undefined;

function resolveHttp(): PluginHttp {
  if (cachedHttp) return cachedHttp;
  if (!requestHttp) {
    if (!hasRequestApi()) {
      throw new Error(
        t({
          zh:
            '当前 SP 版本不支持 PluginAPI.request：请先打开插件配置并选择列表，或等待一次同步完成后再点。',
          en:
            'PluginAPI.request is unavailable in this SP version. Open the plugin settings and pick a list, or wait for the first sync, then try again.',
        }),
      );
    }
    requestHttp = createRequestHttp();
  }
  return requestHttp;
}

/** 取一个有效的 Graph 访问令牌（未连接时抛「未连接」错误）。 */
function getToken(config: Record<string, unknown>, http: PluginHttp): Promise<string> {
  return getAccessToken(getClientId(config, CLIENT_ID), http, secrets);
}

// --- Graph 数据类型 ---

interface TodoList {
  id: string;
  displayName: string;
  wellknownListName?: string;
}

/** 列表总览里的单个列表（含任务，任务在 status 里再补上「是否已链接到 SP」）。 */
interface TodoListInfo {
  id: string;
  displayName: string;
  taskCount: number;
  tasks: { id: string; title: string; status: string }[];
}

/** 全列表总览：所有列表 + 「MS 任务 id → 列表 id」索引（用于反推项目↔列表映射）。 */
interface ListsOverview {
  lists: TodoListInfo[];
  msTaskToList: Map<string, string>;
}

// 加载 To Do 列表（登录后由宿主自动调用，用于 listId 下拉）。
// 未配置 / 未连接 / 网络失败时返回空数组而非抛错 —— 与 SP 官方 Trello 插件的
// loadBoards 保持一致：loadOptions 抛错会让宿主把下拉置为 error 态，某些版本还会
// 反复重试/重渲染，导致配置界面卡死。
async function loadLists(
  config: Record<string, unknown>,
  http: PluginHttp,
): Promise<{ label: string; value: string }[]> {
  cacheContext(config, http);
  try {
    const clientId = getClientId(config, CLIENT_ID);
    // 已配置/已同步的 provider 在本设备可能还没有登录态（setSecret 每设备本地、
    // 不随 SP 同步）。此时先走一次设备码登录，否则永远拉不到列表。而 SP 的
    // 「测试连接」按钮只在新建 provider 时显示，已存在（编辑态）的 provider 没有
    // 登录入口，所以登录只能挂在这里（列表下拉的 loadOptions 触发）。
    if (!(await secrets.get(SECRET_REFRESH_TOKEN))) {
      await runDeviceCodeLogin(clientId, http, createDeviceCodeDialog(), secrets);
    }
    const token = await getAccessToken(clientId, http, secrets);
    const lists = await graphRequestAll<TodoList>('/me/todo/lists', token, 0, http);
    return lists.map((l) => ({ label: l.displayName, value: l.id }));
  } catch {
    return [];
  }
}

// --- 插件注册 ---

const issueProviderDefinition: IssueProviderPluginDefinition = {
  configFields: [
    {
      key: 'clientId',
      type: 'input',
      label: t({ zh: '客户端 ID', en: 'Client ID' }),
      description: t({
        zh:
          '你的 Entra ID 应用「应用程序(客户端) ID」（public client，非机密，无需 redirect URI）。' +
          '构建时若未通过 MS_TODO_CLIENT_ID 环境变量注入，请在此填写，否则无法连接。',
        en:
          'Your Entra ID application (client) ID — a public client, not a secret, no redirect URI needed. ' +
          'Fill this in if it was not injected at build time via the MS_TODO_CLIENT_ID env var, ' +
          'otherwise you cannot connect.',
      }),
    },
    {
      key: 'listId',
      type: 'select',
      label: t({ zh: 'To Do 列表', en: 'To Do list' }),
      description: t({
        zh:
          '与当前 SP 项目对应的 To Do 列表（连接账户后自动加载）。' +
          '不同列表要导入不同项目时，请为每个列表各建一个 provider 实例。',
        en:
          'The To Do list mapped to this SP project (auto-loads after connecting your account). ' +
          'To import different lists into different projects, create one provider instance per list.',
      }),
      required: true,
      options: [],
      loadOptions: loadLists,
    },
    {
      key: 'maxTasks',
      type: 'select',
      label: t({ zh: '加载任务上限', en: 'Max tasks to load' }),
      description: t({
        zh:
          '单次从 To Do 拉取的任务数上限，防止超大列表拖慢/卡死。' +
          '默认 200，设为「全部」则不限制。',
        en:
          'Maximum number of tasks to pull from To Do per sync, to avoid lag on huge lists. ' +
          'Defaults to 200; "All" disables the limit.',
      }),
      options: [
        { label: '50', value: '50' },
        { label: '100', value: '100' },
        { label: '200', value: '200' },
        { label: '500', value: '500' },
        { label: '1000', value: '1000' },
        { label: t({ zh: '全部（无上限）', en: 'All (no limit)' }), value: '0' },
      ],
    },
    {
      key: 'keepRemoteDeletedTasks',
      type: 'checkbox',
      label: t({ zh: '保留远端已删除的任务', en: 'Keep tasks deleted in To Do' }),
      description: t({
        zh:
          '默认不勾选：在 Microsoft To Do 里删除任务时，SP 里的对应任务也会被删除（约 5 分钟内生效；' +
          '已记录时间的任务会先弹确认）。勾选后不删除本地任务，只保留链接，可在「链接体检」里手动解除。',
        en:
          'Unchecked by default: deleting a task in Microsoft To Do also deletes the linked SP task ' +
          '(within ~5 minutes; tasks with tracked time ask first). Check to keep the local task and ' +
          'only keep the link; unlink manually from Link health.',
      }),
    },
    {
      key: 'clientIdHelp',
      type: 'link',
      label: t({
        zh: '如何获取 clientId（Entra 应用注册）',
        en: 'How to get a clientId (Entra app registration)',
      }),
      url: 'https://learn.microsoft.com/entra/identity-platform/quickstart-register-app',
    },
  ],

  // Graph 请求的头（含 Bearer 令牌）由 graphRequest 在每次调用时注入（统一走
  // PluginHttp），故这里返回空即可。
  getHeaders(_config: Record<string, unknown>): Record<string, string> {
    return {};
  },

  async searchIssues(
    searchTerm: string,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginSearchResult[]> {
    cacheContext(config, http);
    try {
      const { listId, maxTasks } = parseConfig(config);
      const token = await getToken(config, http);
      const all = await taskCache.getAll(listId, token, http);
      const tasks = maxTasks > 0 ? all.slice(0, maxTasks) : all;
      const term = searchTerm.trim().toLowerCase();
      return tasks
        .filter((t) => !term || t.title.toLowerCase().includes(term))
        .map(mapTaskToSearchResult);
    } catch {
      return [];
    }
  },

  async getNewIssuesForBacklog(
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginSearchResult[]> {
    cacheContext(config, http);
    try {
      const { listId, maxTasks } = parseConfig(config);
      const token = await getToken(config, http);
      const all = await taskCache.getAll(listId, token, http);
      const tasks = maxTasks > 0 ? all.slice(0, maxTasks) : all;
      return tasks.map(mapTaskToSearchResult);
    } catch {
      return [];
    }
  },

  async getById(
    issueId: string,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginIssue> {
    cacheContext(config, http);
    const { listId } = parseConfig(config);
    const token = await getToken(config, http);
    const { task } = await fetchTodoTask(issueId, listId, token, http);
    return mapTaskToIssue(task);
  },

  // 批量版 getById：宿主刷新大量已链接任务时会优先调用它。不实现时宿主会退化为
  // 「每个任务一次 GET」的串行循环——任务一多一轮跑不完，批次尾部会被下一次
  // timer/上下文切换取消，表现为「MS 端改了、SP 永远收不到」。
  //
  // 实现要点（见 list-cache.ts）：读进程内列表快照（整表每 5 分钟重建一次），
  // 整批失败时返回缓存，避免宿主 catch→[] 丢整批。缺失的 id 不再无条件跳过：
  // 快照可信时做有界确认探测——被挪到别的列表的按正常 issue 返回；确认 404 的
  // 上报 `state='deleted'` 占位 issue，由宿主删除本地任务（deletedStates）。
  async getByIds(
    issueIds: string[],
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginIssue[]> {
    cacheContext(config, http);
    const ids = [...new Set(issueIds.filter((id) => !!id))];
    if (ids.length === 0) return [];
    try {
      const { listId } = parseConfig(config);
      const token = await getToken(config, http);
      const { tasks, missing, trustworthy } = await taskCache.getByIds(
        listId,
        ids,
        token,
        http,
      );
      lifecycle(
        `getByIds list=…${listId.slice(-8)} want=${ids.length} found=${tasks.length} missing=${missing.length} trustworthy=${trustworthy}`,
      );
      for (const task of tasks) {
        rememberIssueList(task.id, taskCache.listIdOf(task.id) ?? listId);
      }
      const issues = filterIssuesByIds(tasks, ids).map(mapTaskToIssue);
      if (trustworthy && missing.length && isDeleteSyncEnabled(config)) {
        issues.push(...(await confirmRemoteDeletions(missing, listId, token, http)));
      }
      return issues;
    } catch (e) {
      // 配置/令牌层面的错误：返回空而不是抛给宿主（抛错会让宿主把整批当失败）。
      lifecycle(
        'getByIds failed: ' + (e instanceof Error ? e.message : String(e)),
      );
      return [];
    }
  },

  getIssueLink(_issueId: string, _config: Record<string, unknown>): string {
    cachedConfig = _config;
    // To Do 任务无稳定公开深链，链接到 Web 版应用首页。
    return TODO_WEB_URL;
  },

  // 「测试连接」按钮复用为登录入口：未登录时先走设备码流，已登录则校验连接。
  async testConnection(
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<boolean> {
    cacheContext(config, http);
    try {
      const clientId = getClientId(config, CLIENT_ID);
      if (!(await secrets.get(SECRET_REFRESH_TOKEN))) {
        await runDeviceCodeLogin(clientId, http, createDeviceCodeDialog(), secrets);
      }
      const token = await getAccessToken(clientId, http, secrets);
      await graphRequest('GET', '/me/todo/lists', token, undefined, http);
      return true;
    } catch (e) {
      PluginAPI.log.warn('Microsoft To Do connection test failed', e);
      return false;
    }
  },

  issueDisplay: [
    { field: 'title', label: t({ zh: '标题', en: 'Title' }), type: 'text' },
    { field: 'status', label: t({ zh: '状态', en: 'Status' }), type: 'text' },
    {
      field: 'importance',
      label: t({ zh: '优先级', en: 'Priority' }),
      type: 'text',
      hideEmpty: true,
    },
    {
      field: 'dueDisplay',
      label: t({ zh: '截止日期', en: 'Due date' }),
      type: 'text',
      hideEmpty: true,
    },
    {
      field: 'body',
      label: t({ zh: '备注', en: 'Notes' }),
      type: 'markdown',
      hideEmpty: true,
    },
  ],

  fieldMappings: FIELD_MAPPINGS,

  // 远端已删除：getByIds 对确认 404 的 id 返回 state='deleted' 的占位 issue，
  // 宿主据此走 `_handleRemoteDeletion`（删本地任务；有计时记录时先弹确认）。
  deletedStates: [DELETED_STATE],

  async updateIssue(
    id: string,
    changes: Record<string, unknown>,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<void> {
    cacheContext(config, http);
    const { listId } = parseConfig(config);
    const token = await getToken(config, http);

    const patch: Record<string, unknown> = {};
    if (changes.title !== undefined) {
      patch.title = changes.title;
    }
    if (changes.state !== undefined) {
      patch.status = changes.state;
    }
    if (changes.body !== undefined) {
      patch.body = { content: changes.body as string, contentType: 'text' };
    }
    // dueWithTime 优先：SP 把「带时刻」与「纯日期」分成两个互斥字段，带时刻的
    // 变更（schedule/reSchedule/unschedule）以 dueWithTime 到达；纯日期（planner
    // planForDay / transferTask）以 dueDate 到达。null 表示清除截止。
    if (changes.dueWithTime !== undefined) {
      // To Do 的“时刻”是 reminderDateTime；dueDateTime 只承载日期（Graph 会归一化
      // 其中的时刻）。定时 → 设提醒并把 due 日期对齐到该时刻的本地日期；
      // 取消定时（null）→ 清掉提醒，保留已有日期。
      const reminder = buildDueDateTimeFromTs(changes.dueWithTime as number | null);
      patch.reminderDateTime = reminder;
      patch.isReminderOn = reminder !== null;
      if (typeof changes.dueWithTime === 'number' && Number.isFinite(changes.dueWithTime)) {
        patch.dueDateTime = buildDueDateTime(localDayOf(changes.dueWithTime));
      }
    }
    // 独立判断而不是 else-if：清掉时刻（dueWithTime=null）时还要同时写日期，
    // 否则「定时 → 纯日期」只会清提醒、日期留在旧值。
    if (changes.dueDate !== undefined) {
      patch.dueDateTime = buildDueDateTime(changes.dueDate as string | null);
    }

    if (Object.keys(patch).length === 0) return;
    let usedListId = knownIssueList(id) ?? listId;
    try {
      await graphRequest('PATCH', `/me/todo/lists/${usedListId}/tasks/${id}`, token, patch, http);
    } catch (e) {
      if (!isNotFoundError(e)) throw e;
      const actual = await findTaskList(id, token, http);
      if (!actual) throw e;
      usedListId = actual;
      await graphRequest('PATCH', `/me/todo/lists/${actual}/tasks/${id}`, token, patch, http);
    }
    rememberIssueList(id, usedListId);
  },

  async createIssue(
    title: string,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<{ issueId: string; issueData: PluginIssue }> {
    cacheContext(config, http);
    const { listId } = parseConfig(config);
    const token = await getToken(config, http);
    const task = await graphRequest<TodoTask>(
      'POST',
      `/me/todo/lists/${listId}/tasks`,
      token,
      { title },
      http,
    );
    taskCache.remember(listId, task);
    return {
      issueId: task.id,
      issueData: mapTaskToIssue(task),
    };
  },

  async deleteIssue(
    id: string,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<void> {
    cacheContext(config, http);
    const { listId } = parseConfig(config);
    const token = await getToken(config, http);
    let usedListId = knownIssueList(id) ?? listId;
    try {
      await graphRequest('DELETE', `/me/todo/lists/${usedListId}/tasks/${id}`, token, undefined, http);
    } catch (e) {
      if (!isNotFoundError(e)) throw e;
      const actual = await findTaskList(id, token, http);
      if (!actual) throw e;
      usedListId = actual;
      await graphRequest('DELETE', `/me/todo/lists/${actual}/tasks/${id}`, token, undefined, http);
    }
    issueListCache.delete(id);
    taskCache.forget(id);
  },

  extractSyncValues,
};

try {
  PluginAPI.registerIssueProvider(issueProviderDefinition);
} catch (e) {
  lifecycle(
    'registerIssueProvider failed: ' + (e instanceof Error ? e.message : String(e)),
  );
}
lifecycle('registerIssueProvider ok');

// --- 独立视图（index.html）的桥接 RPC ---
//
// 顶部按钮不直观、也没有进度打印，因此改用 SP 的 iFrame 视图（manifest iFrame:true +
// index.html）承载「强制同步 / 检测失效链接」操作与进度展示。宿主只把 getConfig /
// getSecret 暴露给 plugin.js（iframe 里拿不到），所以 Graph 读写、令牌、配置都留在这
// 里；视图经宿主的消息桥（PLUGIN_MESSAGE）调用下面的 handler，拿到纯 JSON 结果自行
// 渲染进度。

interface RpcResult {
  ok: boolean;
  error?: string;
  [key: string]: unknown;
}

function toErrorResult(e: unknown): RpcResult {
  return { ok: false, error: e instanceof Error ? e.message : String(e) };
}

function missingConfigError(): Error {
  return new Error(
    t({
      zh:
        '缺少插件配置：请先打开插件配置（设置 → 集成）选择列表，或等待一次同步完成后再试。',
      en:
        'Plugin config is not ready. Open the plugin settings and pick a list, or wait for the first sync, then try again.',
    }),
  );
}

async function resolveConfig(): Promise<Record<string, unknown>> {
  if (cachedConfig) return cachedConfig;
  try {
    const raw = await secrets.get(CONFIG_CACHE_SECRET);
    if (raw) {
      cachedConfig = JSON.parse(raw) as Record<string, unknown>;
      return cachedConfig;
    }
  } catch {
    // 读缓存失败就继续走宿主 API。
  }
  const config = (await withHostDbRetry(() => PluginAPI.getConfig())) ?? undefined;
  if (!config) throw missingConfigError();
  return config;
}

// 全列表总览缓存：打开视图时构建一次，60s 内复用（推送时直接吃缓存，避免重复拉取）。
let overviewCache: ListsOverview | undefined;
let overviewFetchedAt = 0;
const OVERVIEW_TTL_MS = 60_000;

/**
 * 拉取账号下所有 To Do 列表 + 每个列表的任务，并建立「MS 任务 id → 列表 id」索引。
 *
 * 这是「按列表总览」和「项目↔列表映射」的数据基础：Graph 本身就能枚举全部列表，
 * 无需 SP 暴露任何东西；而宿主 issueProvider 单实例模型不会告诉插件每个项目的 listId，
 * 只能靠已链接任务的 issueId 反向定位它落在哪个列表。
 */
async function buildOverview(limit?: number): Promise<ListsOverview> {
  // 默认（未传 limit）走缓存；显式传 limit（例如体检时传 0 = 不限量）强制重新拉取，
  // 且结果不写缓存，避免把超大结果常驻内存。
  if (limit === undefined && overviewCache && Date.now() - overviewFetchedAt < OVERVIEW_TTL_MS) {
    return overviewCache;
  }
  const config = await resolveConfig();
  const clientId = getClientId(config, CLIENT_ID);
  const http = resolveHttp();
  const token = await getAccessToken(clientId, http, secrets);
  const { maxTasks } = parseConfig(config);
  const effectiveMax = limit !== undefined ? limit : maxTasks;

  const lists = await graphRequestAll<TodoList>('/me/todo/lists', token, 0, http);
  const msTaskToList = new Map<string, string>();
  const listsInfo: TodoListInfo[] = [];
  for (const list of lists) {
    const tasks = await graphRequestAll<TodoTask>(
      `/me/todo/lists/${list.id}/tasks?$top=1000`,
      token,
      effectiveMax,
      http,
    );
    for (const task of tasks) {
      msTaskToList.set(task.id, list.id);
    }
    listsInfo.push({
      id: list.id,
      displayName: list.displayName,
      taskCount: tasks.length,
      tasks: tasks.map((t) => ({ id: t.id, title: t.title, status: t.status })),
    });
  }

  const overview: ListsOverview = { lists: listsInfo, msTaskToList };
  if (limit === undefined) {
    overviewCache = overview;
    overviewFetchedAt = Date.now();
  }
  return overview;
}

/** 判断 Graph 错误是否为 404（远端任务已不存在）。graphRequest 会把宿主 status 挂到 Error 上。 */
function isNotFoundError(e: unknown): boolean {
  const status = (e as { status?: unknown } | null | undefined)?.status;
  if (status === 404 || status === 410) return true;
  // To Do 不是所有「此列表里没有该任务」都报 404：按错误的 listId 读取时 Graph 返回
  // 400 invalidRequest + "Parent folder specified does not contain a Task with given
  // Id"（2026-10-29 实测）。不把它当 not-found 就永远不会走跨列表自愈，
  // 表现为该任务的所有读写都失败。
  const message = e instanceof Error ? e.message : String(e ?? '');
  return /does not contain a Task with given Id|ErrorItemNotFound|itemNotFound/i.test(message);
}

/**
 * 定位某个 MS 任务实际所在的列表：逐个列表 GET `/lists/{id}/tasks/{issueId}`，第一个
 * 200 的列表即真实归属；全 404 则返回 undefined（任务已不存在）。非 404 错误（401 /
 * 429 / 网络）直接抛出，不吞掉。
 *
 * MS 任务 id 是全局唯一 GUID，所以「在哪个列表 GET 到它」就唯一确定了它当前所在列表。
 * 用于写路径的 404 自愈：任务被用户从配置列表挪走后，宿主按旧 listId 拼的路径会 404，
 * 这里兜底找到真实列表再重试。
 */
/**
 * 单条读取一个 To Do 任务（404 时跨列表定位），并把结果写入列表快照，
 * 供 getById 与删除确认探测共用。
 */
async function fetchTodoTask(
  issueId: string,
  listId: string,
  token: string,
  http: PluginHttp,
): Promise<{ task: TodoTask; listId: string }> {
  let usedListId = knownIssueList(issueId) ?? listId;
  const task = await graphRequest<TodoTask>(
    'GET',
    `/me/todo/lists/${usedListId}/tasks/${issueId}`,
    token,
    undefined,
    http,
  ).catch(async (e) => {
    // 任务被挪到别的列表后，旧 listId 路径 404：定位真实列表再取一次。
    if (!isNotFoundError(e)) throw e;
    const actual = await findTaskList(issueId, token, http);
    if (!actual) throw e;
    usedListId = actual;
    return graphRequest<TodoTask>(
      'GET',
      `/me/todo/lists/${actual}/tasks/${issueId}`,
      token,
      undefined,
      http,
    );
  });
  rememberIssueList(issueId, usedListId);
  missingIssueCache.delete(issueId);
  taskCache.remember(usedListId, task);
  return { task, listId: usedListId };
}

/** 单轮最多确认多少个远端删除：坏快照时避免一次误删一大批。 */
const MAX_DELETION_PROBES_PER_CYCLE = 5;

/** 占位 issue：只用于把“远端已删除”告诉宿主（provider 上声明了 deletedStates）。 */
function deletedIssue(id: string): PluginIssue {
  return {
    id,
    title: '',
    body: '',
    url: TODO_WEB_URL,
    state: DELETED_STATE,
    lastUpdated: 0,
  };
}

/**
 * 对快照里缺失的 id 做有界确认：
 *   - 10 分钟内已确认缺失的 → 本轮不再上报（宿主的删除/警告已在上一轮发出）；
 *   - getById 成功 → 任务只是被移到别的列表，按正常 issue 返回（并写入缓存）；
 *   - 404（含跨列表探测）→ 确认远端已删，上报 deleted 占位 issue；
 *   - 其它错误（429/网络）→ 什么都不做，绝不当作删除。
 */
async function confirmRemoteDeletions(
  missing: string[],
  listId: string,
  token: string,
  http: PluginHttp,
): Promise<PluginIssue[]> {
  const issues: PluginIssue[] = [];
  let probes = 0;
  for (const id of missing) {
    if (isRecentlyMissing(id)) continue;
    if (probes >= MAX_DELETION_PROBES_PER_CYCLE) break;
    probes++;
    try {
      const { task } = await fetchTodoTask(id, listId, token, http);
      issues.push(mapTaskToIssue(task));
    } catch (e) {
      if (!isNotFoundError(e)) {
        lifecycle(
          `deletion probe failed for ${id}: ${e instanceof Error ? e.message : String(e)}`,
        );
        continue;
      }
      rememberMissing(id);
      taskCache.forget(id);
      lifecycle(`remote deletion confirmed: ${id}`);
      issues.push(deletedIssue(id));
    }
  }
  return issues;
}

async function findTaskList(
  issueId: string,
  token: string,
  http: PluginHttp,
): Promise<string | undefined> {
  if (isRecentlyMissing(issueId)) return undefined;
  const lists = await graphRequestAll<TodoList>('/me/todo/lists', token, 0, http);
  for (const list of lists) {
    try {
      await graphRequest<TodoTask>(
        'GET',
        `/me/todo/lists/${list.id}/tasks/${issueId}`,
        token,
        undefined,
        http,
      );
      rememberIssueList(issueId, list.id);
      return list.id;
    } catch (e) {
      if (!isNotFoundError(e)) throw e;
      // 404 → 不在这个列表，继续找下一个。
    }
  }
  rememberMissing(issueId);
  return undefined;
}

/** 当前项目 id（仅当用户正处在 PROJECT 工作上下文时返回）。 */
async function currentProjectId(): Promise<string | undefined> {
  try {
    const ctx = await PluginAPI.getActiveWorkContext();
    return ctx?.type === 'PROJECT' ? ctx.id : undefined;
  } catch {
    return undefined;
  }
}

// 安全的 getAllProjects / getAllTags：老版本宿主可能没有这些方法，退化为空数组。
async function safeGetAllProjects(): Promise<SpProject[]> {
  try {
    return (await PluginAPI.getAllProjects()) ?? [];
  } catch {
    return [];
  }
}

async function safeGetAllTags(): Promise<SpTag[]> {
  try {
    return (await PluginAPI.getAllTags()) ?? [];
  } catch {
    return [];
  }
}

/** 取全部 SP 任务。宿主任务存储同样是 IndexedDB，移动端会偶发「连接正在关闭」瞬态错误，故重试。 */
async function getAllTasks(): Promise<SpTask[]> {
  // 宿主的 getTasks() 只返回未完成任务：任务一打勾就进 archive，插件看不到。
  // 完成态对账 / 兜底必须把归档一起读出来，否则「本地已完成、远端未完成」的分叉
  // 永远不会被发现（2026-10-28 的「找黑色裤」即如此）。
  const [active, archived] = await Promise.all([
    withHostDbRetry(() => PluginAPI.getTasks()),
    typeof PluginAPI.getArchivedTasks === 'function'
      ? withHostDbRetry(() => PluginAPI.getArchivedTasks?.() ?? Promise.resolve([]))
      : Promise.resolve([]),
  ]);
  if (!archived.length) return active;
  const seen = new Set(active.map((t) => t.id));
  const merged = active.slice();
  for (const task of archived) {
    if (!seen.has(task.id)) merged.push(task);
  }
  return merged;
}

// --- 完成状态兜底同步 ---
//
// 宿主的 two-way-sync 推送只在任务更新动作时触发，失败（瞬时网络 / 桌面 CORS
// status=0）不会重试；provider-changed 跳过后也不会补。拉取方向又可能因批次过长
// 被宿主取消。这里在插件侧补一条“幂等 + 可重试 + 可观测”的完成状态通道。

const baselineStateOf = (task: SpTask): string | undefined => {
  const state = task.issueLastSyncedValues?.['state'];
  return typeof state === 'string' ? state : undefined;
};

// 兜底：个别宿主版本会对 PluginAPI.updateTask 的字段做白名单校验，导致
// issueLastSyncedValues 写不进去（该字段不在公开 Task 类型里）。此时用进程内标记
// 防止同一次打勾被重复推送；远端写入本身已经成功，不影响用户目标。
const pushedCompletionIssues = new Set<string>();
const PUSHED_MARKER_MAX = 3000;

function markPushedCompletion(issueId: string): void {
  pushedCompletionIssues.add(issueId);
  if (pushedCompletionIssues.size > PUSHED_MARKER_MAX) {
    const oldest = pushedCompletionIssues.values().next().value;
    if (oldest) pushedCompletionIssues.delete(oldest);
  }
}

/**
 * 对单个任务执行一次完成状态对齐（读远端 → 决策 → 必要时 PATCH / 仅回写基线）。
 * 返回决策，供视图逐条展示。
 */
async function syncCompletionForTask(
  task: SpTask,
): Promise<{ action: string; reason: string }> {
  const config = await resolveConfig();
  const http = resolveHttp();
  const issueId = task.issueId as string;
  const updateIssue = issueProviderDefinition.updateIssue;
  if (!updateIssue) throw new Error('updateIssue not implemented');

  const fresh = await issueProviderDefinition.getById(issueId, config, http);
  const baseline = task.issueLastSyncedValues;
  const decision = decideCompletionSync({
    isDone: !!task.isDone,
    baselineState: baselineStateOf(task),
    remoteState: fresh.state,
  });

  const writeBaseline = async (lastUpdated?: number): Promise<void> => {
    const changes: Record<string, unknown> = {
      issueLastSyncedValues: withCompletedBaseline(baseline),
    };
    if (typeof lastUpdated === 'number' && Number.isFinite(lastUpdated)) {
      changes['issueLastUpdated'] = lastUpdated;
    }
    // 带 issueLastSyncedValues 的 updateTask 会被宿主推送效果器过滤，不会造成回环。
    try {
      await withHostDbRetry(() => PluginAPI.updateTask(task.id, changes));
    } catch (e) {
      // 宿主白名单拒绝该字段时不要影响远端已完成的既成事实，改用进程内标记兜底。
      lifecycle(
        'baseline write failed for ' + task.id + ': ' +
          (e instanceof Error ? e.message : String(e)),
      );
    }
  };

  if (decision.action === 'push') {
    await updateIssue(issueId, { state: COMPLETED_STATE }, config, http);
    markPushedCompletion(issueId);
    // 写完再读一次，拿到服务端新的 lastModifiedDateTime 作为基线时间戳。
    const post = await issueProviderDefinition
      .getById(issueId, config, http)
      .catch(() => fresh);
    await writeBaseline(post.lastUpdated);
  } else if (decision.action === 'pull') {
    // 远端已完成、本地还没打勾：直接本地置为完成（宿主的拉取可能因批次饿死没跑到）。
    await withHostDbRetry(() =>
      PluginAPI.updateTask(task.id, {
        isDone: true,
        issueLastSyncedValues: withCompletedBaseline(baseline),
        ...(typeof fresh.lastUpdated === 'number' ? { issueLastUpdated: fresh.lastUpdated } : {}),
      }),
    );
  } else if (decision.action === 'baseline-only') {
    await writeBaseline(fresh.lastUpdated);
  }
  return decision;
}

// 同一任务的兜底推送去重（用户快速打勾/取消时只保留最后一次）。
const completionSafetyTimers = new Map<string, ReturnType<typeof setTimeout>>();
const COMPLETION_SAFETY_DELAY_MS = 4000;

/**
 * 打勾后的兜底推送：延迟 4 秒再检查一次基线。宿主的自动推送正常时（实测 ~1.6s）
 * 已把 issueLastSyncedValues.state 写成 completed，这里就什么都不做；只有宿主没推
 * 成功（或被 provider-changed 跳过）时才由插件补推。
 */
function scheduleCompletionSafetyNet(taskId: string, snapshot?: SpTask): void {
  const prev = completionSafetyTimers.get(taskId);
  if (prev) clearTimeout(prev);
  const timer = setTimeout(() => {
    completionSafetyTimers.delete(taskId);
    void (async () => {
      try {
        const all = await getAllTasks();
        const task = all.find((t) => t.id === taskId) ?? snapshot;
        if (!task || !task.isDone || !task.issueId) return;
        if (pushedCompletionIssues.has(task.issueId)) return;
        if (!shouldPushCompletionOnUpdate(task, { isDone: true })) return;
        const decision = await syncCompletionForTask(task);
        lifecycle(`completion safety-net ${taskId}: ${decision.action} (${decision.reason})`);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        lifecycle(`completion safety-net failed for ${taskId}: ${msg}`);
        showSnack(
          t({
            zh: `完成状态推送失败：${msg}`,
            en: `Failed to push completion: ${msg}`,
          }),
          'ERROR',
        );
      }
    })();
  }, COMPLETION_SAFETY_DELAY_MS);
  completionSafetyTimers.set(taskId, timer);
}

/**
 * 注册 taskUpdate hook。只处理「本插件任务被本地打勾」，并排除宿主拉取完成状态时
 * 派发的 updateTask（那种 changes 里带 issueLastSyncedValues / issueWasUpdated）。
 */
function registerCompletionSyncHook(): void {
  if (typeof PluginAPI.registerHook !== 'function') {
    lifecycle('registerHook unavailable, completion safety-net disabled');
    return;
  }
  try {
    PluginAPI.registerHook('taskUpdate', (payload) => {
      const task = payload?.task as SpTask | undefined;
      // 取消打勾：清掉兜底推送标记。否则同一任务“取消 → 再次打勾”后，
      // 进程内标记会让兜底通道永久跳过它（即使宿主这次推送失败）。
      if (task?.issueId && payload?.changes?.['isDone'] === false) {
        pushedCompletionIssues.delete(task.issueId);
        return;
      }
      if (!shouldPushCompletionOnUpdate(task, payload?.changes)) return;
      scheduleCompletionSafetyNet((task as SpTask).id, task as SpTask);
    });
    lifecycle('taskUpdate hook registered (completion safety-net)');
  } catch (e) {
    lifecycle('registerHook failed: ' + (e instanceof Error ? e.message : String(e)));
  }
}

/**
 * 视图「对账」：扫描所有已链接任务的完成状态分叉（本地 isDone vs 同步基线），
 * 逐条读远端并修复：
 *   - 本地完成、远端未完成 → 推 completed
 *   - 远端完成、本地未完成 → 本地打勾
 *   - 两边都完成 → 只回写基线
 * 返回逐条结果（含错误），供视图展示。
 */
async function reconcileCompletions(): Promise<RpcResult> {
  await resolveConfig();
  const all = await getAllTasks();
  const diverged = selectCompletionDivergences(all);
  const dueDiverged = selectDueDivergences(all);
  const results: Array<Record<string, unknown>> = [];
  for (const task of diverged) {
    try {
      const decision = await syncCompletionForTask(task);
      results.push({
        taskId: task.id,
        title: task.title,
        action: decision.action,
        reason: decision.reason,
        ok: true,
      });
    } catch (e) {
      results.push({
        taskId: task.id,
        title: task.title,
        action: 'error',
        ok: false,
        error: e instanceof Error ? e.message : String(e),
      });
    }
  }
  for (const task of dueDiverged) {
    try {
      const decision = await reconcileDueForTask(task);
      results.push({
        taskId: task.id,
        title: task.title,
        action: decision.action,
        reason: decision.reason,
        ok: true,
      });
    } catch (e) {
      results.push({
        taskId: task.id,
        title: task.title,
        action: 'error',
        ok: false,
        error: e instanceof Error ? e.message : String(e),
      });
    }
  }
  return {
    ok: true,
    checked: all.length,
    diverged: diverged.length,
    dueDiverged: dueDiverged.length,
    results,
  };
}

/**
 * 单个任务的日期对账：
 *   - 远端日期与同步基线一致（远端没变）→ 把本地 dueDay/dueWithTime 推上去；
 *   - 远端日期已变化 → 跳过（交给轮询拉取，避免覆盖远端新值）。
 * 只处理日期字段，不碰完成状态。
 */
async function reconcileDueForTask(
  task: SpTask,
): Promise<{ action: string; reason: string }> {
  const config = await resolveConfig();
  const http = resolveHttp();
  const issueId = task.issueId as string;
  const updateIssue = issueProviderDefinition.updateIssue;
  if (!updateIssue) throw new Error('updateIssue not implemented');

  const fresh = await issueProviderDefinition.getById(issueId, config, http);
  const lsv = (task.issueLastSyncedValues ?? {}) as Record<string, unknown>;
  const normDay = (v: unknown): string | null => (typeof v === 'string' && v ? v : null);
  const normTime = (v: unknown): number | null =>
    typeof v === 'number' && Number.isFinite(v) ? v : null;

  const freshDay = normDay(fresh.dueDate);
  const freshTime = normTime(fresh.dueWithTime);
  // 基线里的 dueDate 可能为空而远端有日期：定时任务（dueDateTime=当天 + 提醒）的
  // 日期由 dueWithTime 派生，基线是宿主按「本地只有 dueWithTime」写的。这种
  // 「时刻一致、日期多出来」不是远端改动，否则永远判 due-skip（2026-10-29 实测）。
  const sameInstant =
    freshTime !== null && freshTime === normTime(lsv['dueWithTime']);
  const dayChanged = !sameInstant && freshDay !== normDay(lsv['dueDate']);
  const timeChanged = freshTime !== normTime(lsv['dueWithTime']);
  if (dayChanged || timeChanged) {
    return { action: 'due-skip', reason: 'remote due changed; pull will reconcile' };
  }

  const localDay = normDay(task.dueDay);
  const localTime = normTime(task.dueWithTime);
  const baselineDay = normDay(lsv['dueDate']);
  const baselineTime = normTime(lsv['dueWithTime']);

  // 遗留缺口（0.1.10 之前把定时截止读成纯日期）：本地只剩日期，MS 有提醒时刻，
  // 且基线自洽。用户的约定是「时间以 MS 为准」→ 把远端时刻拉回本地，绝不反向清提醒。
  if (isLegacyTimedBaselineGap(localDay, localTime, baselineDay, baselineTime)) {
    const remoteTime = freshTime ?? baselineTime;
    await withHostDbRetry(() =>
      PluginAPI.updateTask(task.id, {
        dueWithTime: remoteTime,
        issueLastSyncedValues: {
          ...(task.issueLastSyncedValues ?? {}),
          dueDate: freshDay ?? baselineDay,
          dueWithTime: remoteTime,
        },
        ...(typeof fresh.lastUpdated === 'number' ? { issueLastUpdated: fresh.lastUpdated } : {}),
      }),
    );
    return { action: 'due-pull', reason: 'pulled remote time into local (legacy gap)' };
  }

  if (localTime !== null) {
    await updateIssue(issueId, { dueWithTime: localTime }, config, http);
  } else if (baselineTime !== null) {
    // 本地把「时刻」删掉了（定时 → 纯日期）：清远端提醒，同时写本地日期。
    await updateIssue(issueId, { dueWithTime: null, dueDate: localDay }, config, http);
  } else {
    // localDay 为 null 时表示本地已清空日期：显式写 null 清除远端。
    await updateIssue(issueId, { dueDate: localDay }, config, http);
  }

  const post = await issueProviderDefinition.getById(issueId, config, http);
  await withHostDbRetry(() =>
    PluginAPI.updateTask(task.id, {
      issueLastSyncedValues: {
        ...(task.issueLastSyncedValues ?? {}),
        dueDate: normDay(post.dueDate),
        dueWithTime: normTime(post.dueWithTime),
      },
      ...(typeof post.lastUpdated === 'number' ? { issueLastUpdated: post.lastUpdated } : {}),
    }),
  );
  return {
    action: 'due-push',
    reason: localTime !== null ? 'pushed timed due' : 'pushed planned day',
  };
}

/** 单任务补推：完成态 + 计划日期各查一次分叉，有分叉就补推（视图/RPC 用）。 */
async function pushTaskDivergence(taskId: string): Promise<RpcResult> {
  const config = await resolveConfig();
  const all = await getAllTasks();
  const task = all.find((t) => t.id === taskId);
  if (!task) return { ok: false, error: 'task not found: ' + taskId };
  const results: Array<Record<string, unknown>> = [];
  // 附上远端现值：排查「远端已改 → due-skip」时不用再猜。
  let remote: Record<string, unknown> | undefined;
  try {
    const fresh = await issueProviderDefinition.getById(
      task.issueId as string,
      config,
      resolveHttp(),
    );
    remote = {
      state: fresh.state,
      dueDate: fresh.dueDate,
      dueWithTime: fresh.dueWithTime,
      lastUpdated: fresh.lastUpdated,
    };
  } catch {
    remote = undefined;
  }
  if (selectCompletionDivergences([task]).length) {
    results.push({ kind: 'completion', ...(await syncCompletionForTask(task)) });
  }
  if (selectDueDivergences([task]).length) {
    results.push({ kind: 'due', ...(await reconcileDueForTask(task)) });
  }
  return {
    ok: true,
    taskId,
    title: task.title,
    local: { isDone: !!task.isDone, dueDay: task.dueDay, dueWithTime: task.dueWithTime },
    baseline: task.issueLastSyncedValues,
    remote,
    results,
  };
}

// --- 后台对账（跨设备兜底） ---
//
// 另一台设备（手机）上打勾 / 改计划日期时，本机只会在「拉取」方向看到新状态：
// 宿主的推送效果器只对本地任务更新动作触发，不会为同步进来的改动补推，于是 MS
// 永远停在旧状态（2026-10-28 23:25 的「找黑色裤」就是这样漏掉的）。
// 这里每 5 分钟扫一次分叉并补推 —— 任何一台装了插件的设备在线，都能把改动带到 MS。
const SWEEP_INTERVAL_MS = 5 * 60_000;
const SWEEP_FIRST_DELAY_MS = 120_000;
const SWEEP_MAX_PER_RUN = 20;
let sweepStarted = false;
let sweepInFlight = false;

async function runDivergenceSweep(reason: string): Promise<void> {
  if (sweepInFlight) return;
  sweepInFlight = true;
  try {
    const all = await getAllTasks();
    // 归档任务的基线可能写不回去（宿主对归档任务 updateTask 有限制），用进程内标记
    // 防止每轮重复推同一个完成态；取消打勾时 hook 会清掉标记。
    const completions = selectCompletionDivergences(all)
      .filter((task) => !pushedCompletionIssues.has(task.issueId as string))
      .slice(0, SWEEP_MAX_PER_RUN);
    // 自动补推只处理「今天或未来」的日期：历史逾期日期上的分叉多半是 0.1.10 之前
    // 遗留的 UTC/日期换算漂移（用户已约定 MS 为准），自动推会把错误的本地日期写回
    // MS。这类留给「对账」页人工确认。
    const today = localDayOf(Date.now());
    const dues = selectDueDivergences(all)
      .filter((task) => {
        const day =
          typeof task.dueDay === 'string' && task.dueDay
            ? task.dueDay
            : typeof task.dueWithTime === 'number'
              ? localDayOf(task.dueWithTime)
              : null;
        return day !== null && day >= today;
      })
      .slice(0, SWEEP_MAX_PER_RUN);
    if (!completions.length && !dues.length) return;
    let ok = 0;
    const failures: string[] = [];
    for (const task of completions) {
      try {
        const decision = await syncCompletionForTask(task);
        ok++;
        lifecycle(`sweep[${reason}] completion ${decision.action} · ${task.title}`);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        failures.push(`${task.id}: ${msg}`);
        lifecycle(`sweep[${reason}] completion failed ${task.id}: ${msg}`);
      }
    }
    for (const task of dues) {
      try {
        const decision = await reconcileDueForTask(task);
        ok++;
        lifecycle(`sweep[${reason}] due ${decision.action} · ${task.title}`);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        failures.push(`${task.id}: ${msg}`);
        lifecycle(`sweep[${reason}] due failed ${task.id}: ${msg}`);
      }
    }
    lifecycle(
      `sweep[${reason}] done: ${completions.length} completion + ${dues.length} due, ok=${ok}, fail=${failures.length}`,
    );
  } catch (e) {
    lifecycle('sweep failed: ' + (e instanceof Error ? e.message : String(e)));
  } finally {
    sweepInFlight = false;
  }
}

function startDivergenceSweep(): void {
  if (sweepStarted) return;
  sweepStarted = true;
  try {
    setTimeout(() => void runDivergenceSweep('startup'), SWEEP_FIRST_DELAY_MS);
    setInterval(() => void runDivergenceSweep('interval'), SWEEP_INTERVAL_MS);
    lifecycle('divergence sweep scheduled');
  } catch (e) {
    lifecycle('sweep scheduling failed: ' + (e instanceof Error ? e.message : String(e)));
  }
}

/** 给任务补上「所属项目名 / 标签名」等人类可读信息（SP 任务实体只有 id）。 */
function decorateTask(
  task: SpTask,
  projectById: Map<string, string>,
  tagById: Map<string, SpTag>,
): Record<string, unknown> {
  return {
    id: task.id,
    title: task.title,
    notes: task.notes,
    isDone: !!task.isDone,
    dueDay: task.dueDay,
    dueWithTime: task.dueWithTime,
    projectId: task.projectId,
    projectName: task.projectId ? projectById.get(task.projectId) : undefined,
    tags: (task.tagIds ?? [])
      .map((id) => tagById.get(id))
      .filter((t): t is SpTag => !!t)
      .map((t) => ({ title: t.title, color: t.color })),
  };
}

// status：视图载入时拉取任务快照（连接态 / 列表名 / 可推送 / 已链接 / 全列表总览，
// 任务已补可读字段，并反推出「项目↔列表」映射以修正当前列表）。
async function getSyncStatus(): Promise<RpcResult> {
  const connected = !!(await secrets.get(SECRET_REFRESH_TOKEN));

  const [projectId, all, projects, tags] = await Promise.all([
    currentProjectId(),
    getAllTasks(),
    safeGetAllProjects(),
    safeGetAllTags(),
  ]);
  const projectById = new Map(projects.map((p) => [p.id, p.title] as const));
  const tagById = new Map(tags.map((t) => [t.id, t] as const));
  const projectName = projectId ? projectById.get(projectId) : undefined;

  // 全列表总览（含「MS 任务 id → 列表 id」索引）。未连接或拉取失败时退化为空。
  let overview: ListsOverview | undefined;
  if (connected) {
    try {
      overview = await buildOverview();
    } catch {
      overview = undefined;
    }
  }
  const msTaskToList = overview?.msTaskToList ?? new Map<string, string>();
  const listNameById = new Map(
    (overview?.lists ?? []).map((l) => [l.id, l.displayName] as const),
  );

  const linkedTasks = selectLinkedMsTasks(all);
  // 反推项目→列表（一个项目一个 provider 实例 = 一个列表）。
  const projectToList = buildProjectToListMap(linkedTasks, msTaskToList);
  // 每个 MS 任务 id → 它在 SP 里的归属（用于「列表总览」标记已链接/未链接）。
  const spLinkedByIssueId = new Map<string, { taskId: string; projectName?: string }>();
  for (const task of linkedTasks) {
    spLinkedByIssueId.set(task.issueId as string, {
      taskId: task.id,
      projectName: task.projectId ? projectById.get(task.projectId) : undefined,
    });
  }

  // 兜底列表：cachedConfig（最后被宿主轮询到的实例）里的 listId —— 只在「项目↔列表」
  // 反推失败（例如项目里还没有任何已链接任务）时用，可能不准。
  let fallbackListId: string | undefined;
  try {
    if (cachedConfig) fallbackListId = parseConfig(cachedConfig).listId;
  } catch {
    fallbackListId = undefined;
  }

  // 当前列表 = 当前项目绑定的列表（反推）；反推不到才退回 cachedConfig。
  const currentListId = projectId
    ? (projectToList.get(projectId) ?? fallbackListId)
    : fallbackListId;
  const currentListName = currentListId ? listNameById.get(currentListId) : undefined;

  const scoped = projectId ? all.filter((task) => task.projectId === projectId) : all;
  const linked = linkedTasks
    .filter((task) => (projectId ? task.projectId === projectId : true))
    .map((task) => {
      const listId = msTaskToList.get(task.issueId as string);
      return {
        ...decorateTask(task, projectById, tagById),
        issueId: task.issueId as string,
        listId: listId ?? null,
        listName: listId ? (listNameById.get(listId) ?? null) : null,
      };
    });
  const pushable = scoped.filter(isPushableTask).map((task) => {
    const listId = projectToList.get(task.projectId as string) ?? fallbackListId;
    return {
      ...decorateTask(task, projectById, tagById),
      listId: listId ?? null,
      listName: listId ? (listNameById.get(listId) ?? null) : null,
    };
  });

  // 列表总览：每个列表的任务 + 是否已链接到 SP（以及归属项目名）。
  const lists = (overview?.lists ?? []).map((l) => {
    const tasks = l.tasks.map((t) => {
      const linkedInfo = spLinkedByIssueId.get(t.id);
      return {
        id: t.id,
        title: t.title,
        status: t.status,
        linkedTaskId: linkedInfo?.taskId ?? null,
        linkedProjectName: linkedInfo?.projectName ?? null,
      };
    });
    const linkedCount = tasks.filter((t) => t.linkedTaskId).length;
    return {
      id: l.id,
      displayName: l.displayName,
      taskCount: l.taskCount,
      linkedCount,
      tasks,
    };
  });

  return {
    ok: true,
    version: PLUGIN_VERSION,
    connected,
    listId: currentListId ?? null,
    listName: currentListName ?? null,
    projectId: projectId ?? null,
    projectName: projectName ?? null,
    totalTasks: all.length,
    linked,
    pushable,
    lists,
  };
}

// forceSync：把一个或多个 task 推送到目标 To Do 列表并回写链接（支持批量、可逐任务指定目标列表）。
// 目标列表按任务逐个解析：targets[taskId] 显式指定 > 任务所属项目绑定列表 > cachedConfig 兜底。
// 配置 / 令牌 / 任务表只解析一次；每个任务独立 POST 与回写，逐个报告结果。
async function forceSyncMany(
  taskIds: string[],
  targets?: Record<string, string>,
): Promise<RpcResult> {
  const ids = taskIds.filter(Boolean);
  if (ids.length === 0) {
    return { ok: false, error: t({ zh: '没有要推送的任务。', en: 'No tasks to push.' }) };
  }

  const config = await resolveConfig();
  const clientId = getClientId(config, CLIENT_ID);
  const http = resolveHttp();
  const token = await getAccessToken(clientId, http, secrets);

  const all = await getAllTasks();
  const byId = new Map(all.map((task) => [task.id, task] as const));

  // 项目→列表映射（复用 status 的总览缓存；缓存未命中则重建）。拉不到就退回兜底列表。
  let projectToList = new Map<string, string>();
  let fallbackListId: string | undefined;
  try {
    fallbackListId = parseConfig(config).listId;
  } catch {
    fallbackListId = undefined;
  }
  try {
    const overview = await buildOverview();
    projectToList = buildProjectToListMap(
      selectLinkedMsTasks(all),
      overview.msTaskToList,
    );
  } catch {
    // 忽略：无总览时 projectToList 为空，每个任务走兜底列表或显式指定列表。
  }

  interface PushResult {
    id: string;
    ok: boolean;
    title?: string;
    issueId?: string;
    error?: string;
    warning?: string;
  }
  const results: PushResult[] = [];
  let pushed = 0;
  let failed = 0;

  for (const id of ids) {
    if (inflightSync.has(id)) {
      results.push({
        id,
        ok: false,
        error: t({ zh: '该任务正在推送中。', en: 'Already being pushed.' }),
      });
      failed++;
      continue;
    }
    inflightSync.add(id);
    try {
      const task = byId.get(id);
      if (!task) {
        results.push({
          id,
          ok: false,
          error: t({ zh: '未找到该任务（可能已被删除）。', en: 'Task not found.' }),
        });
        failed++;
        continue;
      }
      if (task.issueId) {
        results.push({
          id,
          ok: false,
          title: task.title,
          error: t({ zh: '该任务已同步（已有 MS 链接）。', en: 'Already synced.' }),
        });
        failed++;
        continue;
      }
      // 这里 task 已非空、且无 issueId；只需排除「属于其它 provider」的任务。
      // 注意不能用 isPushableTask(task) 的类型谓词在 if 里取反（会收窄成 never）。
      if (task.issueType && task.issueType !== PLUGIN_ISSUE_TYPE) {
        results.push({
          id,
          ok: false,
          title: task.title,
          error: t({
            zh: '该任务属于其它 provider，不能推送到 Microsoft To Do。',
            en: 'This task belongs to another provider.',
          }),
        });
        failed++;
        continue;
      }

      const boundListId = projectToList.get(task.projectId as string);
      const explicit = targets?.[id];
      const listId = resolvePushListId(task, explicit, projectToList, fallbackListId);
      if (!listId) {
        results.push({
          id,
          ok: false,
          title: task.title,
          error: t({
            zh: '无法确定目标列表：该任务所属项目没有绑定任何 To Do 列表，请在推送下拉里手动选择。',
            en: "Cannot determine the target list: this task's project has no bound To Do list. Pick one in the push dropdown.",
          }),
        });
        failed++;
        continue;
      }

      const created = await graphRequest<TodoTask>(
        'POST',
        `/me/todo/lists/${listId}/tasks`,
        token,
        buildCreateTaskBody(task),
        http,
      );
      const issueData = mapTaskToIssue(created);
      // 回写链接：issueType 固定为本插件；issueProviderId 优先用任务自带的，否则从同项目
      // 已链接任务里抄（自动创建失败的任务两者都缺，缺了它宿主无法路由后续双向同步）。
      const issueProviderId =
        task.issueProviderId ?? findIssueProviderId(all, task.projectId);
      const link: Record<string, unknown> = {
        issueId: created.id,
        issueType: PLUGIN_ISSUE_TYPE,
        issueData,
        issueWasUpdated: false,
        issueLastUpdated: Date.now(),
        issueLastSyncedValues: extractSyncValues(issueData),
      };
      if (issueProviderId) link.issueProviderId = issueProviderId;
      await PluginAPI.updateTask(task.id, link);

      const result: PushResult = { id, ok: true, title: created.title, issueId: created.id };
      if (!issueProviderId) {
        result.warning = t({
          zh:
            '已推送，但未能找到本项目的 provider 实例 id（项目里还没有任何已链接任务）。' +
            '请先正常同步一次，让项目里至少有一个已链接任务，再重试。',
          en:
            'Pushed, but no provider instance id could be found (no linked tasks in this project yet). ' +
            'Run a normal sync once so the project has at least one linked task, then retry.',
        });
      }
      if (explicit && boundListId && explicit !== boundListId) {
        const warn = t({
          zh:
            '已推送到所选列表，但该任务所属项目绑定的列表不同，后续双向同步可能因列表不一致而 404。',
          en:
            "Pushed to the chosen list, but it differs from the list bound to this task's project; later two-way sync may 404.",
        });
        result.warning = result.warning ? result.warning + ' ' + warn : warn;
      }
      results.push(result);
      pushed++;
    } catch (e) {
      results.push({
        id,
        ok: false,
        error: e instanceof Error ? e.message : String(e),
      });
      failed++;
    } finally {
      inflightSync.delete(id);
    }
  }

  return { ok: failed === 0, pushed, failed, results };
}

// scanDeadLinks：链接体检。用全列表总览（不限 maxTasks）给每个已链接任务定位真实列表，
// 区分「失效链接」（远端已删）与「列表不一致」（被挪到别的列表/推错列表）。
async function scanDeadLinks(): Promise<RpcResult> {
  if (isScanning) {
    return {
      ok: false,
      error: t({ zh: '正在检测中，请稍候。', en: 'A scan is already running.' }),
    };
  }
  isScanning = true;
  try {
    const config = await resolveConfig();
    const clientId = getClientId(config, CLIENT_ID);
    const http = resolveHttp();
    const token = await getAccessToken(clientId, http, secrets);

    const projectId = await currentProjectId();
    const all = await getAllTasks();
    let linked = selectLinkedMsTasks(all);
    if (projectId) {
      linked = linked.filter((task) => task.projectId === projectId);
    }
    if (linked.length === 0) {
      return { ok: true, checked: 0, unverifiable: 0, dead: [], mismatched: [], empty: true };
    }

    // 不限量拉全量总览，保证「issueId 找不到」等价于「远端已删除」，而不是被 maxTasks 截断。
    const overview = await buildOverview(0);
    const listNameById = new Map(
      overview.lists.map((l) => [l.id, l.displayName] as const),
    );

    const health = classifyLinkedTasks(linked, overview.msTaskToList);
    const dead = health.dead.map((d) => ({
      taskId: d.taskId,
      title: d.title,
      issueId: d.issueId,
    }));
    const mismatched = health.mismatched.map((m) => ({
      taskId: m.taskId,
      title: m.title,
      issueId: m.issueId,
      listId: m.listId,
      expectedListId: m.expectedListId,
      listName: listNameById.get(m.listId) ?? m.listId,
      expectedListName: listNameById.get(m.expectedListId) ?? m.expectedListId,
    }));

    return { ok: true, checked: health.checked, unverifiable: 0, dead, mismatched };
  } catch (e) {
    return toErrorResult(e);
  } finally {
    isScanning = false;
  }
}

// unlink：解除本地任务的 MS 链接。SP 插件 API 没有「删除任务」接口（只有 updateTask），
// 所以对「失效链接 / 列表不一致」最安全的补救就是把 issue* 字段清空，让任务变回普通本地
// 任务（回到「推送」列表可重新推送），宿主不再拿旧 issueId 去 Graph 轮询、也就停止 404。
async function unlinkTask(taskId: string): Promise<RpcResult> {
  if (!taskId) {
    return { ok: false, error: t({ zh: '缺少任务 id。', en: 'Missing task id.' }) };
  }
  const all = await getAllTasks();
  const task = all.find((candidate) => candidate.id === taskId);
  if (!task) {
    return {
      ok: false,
      error: t({ zh: '未找到该本地任务（可能已被删除）。', en: 'Local task not found.' }),
    };
  }
  // 注意：宿主 updateTask 的 typia 校验只接受 `string | undefined`，传 null 会被拒
  // （报 "invalid type on $input.issueId"）。用 undefined 才是跨宿主版本安全的“清空”。
  await PluginAPI.updateTask(taskId, {
    issueId: undefined,
    issueType: undefined,
    issueProviderId: undefined,
    issueData: undefined,
    issueWasUpdated: false,
    issueLastUpdated: undefined,
    issueLastSyncedValues: undefined,
  });
  return { ok: true, title: task.title };
}

// recreateLink：把「列表不一致」的任务在应属列表里重建一份、重新链接，并删除旧列表里的
// 原任务（等价于「移回正确列表」）。这是破坏性操作：会改变本地任务的 issueId、并删除远端
// 旧任务，因此只在用户显式二次确认后由视图调用。重建依赖「应属列表 = 该项目已链接任务的
// 众数列表」；单链接项目或整项目被整体挪走时无法判定，返回错误让用户手动处理。
async function recreateLink(
  taskId: string,
  expectedListIdOverride?: string,
): Promise<RpcResult> {
  if (!taskId) {
    return { ok: false, error: t({ zh: '缺少任务 id。', en: 'Missing task id.' }) };
  }

  const config = await resolveConfig();
  const clientId = getClientId(config, CLIENT_ID);
  const http = resolveHttp();
  const token = await getAccessToken(clientId, http, secrets);

  const all = await getAllTasks();
  const task = all.find((candidate) => candidate.id === taskId);
  if (!task) {
    return {
      ok: false,
      error: t({ zh: '未找到该本地任务（可能已被删除）。', en: 'Local task not found.' }),
    };
  }
  if (!isMsTodoTask(task)) {
    return {
      ok: false,
      error: t({ zh: '该任务不属于 Microsoft To Do。', en: 'Not a Microsoft To Do task.' }),
    };
  }
  const issueId = task.issueId as string;
  if (!issueId) {
    return {
      ok: false,
      error: t({ zh: '该任务没有 MS 链接（无需重建）。', en: 'This task is not linked to MS.' }),
    };
  }

  // 全量总览定位真实列表，并复用体检的众数启发式确定「应属列表」。
  const overview = await buildOverview(0);
  const actualListId = overview.msTaskToList.get(issueId);
  if (!actualListId) {
    return {
      ok: false,
      error: t({
        zh: '远端已找不到该任务（已失效），请改用「解除链接」。',
        en: 'This task no longer exists in To Do; use Unlink instead.',
      }),
    };
  }
  const health = classifyLinkedTasks(selectLinkedMsTasks(all), overview.msTaskToList);
  const hit = health.mismatched.find((m) => m.taskId === taskId);
  // 显式指定优先：项目里只有一个链接（或整体被挪走）时众数无法判定，但调用方
  // （一次性治理脚本 / 高级用户）知道应属列表时可以直接传入。
  const expectedListId = expectedListIdOverride ?? hit?.expectedListId;
  if (!expectedListId || expectedListId === actualListId) {
    return {
      ok: false,
      error: t({
        zh: '无法确定应属列表（该项目可能只有一个链接，或整项目已被整体挪走）。请手动处理。',
        en: 'Cannot determine the target list (single-link project, or the whole project moved). Handle manually.',
      }),
    };
  }

  // 1) 在应属列表重建同内容任务；2) 回写链接指向新任务；3) 删除旧列表原任务。
  const created = await graphRequest<TodoTask>(
    'POST',
    `/me/todo/lists/${expectedListId}/tasks`,
    token,
    buildCreateTaskBody(task),
    http,
  );
  const issueData = mapTaskToIssue(created);
  const issueProviderId = task.issueProviderId ?? findIssueProviderId(all, task.projectId);
  const link: Record<string, unknown> = {
    issueId: created.id,
    issueType: PLUGIN_ISSUE_TYPE,
    issueData,
    issueWasUpdated: false,
    issueLastUpdated: Date.now(),
    issueLastSyncedValues: extractSyncValues(issueData),
  };
  if (issueProviderId) link.issueProviderId = issueProviderId;
  await PluginAPI.updateTask(task.id, link);

  let warning: string | undefined;
  try {
    await graphRequest(
      'DELETE',
      `/me/todo/lists/${actualListId}/tasks/${issueId}`,
      token,
      undefined,
      http,
    );
  } catch (e) {
    warning =
      t({
        zh: '新任务已创建并链接，但删除旧任务失败（可稍后在 To Do 里手动删除）。',
        en: 'New task created and linked, but deleting the old one failed (delete it manually later).',
      }) + ' ' + (e instanceof Error ? e.message : String(e));
  }

  const result: RpcResult = { ok: true, title: created.title, newIssueId: created.id };
  if (warning) result.warning = warning;
  return result;
}

// deleteMsTask：删除 To Do 远端任务（「列表」页对未链接任务的操作，破坏性，视图侧已二次确认）。
// 直接按任务所在列表 DELETE；404 时复用 findTaskList 定位真实列表再删（任务被挪走也能删对），
// 全部 404 视为已删除（幂等）。成功后清掉总览缓存，否则 60s 内 loadStatus 仍会显示已删任务。
async function deleteMsTask(listId: string, msTaskId: string): Promise<RpcResult> {
  if (!listId || !msTaskId) {
    return { ok: false, error: t({ zh: '缺少列表或任务 id。', en: 'Missing list/task id.' }) };
  }
  const config = await resolveConfig();
  const clientId = getClientId(config, CLIENT_ID);
  const http = resolveHttp();
  const token = await getAccessToken(clientId, http, secrets);
  try {
    await graphRequest('DELETE', `/me/todo/lists/${listId}/tasks/${msTaskId}`, token, undefined, http);
  } catch (e) {
    if (!isNotFoundError(e)) throw e;
    const actual = await findTaskList(msTaskId, token, http);
    if (!actual) return { ok: true };
    await graphRequest('DELETE', `/me/todo/lists/${actual}/tasks/${msTaskId}`, token, undefined, http);
  }
  overviewCache = undefined;
  return { ok: true };
}

/** 把视图传来的 `targets`（taskId → 列表 id 映射）清洗为纯 Record；仅保留非空字符串值。 */
function parseTargets(raw: unknown): Record<string, string> | undefined {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return undefined;
  const out: Record<string, string> = {};
  let any = false;
  for (const [k, v] of Object.entries(raw as Record<string, unknown>)) {
    if (typeof v === 'string' && v) {
      out[k] = v;
      any = true;
    }
  }
  return any ? out : undefined;
}

/** 顶部按钮 / 同步动作的结果反馈（不抛错，snack 自身失败仅忽略）。 */
function showSnack(msg: string, type: 'SUCCESS' | 'ERROR' | 'WARNING' | 'INFO'): void {
  try {
    PluginAPI.showSnack({ msg, type, ico: type === 'ERROR' ? 'error' : undefined });
  } catch {
    // 忽略 snack 失败。
  }
}

// 手动同步的去重锁：避免用户连点顶部按钮时并发触发多个 forceSyncMany。
let isSyncNowRunning = false;

// 顶部按钮「手动触发同步」：把当前范围内所有未同步（可推送）的本插件任务一次性
// 推到 Microsoft To Do。范围与独立视图一致 —— 处于项目上下文时只推该项目，否则全部。
// 目标列表沿用 forceSyncMany 的解析（显式 > 项目绑定 > 兜底列表）。
// 背景：桌面端自动创建可能因宿主 CORS 预检失败为 status=0，这里提供一键手动补救；
// MS → SP 的拉取方向仍由宿主按 pollIntervalMs 自动完成（插件无触发入口）。
async function syncNow(): Promise<void> {
  if (isSyncNowRunning) {
    showSnack(t({ zh: '正在同步中，请稍候…', en: 'Sync already in progress…' }), 'INFO');
    return;
  }
  isSyncNowRunning = true;
  try {
    const projectId = await currentProjectId();
    const all = await getAllTasks();
    const ids = selectTasksToSync(all, projectId).map((task) => task.id);
    if (ids.length === 0) {
      showSnack(
        t({
          zh: '没有待同步的任务（无 MS 链接的本地任务会出现在这里）。',
          en: 'No tasks to sync (local tasks without an MS link show up here).',
        }),
        'INFO',
      );
      return;
    }
    // 头部按钮无法像独立视图那样滚动打印逐任务进度，故用「开始/结束」两条 snackbar
    // 给出可感知反馈；详细进度见左侧菜单的独立视图。
    showSnack(
      t({
        zh: `开始同步 ${ids.length} 个任务到 Microsoft To Do…`,
        en: `Syncing ${ids.length} task(s) to Microsoft To Do…`,
      }),
      'INFO',
    );
    const result = await forceSyncMany(ids);
    const pushed = Number(result.pushed ?? 0);
    const failed = Number(result.failed ?? 0);
    if (failed === 0) {
      showSnack(
        t({
          zh: `同步完成：已推送 ${pushed} 个任务。`,
          en: `Sync complete: pushed ${pushed} task(s).`,
        }),
        'SUCCESS',
      );
    } else if (pushed > 0) {
      showSnack(
        t({
          zh: `同步部分完成：成功 ${pushed}，失败 ${failed}。`,
          en: `Sync partially complete: ${pushed} ok, ${failed} failed.`,
        }),
        'WARNING',
      );
    } else {
      const results = (result.results ?? []) as Array<{ error?: string }>;
      const firstErr = results.find((r) => r.error)?.error;
      showSnack(
        firstErr
          ? t({
              zh: `同步失败：${failed} 个任务未推送。${firstErr}`,
              en: `Sync failed: ${failed} task(s) not pushed. ${firstErr}`,
            })
          : t({
              zh: `同步失败：${failed} 个任务未推送。`,
              en: `Sync failed: ${failed} task(s) not pushed.`,
            }),
        'ERROR',
      );
    }
  } catch (e) {
    showSnack(e instanceof Error ? e.message : String(e), 'ERROR');
  } finally {
    isSyncNowRunning = false;
  }
}

function registerViewEntry(): void {
  const open = (): void => {
    try {
      PluginAPI.showIndexHtmlAsView();
    } catch (e) {
      lifecycle('showIndexHtmlAsView failed: ' + (e instanceof Error ? e.message : String(e)));
      showSnack(e instanceof Error ? e.message : String(e), 'ERROR');
    }
  };
  const label = t({ zh: 'Microsoft To Do', en: 'Microsoft To Do' });
  // check_box（方框打勾）与插件 icon.svg 的视觉一致；check_circle 是圆形，不搭。
  // 顶部按钮用 check_box 而非 sync 图标：sync 会与宿主自己的「同步」按钮高度相似、难以区分。
  // 桌面端右上角曾有两个重复入口：sidePanelButton 由宿主先 dispatch
  // togglePluginPanel 再回调 onClick（无法干净复用为纯动作），且在移动端只会渲染成
  // 空白面板（宿主 #8394），故移除。现在：
  //   - 顶部按钮（registerHeaderButton）＝手动触发同步（把当前范围内未同步任务推到 MS）；
  //   - 左侧菜单（registerMenuEntry）＝打开独立面板（承载推送/体检/清理等完整操作）。
  // 两者彼此独立注册：任何一个失败都不影响另一个。
  const safeRegister = (name: string, fn: () => void): void => {
    try {
      fn();
      lifecycle(name + ' ok');
    } catch (e) {
      lifecycle(name + ' failed: ' + (e instanceof Error ? e.message : String(e)));
    }
  };
  safeRegister('registerHeaderButton', () =>
    PluginAPI.registerHeaderButton({
      label: t({ zh: '同步 Microsoft To Do', en: 'Sync Microsoft To Do' }),
      icon: 'check_box',
      onClick: () => {
        void syncNow();
      },
    }),
  );
  safeRegister('registerMenuEntry', () =>
    PluginAPI.registerMenuEntry({ label, icon: 'check_box', onClick: open }),
  );
}

const inflightSync = new Set<string>();
let isScanning = false;

// 视图（index.html）经宿主消息桥调用这里；所有 op 都返回纯 JSON，供视图渲染进度。
if (typeof PluginAPI.onMessage === 'function') {
  try {
    PluginAPI.onMessage(async (message) => {
      const msg = (message ?? {}) as {
        op?: string;
        taskIds?: unknown;
        targets?: unknown;
        taskId?: unknown;
        listId?: unknown;
      };
      try {
        switch (msg.op) {
          case 'status':
            return await getSyncStatus();
          case 'forceSync':
            return await forceSyncMany(
              Array.isArray(msg.taskIds) ? msg.taskIds.map((id) => String(id)) : [],
              parseTargets(msg.targets),
            );
          case 'scanDeadLinks':
            return await scanDeadLinks();
          case 'reconcile':
            return await reconcileCompletions();
          case 'pushTask':
            return await pushTaskDivergence(String(msg.taskId ?? ''));
          case 'unlink':
            return await unlinkTask(String(msg.taskId ?? ''));
          case 'recreateLink':
            return await recreateLink(
              String(msg.taskId ?? ''),
              typeof msg.listId === 'string' && msg.listId ? msg.listId : undefined,
            );
          case 'deleteMsTask':
            return await deleteMsTask(String(msg.listId ?? ''), String(msg.taskId ?? ''));
          default:
            return { ok: false, error: t({ zh: '未知操作。', en: 'Unknown operation.' }) };
        }
      } catch (e) {
        return toErrorResult(e);
      }
    });
    lifecycle('onMessage registered');
  } catch (e) {
    lifecycle('onMessage failed: ' + (e instanceof Error ? e.message : String(e)));
  }
} else {
  lifecycle('onMessage unavailable');
}

// 对齐 vision-board-super-productivity（移动端/桌面端都正常）：把视图入口注册放进
// onReady，由宿主在插件加载完成后统一触发，避免顶层注册在部分平台上过早丢失。
// 若宿主版本没有 onReady，则退化为顶层直接注册（旧版本兜底）。
if (typeof PluginAPI.onReady === 'function') {
  lifecycle('onReady available, deferring view entry registration');
  PluginAPI.onReady(() => {
    lifecycle('onReady fired');
    registerCompletionSyncHook();
    registerViewEntry();
    startDivergenceSweep();
  });
} else {
  lifecycle('onReady unavailable, registering view entry directly');
  registerCompletionSyncHook();
  registerViewEntry();
  startDivergenceSweep();
}

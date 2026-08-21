// Microsoft To Do issue provider for Super Productivity（桌面端 + 移动端 + Web）。
//
// 通过 SP 的 issueProvider 一等接口实现，双向同步由宿主 two-way-sync 引擎驱动；
// 本插件只需：OAuth 配置字段、Graph 读写、字段映射。
//
// 数据路径：Graph 请求经 graph.ts 全程走宿主 PluginHttp（`http` 权限，普通浏览器
// 请求，桌面 / 移动 / Web 三端一致，Graph 返回 CORS）；令牌经
// PluginAPI.getOAuthToken()（宿主自动 PKCE + refresh）。纯逻辑（配置解析、映射、
// 字段映射）在 domain.ts，便于单元测试。

import type {
  IssueProviderPluginDefinition,
  PluginHttp,
  PluginIssue,
  PluginSearchResult,
} from './types';
import {
  buildDueDateTime,
  extractSyncValues,
  FIELD_MAPPINGS,
  mapTaskToIssue,
  mapTaskToSearchResult,
  parseConfig,
  TODO_WEB_URL,
} from './domain';
import type { TodoTask } from './domain';
import { graphRequest, graphRequestAll, requireToken } from './graph';
import { t } from './i18n';

declare const PluginAPI: {
  registerIssueProvider(definition: IssueProviderPluginDefinition): void;
  log: {
    err: (...args: unknown[]) => void;
    log: (...args: unknown[]) => void;
    warn: (...args: unknown[]) => void;
    info: (...args: unknown[]) => void;
    debug: (...args: unknown[]) => void;
  };
};

// --- OAuth / Graph 常量 ---

const MS_AUTH_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
const MS_TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
const SCOPES = ['Tasks.ReadWrite', 'offline_access'];

// SP 的 Electron loopback 强制 127.0.0.1（不能 localhost，避开 ::1），端口 [1024,65535]。
// 固定端口以便在 Entra 应用注册里精确登记该 redirect URI。
const REDIRECT_URI = 'http://127.0.0.1:53682';

// 供 esbuild `define` 在构建期注入环境变量（编译期替换，非运行时读取）。
declare const process: { env: Record<string, string | undefined> };

// 默认 clientId：构建时从环境变量 MS_TODO_CLIENT_ID 注入；未设置则为空串。
// 运行时用户可在插件配置里用「Client ID」字段覆盖（写入 pluginConfig.oauthOverrides）。
const CLIENT_ID = process.env.MS_TODO_CLIENT_ID || '';

// --- Graph 数据类型 ---

interface TodoList {
  id: string;
  displayName: string;
  wellknownListName?: string;
}

// 加载 To Do 列表（登录后由宿主自动调用，用于 listId 下拉）。
async function loadLists(
  _config: Record<string, unknown>,
  http: PluginHttp,
): Promise<{ label: string; value: string }[]> {
  const token = await requireToken();
  const lists = await graphRequestAll<TodoList>('/me/todo/lists', token, 0, http);
  return lists.map((l) => ({ label: l.displayName, value: l.id }));
}

// --- 插件注册 ---

PluginAPI.registerIssueProvider({
  configFields: [
    {
      key: 'oauthOverrides.clientId',
      type: 'input',
      label: t({ zh: '客户端 ID', en: 'Client ID' }),
      description: t({
        zh:
          '你的 Entra ID 应用「应用程序(客户端) ID」（public client，非机密）。' +
          '构建时若未通过 MS_TODO_CLIENT_ID 环境变量注入，请在此填写，否则无法连接。',
        en:
          'Your Entra ID application (client) ID — a public client, not a secret. ' +
          'Fill this in if it was not injected at build time via the MS_TODO_CLIENT_ID env var, ' +
          'otherwise you cannot connect.',
      }),
      required: true,
    },
    {
      key: 'oauth',
      type: 'oauthButton',
      label: t({ zh: '连接 Microsoft 账户', en: 'Connect Microsoft account' }),
      description: t({
        zh:
          '登录 Microsoft 并授权读写 To Do 任务。需先在 Entra 应用清单登记重定向 URI ' +
          '（replyUrlsWithType，type 用 InstalledClient）并把 allowPublicClient 设为 true。' +
          '桌面端：http://127.0.0.1:53682；Android：com.superproductivity.superproductivity:/plugin-oauth-callback；' +
          'iOS：com.super-productivity.app:/plugin-oauth-callback。',
        en:
          'Sign in to Microsoft and authorize read/write access to your To Do tasks. ' +
          'First register these redirect URIs in the Entra app manifest (replyUrlsWithType, ' +
          'type InstalledClient) and set allowPublicClient to true. ' +
          'Desktop: http://127.0.0.1:53682; Android: com.superproductivity.superproductivity:/plugin-oauth-callback; ' +
          'iOS: com.super-productivity.app:/plugin-oauth-callback.',
      }),
      oauthConfig: {
        authUrl: MS_AUTH_URL,
        tokenUrl: MS_TOKEN_URL,
        clientId: CLIENT_ID,
        scopes: SCOPES,
        redirectUri: REDIRECT_URI,
      },
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
    const { listId, maxTasks } = parseConfig(config);
    const token = await requireToken();
    const tasks = await graphRequestAll<TodoTask>(
      `/me/todo/lists/${listId}/tasks`,
      token,
      maxTasks,
      http,
    );
    const term = searchTerm.trim().toLowerCase();
    return tasks
      .filter((t) => !term || t.title.toLowerCase().includes(term))
      .map(mapTaskToSearchResult);
  },

  async getNewIssuesForBacklog(
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginSearchResult[]> {
    const { listId, maxTasks } = parseConfig(config);
    const token = await requireToken();
    const tasks = await graphRequestAll<TodoTask>(
      `/me/todo/lists/${listId}/tasks`,
      token,
      maxTasks,
      http,
    );
    return tasks.map(mapTaskToSearchResult);
  },

  async getById(
    issueId: string,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginIssue> {
    const { listId } = parseConfig(config);
    const token = await requireToken();
    const task = await graphRequest<TodoTask>(
      'GET',
      `/me/todo/lists/${listId}/tasks/${issueId}`,
      token,
      undefined,
      http,
    );
    return mapTaskToIssue(task);
  },

  getIssueLink(_issueId: string, _config: Record<string, unknown>): string {
    // To Do 任务无稳定公开深链，链接到 Web 版应用首页。
    return TODO_WEB_URL;
  },

  async testConnection(
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<boolean> {
    try {
      const token = await requireToken();
      await graphRequest('GET', '/me/todo/lists', token, undefined, http);
      return true;
    } catch {
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
      field: 'dueDate',
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

  async updateIssue(
    id: string,
    changes: Record<string, unknown>,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<void> {
    const { listId } = parseConfig(config);
    const token = await requireToken();

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
    if (changes.dueDate !== undefined) {
      patch.dueDateTime = buildDueDateTime(changes.dueDate as string | null);
    }

    if (Object.keys(patch).length === 0) return;
    await graphRequest('PATCH', `/me/todo/lists/${listId}/tasks/${id}`, token, patch, http);
  },

  async createIssue(
    title: string,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<{ issueId: string; issueData: PluginIssue }> {
    const { listId } = parseConfig(config);
    const token = await requireToken();
    const task = await graphRequest<TodoTask>(
      'POST',
      `/me/todo/lists/${listId}/tasks`,
      token,
      { title },
      http,
    );
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
    const { listId } = parseConfig(config);
    const token = await requireToken();
    await graphRequest('DELETE', `/me/todo/lists/${listId}/tasks/${id}`, token, undefined, http);
  },

  extractSyncValues,
});

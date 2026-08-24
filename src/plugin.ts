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
} from './types';
import {
  buildDueDateTime,
  extractSyncValues,
  FIELD_MAPPINGS,
  getClientId,
  mapTaskToIssue,
  mapTaskToSearchResult,
  parseConfig,
  TODO_WEB_URL,
} from './domain';
import type { TodoTask } from './domain';
import { graphRequest, graphRequestAll } from './graph';
import {
  createDeviceCodeDialog,
  createSecretStore,
  getAccessToken,
  runDeviceCodeLogin,
  SECRET_REFRESH_TOKEN,
} from './auth';
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

// 供 esbuild `define` 在构建期注入环境变量（编译期替换，非运行时读取）。
declare const process: { env: Record<string, string | undefined> };

// 默认 clientId：构建时从环境变量 MS_TODO_CLIENT_ID 注入；未设置则为空串。
// 运行时用户可在插件配置里用「Client ID」字段覆盖。
const CLIENT_ID = process.env.MS_TODO_CLIENT_ID || '';

// 密钥存储（本地、每设备、不进 sync/备份），模块级单例即可。
const secrets = createSecretStore();

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

// 加载 To Do 列表（登录后由宿主自动调用，用于 listId 下拉）。
// 未配置 / 未连接 / 网络失败时返回空数组而非抛错 —— 与 SP 官方 Trello 插件的
// loadBoards 保持一致：loadOptions 抛错会让宿主把下拉置为 error 态，某些版本还会
// 反复重试/重渲染，导致配置界面卡死。
async function loadLists(
  config: Record<string, unknown>,
  http: PluginHttp,
): Promise<{ label: string; value: string }[]> {
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

PluginAPI.registerIssueProvider({
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
    try {
      const { listId, maxTasks } = parseConfig(config);
      const token = await getToken(config, http);
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
    } catch {
      return [];
    }
  },

  async getNewIssuesForBacklog(
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginSearchResult[]> {
    try {
      const { listId, maxTasks } = parseConfig(config);
      const token = await getToken(config, http);
      const tasks = await graphRequestAll<TodoTask>(
        `/me/todo/lists/${listId}/tasks`,
        token,
        maxTasks,
        http,
      );
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
    const { listId } = parseConfig(config);
    const token = await getToken(config, http);
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

  // 「测试连接」按钮复用为登录入口：未登录时先走设备码流，已登录则校验连接。
  async testConnection(
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<boolean> {
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
    const token = await getToken(config, http);
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
    const token = await getToken(config, http);
    await graphRequest('DELETE', `/me/todo/lists/${listId}/tasks/${id}`, token, undefined, http);
  },

  extractSyncValues,
});

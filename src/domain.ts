// 纯函数与常量（无副作用、不依赖宿主 PluginAPI 的运行时对象）。
//
// 从 plugin.ts 拆出，便于单元测试：配置解析、日期转换、任务↔issue 映射、
// 双向同步字段映射、extractSyncValues。plugin.ts 只保留 OAuth/Graph 读写与
// registerIssueProvider 的副作用组装。

import type {
  PluginFieldMapping,
  PluginIssue,
  PluginSearchResult,
} from './types';
import { t } from './i18n';

// To Do 任务无稳定公开深链，链接到 Web 版应用首页。
export const TODO_WEB_URL = 'https://to-do.office.com/tasks/';

// --- Graph 数据类型 ---

export interface TodoTask {
  id: string;
  title: string;
  status: 'notStarted' | 'inProgress' | 'completed' | 'waitingOnOthers' | 'deferred';
  importance?: 'low' | 'normal' | 'high';
  body?: { content: string; contentType: 'text' | 'html' };
  dueDateTime?: { dateTime: string; timeZone: string } | null;
  lastModifiedDateTime?: string;
}

// --- 配置 ---

interface MsTodoConfig {
  listId?: string;
  maxTasks?: string;
}

// 未配置时的默认加载上限；设为 0 表示不限制（无脑全拉）。
export const DEFAULT_MAX_TASKS = 200;

export interface ParsedConfig {
  listId: string;
  maxTasks: number;
}

export function parseConfig(config: Record<string, unknown>): ParsedConfig {
  const c = config as MsTodoConfig;
  const listId = c.listId?.trim();
  if (!listId) {
    throw new Error(
      t({
        zh: '未选择 To Do 列表。请在插件配置中选择一个列表。',
        en: 'No To Do list selected. Please pick a list in the plugin settings.',
      }),
    );
  }
  // select 返回字符串；非法/空值回退默认值，0 表示不限制。
  const rawMax = parseInt((c.maxTasks ?? String(DEFAULT_MAX_TASKS)).toString(), 10);
  const maxTasks = Number.isFinite(rawMax) && rawMax >= 0 ? rawMax : DEFAULT_MAX_TASKS;
  return { listId, maxTasks };
}

// 取 Entra 应用的 clientId：优先用插件配置里的值，否则回退到构建时注入的默认值。
// 两者都为空时抛错（设备码流没有 clientId 无法发起）。
export function getClientId(
  config: Record<string, unknown>,
  buildTimeDefault = '',
): string {
  const fromConfig = (config as { clientId?: string }).clientId?.trim();
  const id = fromConfig || buildTimeDefault;
  if (!id) {
    throw new Error(
      t({
        zh:
          '未配置客户端 ID。请在插件配置中填写，或构建时通过 MS_TODO_CLIENT_ID 环境变量注入。',
        en:
          'No Client ID configured. Fill it in the plugin settings, or inject it at build time via the MS_TODO_CLIENT_ID env var.',
      }),
    );
  }
  return id;
}

// --- 日期转换 ---

// Graph 的 dueDateTime 形如 { dateTime: 'YYYY-MM-DDTHH:mm:ss.0000000', timeZone }。
// SP 的 dueDay 是 'YYYY-MM-DD'。v1 取 date 部分、以 UTC 午夜写回（非 UTC 用户
// 可能看到轻微时区偏移，见 README 已知限制）。
export function extractDueDate(due?: { dateTime: string; timeZone: string } | null): string | undefined {
  if (!due?.dateTime) return undefined;
  return due.dateTime.slice(0, 10);
}

export function buildDueDateTime(
  dueDay: string | null | undefined,
): { dateTime: string; timeZone: string } | null {
  if (!dueDay) return null;
  return { dateTime: dueDay + 'T00:00:00.0000000', timeZone: 'UTC' };
}

// --- 映射 ---

export function mapTaskToIssue(task: TodoTask): PluginIssue {
  return {
    id: task.id,
    title: task.title,
    body: task.body?.content ?? '',
    url: TODO_WEB_URL,
    state: task.status,
    lastUpdated: task.lastModifiedDateTime
      ? new Date(task.lastModifiedDateTime).getTime()
      : undefined,
    // 扩展字段（展示用）
    status: task.status,
    importance: task.importance,
    dueDate: extractDueDate(task.dueDateTime),
    dueDateTime: task.dueDateTime,
  };
}

export function mapTaskToSearchResult(task: TodoTask): PluginSearchResult {
  // SP 导入时把 PluginSearchResult 直接当 issueData 传给字段映射引擎，并调用
  // _computeIsDone / _buildBaseIssueTask，后者读取的是与 PluginIssue 同名的字段
  // （state / body / dueDate / lastUpdated），而非 status / description。因此这里
  // 展开 mapTaskToIssue 的结果，确保完成状态、备注、截止日期在导入时一并带入。
  return {
    ...mapTaskToIssue(task),
    description: task.body?.content ?? '',
    labels: task.importance && task.importance !== 'normal' ? [task.importance] : [],
  };
}

// --- 双向同步字段映射 ---

export const FIELD_MAPPINGS = [
  {
    taskField: 'isDone',
    issueField: 'state',
    defaultDirection: 'both',
    toIssueValue: (taskValue: unknown): string =>
      taskValue ? 'completed' : 'notStarted',
    toTaskValue: (issueValue: unknown): boolean => issueValue === 'completed',
  },
  {
    taskField: 'title',
    issueField: 'title',
    defaultDirection: 'both',
    toIssueValue: (taskValue: unknown): string => (taskValue as string) ?? '',
    toTaskValue: (issueValue: unknown): string => (issueValue as string) ?? '',
  },
  {
    taskField: 'notes',
    issueField: 'body',
    defaultDirection: 'both',
    toIssueValue: (taskValue: unknown): string => (taskValue as string) ?? '',
    toTaskValue: (issueValue: unknown): string => (issueValue as string) ?? '',
  },
  {
    taskField: 'dueDay',
    issueField: 'dueDate',
    defaultDirection: 'both',
    toIssueValue: (taskValue: unknown): string | null =>
      (taskValue as string) || null,
    toTaskValue: (issueValue: unknown): string | undefined =>
      (issueValue as string) || undefined,
  },
] satisfies PluginFieldMapping[];

// --- extractSyncValues ---

export function extractSyncValues(issue: PluginIssue): Record<string, unknown> {
  return {
    state: issue.state,
    title: issue.title,
    body: issue.body ?? '',
    dueDate: issue.dueDate ?? null,
  };
}

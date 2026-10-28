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
  /**
   * To Do 的“截止日期”。实测 Graph 会忽略/归一化其中的时刻（例如写
   * 12-10T06:30Z 会变成 12-10T00:00Z），所以它只承载**日期**。
   */
  dueDateTime?: { dateTime: string; timeZone: string } | null;
  /** To Do 的“提醒”（真正的日期+时刻）。isReminderOn 为 true 时才有效。 */
  reminderDateTime?: { dateTime: string; timeZone: string } | null;
  isReminderOn?: boolean;
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
//
// Graph 的 dueDateTime 形如 { dateTime: 'YYYY-MM-DDTHH:mm:ss.0000000', timeZone }。
//
// MS To Do 没有独立的 all-day 标志：一个“纯日期”提醒实际存成**用户时区当地午夜**
// 的绝对时刻（中国 +8 即前一天 16:00Z）。写回时必须按“本机当地午夜”编码；只写
// UTC 午夜（T00:00:00Z）会变成中国时间 08:00 的**定时**提醒。读回时也不能看墙钟
// 是否 00:00，而要把值还原成绝对时刻后判断“本机是不是正好午夜”。

export interface GraphDue {
  dateTime: string;
  timeZone: string;
}

export interface ParsedGraphDue {
  /** 绝对时刻（毫秒）。 */
  ms: number;
  /** 该时刻在本机时区是否正好是午夜 → 等价于 MS 的“纯日期”。 */
  isDateOnly: boolean;
  /** 纯日期时，本机时区下的日期 YYYY-MM-DD。 */
  localDay: string;
}

const pad2 = (n: number): string => String(n).padStart(2, '0');

interface WallClock {
  y: number;
  mo: number;
  d: number;
  h: number;
  mi: number;
  s: number;
}

function parseWallClock(dateTime: string): WallClock | undefined {
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/.exec(dateTime);
  if (!m) return undefined;
  return {
    y: Number(m[1]),
    mo: Number(m[2]) - 1,
    d: Number(m[3]),
    h: Number(m[4]),
    mi: Number(m[5]),
    s: Number(m[6]),
  };
}

// 墙钟 → 绝对毫秒：UTC 按 UTC；其它时区名按本机墙钟解析。
// （用户设备与 To Do 通常同一时区；跨时区场景见 README 已知限制。）
function wallClockToMs(w: WallClock, timeZone: string): number {
  return timeZone === 'UTC'
    ? Date.UTC(w.y, w.mo, w.d, w.h, w.mi, w.s)
    : new Date(w.y, w.mo, w.d, w.h, w.mi, w.s, 0).getTime();
}

/** 解析 Graph dueDateTime → 绝对时刻 + 是否纯日期 + 本机日期。 */
export function parseGraphDue(due?: GraphDue | null): ParsedGraphDue | undefined {
  if (!due?.dateTime) return undefined;
  const w = parseWallClock(due.dateTime);
  if (!w) return undefined;
  const ms = wallClockToMs(w, due.timeZone);
  if (!Number.isFinite(ms)) return undefined;
  const local = new Date(ms);
  const isLocalMidnight =
    local.getHours() === 0 && local.getMinutes() === 0 && local.getSeconds() === 0;
  // 兼容旧版插件写回的 `T00:00:00Z`（UTC 午夜）：那是当时“纯日期”的编码方式。
  // 不兼容它会让旧任务的同步基线看起来“远端已变”，宿主因此永远跳过本地日期推送。
  const isLegacyUtcMidnight =
    due.timeZone === 'UTC' && w.h === 0 && w.mi === 0 && w.s === 0;
  const isDateOnly = isLocalMidnight || isLegacyUtcMidnight;
  return {
    ms,
    isDateOnly,
    localDay: `${local.getFullYear()}-${pad2(local.getMonth() + 1)}-${pad2(local.getDate())}`,
  };
}

/** Graph → SP 的纯日期（dueDay）；定时提醒返回 undefined。 */
export function extractDueDate(due?: GraphDue | null): string | undefined {
  const parsed = parseGraphDue(due);
  return parsed?.isDateOnly ? parsed.localDay : undefined;
}

/** Graph → SP 的带时刻截止（dueWithTime，绝对毫秒）；纯日期返回 undefined。 */
export function extractDueWithTime(due?: GraphDue | null): number | undefined {
  const parsed = parseGraphDue(due);
  if (!parsed || parsed.isDateOnly) return undefined;
  return parsed.ms;
}

/** 设备时区的 IANA 名称（Graph To Do 接受 IANA；不可用时退回 UTC）。 */
export function deviceTimeZone(): string {
  try {
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    return typeof tz === 'string' && tz ? tz : 'UTC';
  } catch {
    return 'UTC';
  }
}

/**
 * SP 的纯日期（dueDay）→ Graph dueDateTime。
 *
 * 实测（2026-10-28，Graph v1.0）：纯日期必须写成 `<日期>T00:00:00` + **用户时区名**
 * （IANA 如 `Asia/Shanghai` 或 Windows 如 `China Standard Time` 都接受），MS 会把它
 * 归一化成“当地午夜的 UTC 即时”返回。若写成 UTC 午夜（`T00:00:00Z`）会被当成 08:00
 * 的定时提醒；写成“当地午夜的 UTC 即时 + timeZone:'UTC'”则会被 MS 丢弃时刻、只取
 * UTC 日期，导致日期偏移。
 */
export function buildDueDateTime(
  dueDay: string | null | undefined,
  timeZone: string = deviceTimeZone(),
): GraphDue | null {
  if (!dueDay) return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dueDay);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const d = Number(m[3]);
  const local = new Date(y, mo, d, 0, 0, 0, 0);
  // 拒绝 2026-13-40 这类会被 Date 自动进位的输入。
  if (local.getFullYear() !== y || local.getMonth() !== mo || local.getDate() !== d) {
    return null;
  }
  return { dateTime: `${dueDay}T00:00:00.0000000`, timeZone };
}

// 从毫秒时间戳构造 Graph dueDateTime（写回带时刻的截止用）。与 buildDueDateTime
// 一致以 UTC 表示绝对时刻：SP 的 dueWithTime 是绝对毫秒时间戳，写回 UTC 保证
// 提醒触发的绝对时刻正确，To Do 展示时按用户本地时区换算。
export function buildDueDateTimeFromTs(
  ms: number | null | undefined,
): { dateTime: string; timeZone: string } | null {
  if (typeof ms !== 'number' || !Number.isFinite(ms)) return null;
  const d = new Date(ms);
  const p = (n: number): string => String(n).padStart(2, '0');
  const dateTime =
    `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())}` +
    `T${p(d.getUTCHours())}:${p(d.getUTCMinutes())}:${p(d.getUTCSeconds())}.0000000`;
  return { dateTime, timeZone: 'UTC' };
}

/** 绝对时刻 → 本机 'YYYY-MM-DD HH:mm'。 */
export function formatLocalInstant(ms: number): string {
  const d = new Date(ms);
  return (
    `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ` +
    `${pad2(d.getHours())}:${pad2(d.getMinutes())}`
  );
}

/** 本机时区下某时刻的日期 YYYY-MM-DD。 */
export function localDayOf(ms: number): string {
  const d = new Date(ms);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

/**
 * 展示用的截止文案：
 *   - 有提醒（isReminderOn）→ 提醒的日期+时刻；
 *   - 否则有截止日期 → 日期（旧版写入的时刻仍显示为日期+时刻，兼容历史数据）。
 */
export function formatDueDisplay(
  due?: GraphDue | null,
  reminder?: GraphDue | null,
  isReminderOn?: boolean,
): string | undefined {
  if (isReminderOn && reminder) {
    const r = parseGraphDue(reminder);
    if (r) return formatLocalInstant(r.ms);
  }
  const parsed = parseGraphDue(due);
  if (!parsed) return undefined;
  return parsed.isDateOnly ? parsed.localDay : formatLocalInstant(parsed.ms);
}

// --- 映射 ---

export function mapTaskToIssue(task: TodoTask): PluginIssue {
  // To Do 的数据模型：dueDateTime 只承载**日期**，reminderDateTime 承载**时刻**。
  //   - dueDateTime → dueDate（SP 的 dueDay）
  //   - reminderDateTime（isReminderOn=true）→ dueWithTime
  // 兼容旧数据：旧版把时刻写进 dueDateTime（Graph 现在会归一化，但历史值仍在），
  // 没有 reminder 时仍把其时刻当作 dueWithTime。
  const due = parseGraphDue(task.dueDateTime);
  const reminder =
    task.isReminderOn && task.reminderDateTime
      ? parseGraphDue(task.reminderDateTime)
      : undefined;
  const dueDate = due?.localDay;
  const dueWithTime = reminder?.ms ?? (due && !due.isDateOnly ? due.ms : undefined);
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
    dueDate,
    dueWithTime,
    dueDisplay: formatDueDisplay(
      task.dueDateTime,
      task.reminderDateTime,
      task.isReminderOn,
    ),
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
    // 远端没有日期时返回 null（而不是 undefined），让宿主能清除本地旧值：
    // 纯日期 ↔ 定时切换、远端清空日期都依赖这个“显式置空”。
    toTaskValue: (issueValue: unknown): string | null =>
      typeof issueValue === 'string' && issueValue ? issueValue : null,
  },
  {
    taskField: 'dueWithTime',
    issueField: 'dueWithTime',
    defaultDirection: 'both',
    toIssueValue: (taskValue: unknown): number | null =>
      typeof taskValue === 'number' ? taskValue : null,
    toTaskValue: (issueValue: unknown): number | null =>
      typeof issueValue === 'number' ? issueValue : null,
  },
] satisfies PluginFieldMapping[];

// --- extractSyncValues ---

export function extractSyncValues(issue: PluginIssue): Record<string, unknown> {
  return {
    state: issue.state,
    title: issue.title,
    body: issue.body ?? '',
    dueDate: issue.dueDate ?? null,
    dueWithTime: typeof issue.dueWithTime === 'number' ? issue.dueWithTime : null,
  };
}

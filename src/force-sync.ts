// 「强制同步至 Microsoft To Do」按钮的纯逻辑（无宿主副作用），便于单元测试。
//
// 背景：SP 桌面版（Electron）渲染进程里，Graph 的写请求（POST/PATCH/DELETE）会因
// 宿主的 CORS 预检处理失败为 status=0，导致 createIssue 自动创建失败。此按钮提供
// 手动补救：选中一个本插件未同步任务，走与 createIssue 相同的宿主 http
// 通道重新 POST 到 Graph，成功后回写 issueId 完成链接。
//
// 说明：按钮全程走渲染进程 PluginHttp（`http` 权限），移动端 / 桌面端 / Web 三端
// 一致可用 —— 不依赖 nodeExecution（那会令 SP 在移动端拒载本插件）。桌面端若仍受
// 宿主 status=0 限制，则属宿主层面的问题，插件侧无法在不牺牲移动端的前提下绕过。
//
// 本文件只放纯函数；接宿主 PluginAPI 的副作用组装在 plugin.ts。

import type { SpTask } from './types';
import { buildDueDateTimeFromTs } from './domain';

export const PLUGIN_ID = 'ms-todo-provider';
export const PLUGIN_ISSUE_TYPE = `plugin:${PLUGIN_ID}`;

/** Graph / To Do 的“已完成”状态值（字段映射与对账都以它为准）。 */
export const COMPLETED_STATE = 'completed';

/** 判断任务是否属于本插件（issueType 形如 plugin:ms-todo-provider）。 */
export function isMsTodoTask(task: SpTask | null | undefined): task is SpTask {
  return !!task && task.issueType === PLUGIN_ISSUE_TYPE;
}

/** 判断任务是否尚未推送到 MS（没有 issueId）。 */
export function isUnsynced(task: SpTask): boolean {
  return !task.issueId;
}

/**
 * 判断任务是否可被「强制同步」推送并链接到 MS To Do。
 *
 * 关键：自动创建失败的任务（桌面端 status=0 时宿主会弹「Auto-create issue failed」，
 * 但**不会**改写任务，见宿主 issue-two-way-sync.effects 的 autoCreateIssueOnTaskAdd$）
 * 既没有 issueId 也没有 issueType —— 若只按 isMsTodoTask 过滤会被误判为「非本插件
 * 任务」。因此这里放宽为：
 *   - 有 issueId 的一律不可推（已同步）；
 *   - 无 issueType 的本地任务视为「创建失败待推送」，可推；
 *   - issueType 属于本插件的可推；
 *   - issueType 属于其它 provider 的不可推。
 */
export function isPushableTask(task: SpTask | null | undefined): task is SpTask {
  if (!task || task.issueId) return false;
  return !task.issueType || task.issueType === PLUGIN_ISSUE_TYPE;
}

/** 选出要推送的任务：仅当前选中的、尚未同步（无 issueId）且非其它 provider 的任务。 */
export function selectTasksToPush(
  selected: SpTask | null | undefined,
): SpTask[] {
  return isPushableTask(selected) ? [selected] : [];
}

/**
 * 选出「手动同步」要推送的全部任务：处于项目上下文时仅该项目的任务，否则全部；
 * 再过滤为可推送（无 issueId、非其它 provider）的任务。与独立视图的 scope 规则一致。
 */
export function selectTasksToSync(
  all: SpTask[],
  projectId?: string,
): SpTask[] {
  const scoped = projectId ? all.filter((task) => task.projectId === projectId) : all;
  return scoped.filter(isPushableTask);
}

/**
 * 从已链接的本插件任务里找同一项目的 provider 实例 id（issueProviderId）。
 *
 * 自动创建失败的任务没有 issueProviderId，但成功回写链接时必须带上它，否则宿主
 * two-way-sync 无法把后续更新路由到正确的 provider 实例。每个项目只有一个 provider
 * 实例，所以同项目里任意一个已链接任务携带的 issueProviderId 就是我们要的。
 */
export function findIssueProviderId(
  tasks: SpTask[],
  projectId?: string,
): string | undefined {
  for (const task of tasks) {
    if (!isMsTodoTask(task) || !task.issueId || !task.issueProviderId) continue;
    if (projectId && task.projectId !== projectId) continue;
    return task.issueProviderId;
  }
  return undefined;
}

/** 选出「已链接」（带 issueId）的本插件任务，用于失效链接检测。 */
export function selectLinkedMsTasks(all: SpTask[]): SpTask[] {
  return all.filter((t) => isMsTodoTask(t) && !!t.issueId);
}

// --- 完成状态对账（completion axis）---
//
// 背景：宿主的 two-way-sync 只在“任务更新动作”发生时推送；推送失败（瞬时网络 /
// 桌面 CORS status=0）或 provider-changed 跳过后不会重试。拉取方向又可能因批次
// 过长被宿主重启而饿死。结果就是本地 isDone 与 issueLastSyncedValues.state 长期
// 分叉、且无人修复。下面这组纯函数供两处使用：
//   1) taskUpdate hook：用户在本插件任务上打勾时，插件自己兜底推一次；
//   2) 视图「对账」：双向扫描历史分叉并逐条修复，同时给出可读原因。

export type CompletionSyncAction = 'push' | 'pull' | 'baseline-only' | 'none';

export interface CompletionSyncInput {
  /** 本地任务是否已完成。 */
  isDone: boolean;
  /** 本地记录的“上次同步时远端的状态”（issueLastSyncedValues.state）。 */
  baselineState?: string;
  /** 刚刚从 Graph 读到的远端 state；undefined 表示没读到（失败）。 */
  remoteState?: string;
}

export interface CompletionSyncDecision {
  action: CompletionSyncAction;
  reason: string;
}

/**
 * 决定完成状态该怎么对齐。
 *
 * - 远端已完成、本地未完成 → pull（把本地打勾）
 * - 本地已完成、远端未完成 → push（把远端置为 completed；显式打勾的本地意图优先）
 * - 本地已完成、远端也已完成 → 只回写基线（远端已经是目标状态，无需写）
 * - 远端状态读不到 → none（不做任何写操作）
 */
export function decideCompletionSync(input: CompletionSyncInput): CompletionSyncDecision {
  const { isDone, baselineState, remoteState } = input;
  if (remoteState === undefined || remoteState === null) {
    return { action: 'none', reason: 'remote state unavailable' };
  }
  const remoteDone = remoteState.toLowerCase() === COMPLETED_STATE;
  if (isDone) {
    if (remoteDone) {
      return { action: 'baseline-only', reason: 'remote already completed' };
    }
    if (baselineState !== undefined && baselineState !== remoteState) {
      return {
        action: 'push',
        reason: 'remote changed since baseline; local completion wins',
      };
    }
    return { action: 'push', reason: 'push local completion' };
  }
  if (remoteDone) {
    return { action: 'pull', reason: 'remote completed; mark local task done' };
  }
  return { action: 'none', reason: 'both open' };
}

/**
 * taskUpdate hook 的守卫：只有“本插件任务被本地打勾”才需要兜底推送。
 *
 * 关键排除项：宿主轮询拉取完成状态时也会派发 updateTask（changes 里带
 * `issueLastSyncedValues` / `issueWasUpdated`），那不是用户意图，不能反向再推一次。
 */
export function shouldPushCompletionOnUpdate(
  task: SpTask | null | undefined,
  changes: Record<string, unknown> | null | undefined,
): task is SpTask {
  if (!task || !isMsTodoTask(task) || !task.issueId) return false;
  if (!changes || changes['isDone'] !== true) return false;
  if ('issueLastSyncedValues' in changes || 'issueWasUpdated' in changes) return false;
  const baseline = task.issueLastSyncedValues as Record<string, unknown> | undefined;
  const state = baseline && typeof baseline['state'] === 'string' ? baseline['state'] : undefined;
  return state !== COMPLETED_STATE;
}

/**
 * 选出完成状态分叉的任务：本地 isDone 与基线 state 不一致。
 * 两个方向都会入选（push 与 pull），由 decideCompletionSync 决定具体动作。
 */
export function selectCompletionDivergences(tasks: SpTask[]): SpTask[] {
  return tasks.filter((task) => {
    if (!isMsTodoTask(task) || !task.issueId) return false;
    const baseline = task.issueLastSyncedValues as Record<string, unknown> | undefined;
    const state = baseline && typeof baseline['state'] === 'string' ? baseline['state'] : undefined;
    return !!task.isDone !== (state === COMPLETED_STATE);
  });
}

/** 把已完成状态写进同步基线（保留其它已同步字段）。 */
export function withCompletedBaseline(
  baseline: Record<string, unknown> | undefined,
): Record<string, unknown> {
  return { ...(baseline ?? {}), state: COMPLETED_STATE };
}

/** getByIds 批量接口用：只保留请求过的 issue，保持 provider 返回顺序。 */
export function filterIssuesByIds<T extends { id: string }>(issues: T[], ids: string[]): T[] {
  const wanted = new Set(ids);
  return issues.filter((issue) => wanted.has(issue.id));
}

/** Graph 创建任务请求体。 */
export interface CreateTaskBody {
  title: string;
  status?: 'notStarted' | 'completed';
  body?: { content: string; contentType: 'text' };
  dueDateTime?: { dateTime: string; timeZone: string } | null;
}

/** 用本地任务字段构建 Graph 创建任务请求体（与 domain.ts 的字段映射一致）。 */
export function buildCreateTaskBody(
  task: Pick<SpTask, 'title' | 'notes' | 'dueDay' | 'dueWithTime' | 'isDone'>,
): CreateTaskBody {
  const body: CreateTaskBody = { title: task.title };
  if (task.isDone) body.status = 'completed';
  if (task.notes) body.body = { content: task.notes, contentType: 'text' };
  if (typeof task.dueWithTime === 'number') {
    body.dueDateTime = buildDueDateTimeFromTs(task.dueWithTime);
  } else if (task.dueDay) {
    body.dueDateTime = { dateTime: task.dueDay + 'T00:00:00.0000000', timeZone: 'UTC' };
  }
  return body;
}

/**
 * 由「MS 任务 id → 列表 id」映射重建 projectId → listId。
 *
 * 已链接任务带 projectId + issueId；MS 任务只会落在唯一一个列表里，所以能据此反推出
 * 每个 SP 项目绑定的 To Do 列表（一个项目一个 provider 实例 = 一个列表）。同一项目下
 * 多个已链接任务若映射到不同列表（异常数据），取第一个。
 */
export function buildProjectToListMap(
  linkedTasks: SpTask[],
  msTaskToList: Map<string, string>,
): Map<string, string> {
  const map = new Map<string, string>();
  for (const task of linkedTasks) {
    const listId = task.issueId ? msTaskToList.get(task.issueId) : undefined;
    if (listId && task.projectId && !map.has(task.projectId)) {
      map.set(task.projectId, listId);
    }
  }
  return map;
}

/**
 * 解析一次推送的目标列表：显式指定 > 任务所属项目绑定列表 > 兜底列表。
 * 都拿不到时返回 undefined（调用方报「无法确定目标列表」）。
 */
export function resolvePushListId(
  task: SpTask,
  explicitListId: string | undefined,
  projectToList: Map<string, string>,
  fallbackListId: string | undefined,
): string | undefined {
  if (explicitListId) return explicitListId;
  return projectToList.get(task.projectId as string) ?? fallbackListId;
}

/** 众数：取出现次数最多的非空值；平票时取先出现的（Map 按插入序迭代）。 */
export function majorityList(values: (string | undefined)[]): string | undefined {
  const counts = new Map<string, number>();
  for (const value of values) {
    if (!value) continue;
    counts.set(value, (counts.get(value) ?? 0) + 1);
  }
  let best: string | undefined;
  let bestCount = 0;
  for (const [value, count] of counts) {
    if (count > bestCount) {
      best = value;
      bestCount = count;
    }
  }
  return best;
}

export interface DeadLinkInfo {
  taskId: string;
  title: string;
  issueId: string;
}

export interface MismatchedLinkInfo {
  taskId: string;
  title: string;
  issueId: string;
  /** 任务实际所在的列表。 */
  listId: string;
  /** 任务所属项目「应属」的列表（该项目已链接任务的众数列表）。 */
  expectedListId: string;
}

export interface LinkHealth {
  checked: number;
  dead: DeadLinkInfo[];
  mismatched: MismatchedLinkInfo[];
}

/**
 * 把已链接任务按「MS 任务 id → 列表 id」索引做体检：
 *   - issueId 在任何列表都找不到 → dead（远端已删除）；
 *   - 找到了、但与其项目「应属列表」（众数）不同 → mismatched（被挪走/推错列表）。
 *
 * 判定依据：一个 SP 项目 = 一个 provider 实例 = 一个 To Do 列表，所以同一项目里
 * 已链接任务应全部落在同一列表；出现少数落点不同即为不一致。项目只有一个链接、
 * 或整项目任务被一起挪走时无法靠众数发现（此时由写路径的 404 自动定位兜底）。
 */
export function classifyLinkedTasks(
  linkedTasks: SpTask[],
  msTaskToList: Map<string, string>,
): LinkHealth {
  const listsByProject = new Map<string, (string | undefined)[]>();
  for (const task of linkedTasks) {
    const actual = task.issueId ? msTaskToList.get(task.issueId) : undefined;
    const projectId = task.projectId ?? '';
    const arr = listsByProject.get(projectId);
    if (arr) arr.push(actual);
    else listsByProject.set(projectId, [actual]);
  }
  const boundByProject = new Map<string, string>();
  for (const [projectId, arr] of listsByProject) {
    const bound = majorityList(arr);
    if (bound) boundByProject.set(projectId, bound);
  }

  const dead: DeadLinkInfo[] = [];
  const mismatched: MismatchedLinkInfo[] = [];
  for (const task of linkedTasks) {
    const issueId = task.issueId as string;
    const actual = msTaskToList.get(issueId);
    if (!actual) {
      dead.push({ taskId: task.id, title: task.title, issueId });
      continue;
    }
    const expected = boundByProject.get(task.projectId ?? '');
    if (expected && actual !== expected) {
      mismatched.push({
        taskId: task.id,
        title: task.title,
        issueId,
        listId: actual,
        expectedListId: expected,
      });
    }
  }
  return { checked: linkedTasks.length, dead, mismatched };
}


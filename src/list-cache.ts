// 每个 To Do 列表的内存任务快照 + 增量刷新。
//
// 背景（2026-10-26 实测）：旧版 getByIds 为了覆盖“已被挪走/已删除”的 id，
// 会在每次 60s 轮询里把整张列表分页拉完（最多 40 页），再对最多 20 个缺失 id
// 逐个 getById → 404 → findTaskList（列全部列表 + 每个列表 probe）。5 个 provider
// 实例叠加后，每分钟数百次 Graph 请求，直接触发 429 / ERR_HTTP2_PROTOCOL_ERROR。
// 而宿主 `_getFreshDataForProviderTasks` 对 getByIds 的任何异常都是 `catch → []`，
// 于是“一次限流 = 这一轮该 provider 的所有远端更新（含完成状态）全部丢弃”。
//
// 新策略：
//   - 每个列表在进程内缓存一份 id → task 快照；冷启动整表拉一次；
//   - 之后每 TTL 只发一条 `$filter=lastModifiedDateTime gt <上次同步-2min>` 增量；
//   - getByIds 只读缓存，绝不因为缺失 id 逐个探测（缺失由视图“体检”处理）；
//   - 任一请求失败都保留旧缓存并返回可用结果，不抛错、不丢整批。
//
// 没有持久化：插件代码/存储是设备本地的，进程重启后重新整表拉一次（每进程一次）。

import type { TodoTask } from './domain';
import { graphRequestAll } from './graph';
import type { PluginHttp } from './types';

/** 增量刷新后多久内复用缓存（毫秒）。宿主轮询是 60s，这里只需挡住同轮多个 provider。 */
const REFRESH_TTL_MS = 15_000;

/** 增量过滤的重叠窗口：避免时钟误差/服务端写入延迟漏掉边界上的修改。 */
const OVERLAP_MS = 120_000;

/** 定期用整表刷新一次，顺带清理远端已删除的缓存条目。 */
const FULL_REFRESH_INTERVAL_MS = 30 * 60_000;

/** 整表拉取的安全上限，防止超大列表把渲染进程拖死（正常列表远小于此）。 */
const FULL_FETCH_LIMIT = 5000;

/** Graph To Do 支持 $top=1000（实测单页可返回 900+ 条），把整表/增量都压成 1 次请求。 */
const PAGE_SIZE = 1000;

interface ListEntry {
  tasks: Map<string, TodoTask>;
  fullLoaded: boolean;
  /** 上次成功同步（整表或增量）的时间。 */
  lastSyncAt: number;
  /** 上次整表拉取的时间。 */
  lastFullAt: number;
}

export interface ListTasksResult {
  tasks: TodoTask[];
  missing: string[];
}

export interface RefreshInfo {
  listId: string;
  mode: 'full' | 'incremental' | 'skipped';
  /** 增量过滤的起点（ISO 秒）。 */
  since?: string;
  /** 本次刷新合并到的远端条目数。 */
  changed?: number;
  /** 刷新后缓存里的条目数。 */
  cached: number;
  error?: string;
}

export interface TodoTaskCacheOptions {
  now?: () => number;
  onError?: (message: string, error: unknown) => void;
  onRefresh?: (info: RefreshInfo) => void;
}

const lastModifiedOf = (task: TodoTask): number => {
  const ms = task.lastModifiedDateTime ? Date.parse(task.lastModifiedDateTime) : NaN;
  return Number.isFinite(ms) ? ms : 0;
};

const isoSeconds = (ms: number): string => new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z');

const tasksPath = (listId: string): string => `/me/todo/lists/${listId}/tasks`;

export class TodoTaskCache {
  private lists = new Map<string, ListEntry>();
  private inflight = new Map<string, Promise<void>>();
  private readonly now: () => number;
  private readonly onError?: (message: string, error: unknown) => void;
  private readonly onRefresh?: (info: RefreshInfo) => void;

  constructor(opts: TodoTaskCacheOptions = {}) {
    this.now = opts.now ?? (() => Date.now());
    this.onError = opts.onError;
    this.onRefresh = opts.onRefresh;
  }

  /** 仅供测试/诊断：清空全部缓存。 */
  reset(): void {
    this.lists.clear();
    this.inflight.clear();
  }

  /** 缓存中任务总数（跨列表去重前）。 */
  size(): number {
    let n = 0;
    for (const entry of this.lists.values()) n += entry.tasks.size;
    return n;
  }

  /** 某个 MS 任务 id 当前所在列表（按最新副本判断）；未缓存返回 undefined。 */
  listIdOf(issueId: string): string | undefined {
    return this.bestCopy(issueId)?.listId;
  }

  /** 单条缓存读取（最新副本）。 */
  getCached(issueId: string): TodoTask | undefined {
    return this.bestCopy(issueId)?.task;
  }

  /**
   * 把单条任务写入缓存（getById / createIssue 的返回值）。
   * 列表尚未整表加载时不标记 fullLoaded，下次 getByIds 仍会整表拉取并覆盖。
   */
  remember(listId: string, task: TodoTask): void {
    if (!listId || !task?.id) return;
    const entry = this.entryFor(listId);
    entry.tasks.set(task.id, task);
  }

  forget(issueId: string): void {
    for (const entry of this.lists.values()) entry.tasks.delete(issueId);
  }

  /**
   * 整张列表的当前快照（搜索 / backlog 自动导入用）。
   * 与 getByIds 共用缓存，避免这些路径再次触发整表分页。
   */
  async getAll(listId: string, token: string, http: PluginHttp): Promise<TodoTask[]> {
    await this.refresh(listId, token, http);
    return [...this.entryFor(listId).tasks.values()];
  }

  /**
   * 返回请求 id 对应的任务（跨列表取最新副本），并给出缓存里完全不存在的 id。
   * 绝不抛错：刷新失败时返回旧缓存。
   */
  async getByIds(
    listId: string,
    issueIds: string[],
    token: string,
    http: PluginHttp,
  ): Promise<ListTasksResult> {
    const ids = [...new Set(issueIds.filter((id) => !!id))];
    if (ids.length === 0) return { tasks: [], missing: [] };

    await this.refresh(listId, token, http);

    const tasks: TodoTask[] = [];
    const missing: string[] = [];
    for (const id of ids) {
      const best = this.bestCopy(id);
      if (best) tasks.push(best.task);
      else missing.push(id);
    }
    return { tasks, missing };
  }

  private entryFor(listId: string): ListEntry {
    let entry = this.lists.get(listId);
    if (!entry) {
      entry = { tasks: new Map(), fullLoaded: false, lastSyncAt: 0, lastFullAt: 0 };
      this.lists.set(listId, entry);
    }
    return entry;
  }

  private bestCopy(issueId: string): { task: TodoTask; listId: string } | undefined {
    let best: { task: TodoTask; listId: string } | undefined;
    for (const [listId, entry] of this.lists) {
      const task = entry.tasks.get(issueId);
      if (!task) continue;
      if (!best || lastModifiedOf(task) > lastModifiedOf(best.task)) {
        best = { task, listId };
      }
    }
    return best;
  }

  /** 刷新一个列表：同列表并发调用共享同一个 Promise；TTL 内直接复用缓存。 */
  private async refresh(listId: string, token: string, http: PluginHttp): Promise<void> {
    const running = this.inflight.get(listId);
    if (running) return running;

    const entry = this.entryFor(listId);
    const now = this.now();
    if (entry.fullLoaded && now - entry.lastSyncAt < REFRESH_TTL_MS) return;

    const p = this.doRefresh(listId, entry, token, http).finally(() => {
      this.inflight.delete(listId);
    });
    this.inflight.set(listId, p);
    return p;
  }

  private async doRefresh(
    listId: string,
    entry: ListEntry,
    token: string,
    http: PluginHttp,
  ): Promise<void> {
    const now = this.now();
    const needsFull = !entry.fullLoaded || now - entry.lastFullAt >= FULL_REFRESH_INTERVAL_MS;
    try {
      if (needsFull) {
        const tasks = await graphRequestAll<TodoTask>(
          `${tasksPath(listId)}?$top=${PAGE_SIZE}`,
          token,
          FULL_FETCH_LIMIT,
          http,
        );
        const remembered = entry.tasks;
        entry.tasks = new Map(tasks.map((task) => [task.id, task] as const));
        // 整表结果更权威；只保留整表里没有的、通过单条接口写入的任务。
        for (const [id, task] of remembered) {
          if (!entry.tasks.has(id)) entry.tasks.set(id, task);
        }
        entry.fullLoaded = true;
        entry.lastFullAt = now;
        entry.lastSyncAt = now;
        this.onRefresh?.({
          listId,
          mode: 'full',
          changed: tasks.length,
          cached: entry.tasks.size,
        });
        return;
      }

      const since = Math.max(0, entry.lastSyncAt - OVERLAP_MS);
      const filter = encodeURIComponent(`lastModifiedDateTime gt ${isoSeconds(since)}`);
      const changed = await graphRequestAll<TodoTask>(
        `${tasksPath(listId)}?$filter=${filter}&$top=${PAGE_SIZE}`,
        token,
        0,
        http,
      );
      for (const task of changed) entry.tasks.set(task.id, task);
      entry.lastSyncAt = now;
      this.onRefresh?.({
        listId,
        mode: 'incremental',
        since: isoSeconds(since),
        changed: changed.length,
        cached: entry.tasks.size,
      });
    } catch (e) {
      // 保留旧缓存：宁可这一轮数据旧一点，也不能让宿主的整批 pull 变空。
      this.onError?.(`task cache refresh failed for list ${listId}`, e);
      this.onRefresh?.({
        listId,
        mode: needsFull ? 'full' : 'incremental',
        cached: entry.tasks.size,
        error: e instanceof Error ? e.message : String(e),
      });
    }
  }
}

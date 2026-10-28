import { describe, expect, it, vi } from 'vitest';
import { TodoTaskCache } from '../../src/list-cache';
import type { PluginHttp } from '../../src/types';

const BASE = 'https://graph.microsoft.com/v1.0/me/todo/lists';
const tasksUrl = (listId: string): string => `${BASE}/${listId}/tasks`;
const page = (tasks: { id: string; lastModifiedDateTime?: string }[], next?: string) => ({
  value: tasks.map((t) => ({ title: t.id, status: 'notStarted', ...t })),
  ...(next ? { '@odata.nextLink': next } : {}),
});

type Handler = (method: string, url: string) => unknown;

function httpOf(handler: Handler): { http: PluginHttp; calls: string[] } {
  const calls: string[] = [];
  const http = {
    request: vi.fn(async (method: string, url: string) => {
      calls.push(`${method} ${url}`);
      const res = handler(method, url);
      if (res instanceof Error) throw res;
      return res;
    }),
  } as unknown as PluginHttp;
  return { http, calls };
}

/** 每个列表第一次请求返回整表；带 $filter 的请求返回 handler 指定的增量。 */
function fullList(listId: string, tasks: { id: string; lastModifiedDateTime?: string }[]): Handler {
  return (method, url) => {
    if (method !== 'GET') throw new Error('unexpected ' + method);
    if (!url.startsWith(tasksUrl(listId))) throw new Error('unexpected url ' + url);
    if (url.includes('$filter')) throw new Error('unexpected filter');
    return page(tasks);
  };
}

describe('TodoTaskCache — poll 请求预算', () => {
  it('冷启动只做一次整表拉取，缺失 id 不再逐个探测', async () => {
    const { http, calls } = httpOf(fullList('L1', [{ id: 'a' }, { id: 'b' }]));
    const cache = new TodoTaskCache({ now: () => 1_000 });
    const res = await cache.getByIds('L1', ['a', 'zzz'], 'tok', http);

    expect(res.tasks.map((t) => t.id)).toEqual(['a']);
    expect(res.missing).toEqual(['zzz']);
    expect(calls).toHaveLength(1);
    expect(calls[0]).toContain(tasksUrl('L1'));
    expect(calls.some((c) => c.includes('zzz'))).toBe(false);
  });

  it('TTL 内重复调用不再发请求（多个 provider 共享同一列表）', async () => {
    let now = 10_000;
    const { http, calls } = httpOf(fullList('L1', [{ id: 'a' }]));
    const cache = new TodoTaskCache({ now: () => now });

    await cache.getByIds('L1', ['a'], 'tok', http);
    now += 5_000;
    await cache.getByIds('L1', ['a'], 'tok', http);

    expect(calls).toHaveLength(1);
  });

  it('超过 TTL 后只拉 lastModifiedDateTime 之后的增量', async () => {
    let now = 10_000;
    const { http, calls } = httpOf((method, url) => {
      if (!url.startsWith(tasksUrl('L1'))) throw new Error('unexpected url ' + url);
      if (url.includes('$filter')) return page([{ id: 'b', lastModifiedDateTime: '2026-01-01T00:00:00Z' }]);
      return page([{ id: 'a' }]);
    });
    const cache = new TodoTaskCache({ now: () => now });

    await cache.getByIds('L1', ['a'], 'tok', http);
    now += 60_000;
    const res = await cache.getByIds('L1', ['a', 'b'], 'tok', http);

    expect(res.tasks.map((t) => t.id).sort()).toEqual(['a', 'b']);
    expect(calls).toHaveLength(2);
    expect(decodeURIComponent(calls[1])).toContain('lastModifiedDateTime gt');
  });

  it('增量请求失败时返回缓存数据而不是丢掉整批', async () => {
    let now = 10_000;
    let fail = false;
    const { http } = httpOf((_method, url) => {
      if (!url.startsWith(tasksUrl('L1'))) throw new Error('unexpected url ' + url);
      if (url.includes('$filter')) return fail ? new Error('status=429 throttled') : page([{ id: 'b' }]);
      return page([{ id: 'a' }]);
    });
    const errors: string[] = [];
    const cache = new TodoTaskCache({ now: () => now, onError: (msg) => errors.push(msg) });

    await cache.getByIds('L1', ['a'], 'tok', http);
    now += 60_000;
    fail = true;
    const res = await cache.getByIds('L1', ['a'], 'tok', http);

    expect(res.tasks.map((t) => t.id)).toEqual(['a']);
    expect(errors.length).toBe(1);
  });

  it('冷启动拉取失败时返回空结果且不抛错', async () => {
    const { http } = httpOf(() => Object.assign(new Error('bad request'), { status: 400 }));
    const cache = new TodoTaskCache({ now: () => 1_000 });

    const res = await cache.getByIds('L1', ['a'], 'tok', http);
    expect(res.tasks).toEqual([]);
    expect(res.missing).toEqual(['a']);
  });

  it('同一 id 出现在多个列表缓存时取 lastModifiedDateTime 最新的一份（跨列表移动）', async () => {
    const now = 10_000;
    const { http } = httpOf((_method, url) => {
      if (url.startsWith(tasksUrl('A'))) {
        return page([{ id: 'x', lastModifiedDateTime: '2026-01-01T00:00:00Z' }]);
      }
      if (url.startsWith(tasksUrl('B'))) {
        return page([{ id: 'x', lastModifiedDateTime: '2026-02-02T00:00:00Z' }]);
      }
      throw new Error('unexpected url ' + url);
    });
    const cache = new TodoTaskCache({ now: () => now });

    await cache.getByIds('A', ['x'], 'tok', http);
    await cache.getByIds('B', ['x'], 'tok', http);
    const res = await cache.getByIds('A', ['x'], 'tok', http);

    expect(res.tasks[0].lastModifiedDateTime).toBe('2026-02-02T00:00:00Z');
    expect(cache.listIdOf('x')).toBe('B');
  });

  it('getAll（搜索/backlog）复用同一份列表缓存，不额外发请求', async () => {
    const { http, calls } = httpOf(fullList('L1', [{ id: 'a' }, { id: 'b' }]));
    const cache = new TodoTaskCache({ now: () => 1_000 });

    const all = await cache.getAll('L1', 'tok', http);
    expect(all.map((t) => t.id).sort()).toEqual(['a', 'b']);
    const viaIds = await cache.getByIds('L1', ['a'], 'tok', http);
    expect(viaIds.tasks.map((t) => t.id)).toEqual(['a']);
    expect(calls).toHaveLength(1);
  });

  it('remember/forget 保持缓存与单条读取一致', async () => {
    const { http } = httpOf(fullList('L1', []));
    const cache = new TodoTaskCache({ now: () => 1_000 });
    await cache.getByIds('L1', [], 'tok', http);

    cache.remember('L1', { id: 'x', title: 'x', status: 'notStarted' });
    expect(cache.getCached('x')?.title).toBe('x');
    cache.forget('x');
    expect(cache.getCached('x')).toBeUndefined();
  });
});

describe('TodoTaskCache — 远端删除检测', () => {
  it('整表刷新是权威的：远端已删的任务会从快照消失并出现在 missing 里', async () => {
    let now = 10_000;
    let remote = [{ id: 'a' }, { id: 'b' }];
    const { http } = httpOf((_method, url) => {
      if (!url.startsWith(tasksUrl('L1'))) throw new Error('unexpected url ' + url);
      if (url.includes('$filter')) throw new Error('unexpected filter');
      return page(remote);
    });
    const cache = new TodoTaskCache({ now: () => now });

    const first = await cache.getByIds('L1', ['a', 'b'], 'tok', http);
    expect(first.tasks.map((t) => t.id).sort()).toEqual(['a', 'b']);
    expect(first.trustworthy).toBe(true);

    // 远端删除了 b；等到下一次整表刷新。
    remote = [{ id: 'a' }];
    now += 31 * 60_000;
    const second = await cache.getByIds('L1', ['a', 'b'], 'tok', http);

    expect(second.tasks.map((t) => t.id)).toEqual(['a']);
    expect(second.missing).toEqual(['b']);
    expect(cache.getCached('b')).toBeUndefined();
  });

  it('整表刷新期间通过单条接口写入的任务不会被误删', async () => {
    let cache: TodoTaskCache;
    const now = 10_000;
    const { http } = httpOf((_method, url) => {
      if (!url.startsWith(tasksUrl('L1'))) throw new Error('unexpected url ' + url);
      if (url.includes('$filter')) throw new Error('unexpected filter');
      // 模拟 createIssue/getById 在整表请求进行中写入了 x（远端整表结果还没有它）。
      cache.remember('L1', { id: 'x', title: 'x', status: 'notStarted' });
      return page([{ id: 'a' }]);
    });
    cache = new TodoTaskCache({ now: () => now });

    const res = await cache.getByIds('L1', ['a', 'x'], 'tok', http);
    expect(res.tasks.map((t) => t.id).sort()).toEqual(['a', 'x']);
  });

  it('整表刷新失败时 trustworthy=false（禁止把 missing 当删除）', async () => {
    const { http } = httpOf(() => Object.assign(new Error('status=400'), { status: 400 }));
    const cache = new TodoTaskCache({ now: () => 1_000 });
    const res = await cache.getByIds('L1', ['a'], 'tok', http);
    expect(res.trustworthy).toBe(false);
    expect(res.missing).toEqual(['a']);
  });

  it('增量刷新失败后 trustworthy=false', async () => {
    let now = 10_000;
    let fail = false;
    const { http } = httpOf((_method, url) => {
      if (!url.startsWith(tasksUrl('L1'))) throw new Error('unexpected url ' + url);
      if (url.includes('$filter')) return fail ? Object.assign(new Error('status=400'), { status: 400 }) : page([]);
      return page([{ id: 'a' }]);
    });
    const cache = new TodoTaskCache({ now: () => now });
    expect((await cache.getByIds('L1', ['a'], 'tok', http)).trustworthy).toBe(true);

    now += 60_000;
    fail = true;
    const res = await cache.getByIds('L1', ['a'], 'tok', http);
    expect(res.trustworthy).toBe(false);
    expect(res.tasks.map((t) => t.id)).toEqual(['a']);
  });
});

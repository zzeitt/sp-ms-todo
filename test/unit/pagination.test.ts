import { describe, it, expect } from 'vitest';
import { graphRequestAll } from '../../src/graph';
import type { PluginHttp } from '../../src/types';

// graphRequestAll → graphRequest → http.request（宿主 PluginHttp）。这里脚本化
// http.request 的返回值，端到端覆盖：分页循环跟随 nextLink、limit 截断、
// 绝对 nextLink 直连、以及请求错误的透传。

const BASE = 'https://graph.microsoft.com/v1.0';

interface Page {
  value: unknown[];
  '@odata.nextLink'?: string;
}

interface Call {
  method: string;
  url: string;
}

function httpPages(pages: Page[]): { calls: Call[]; http: PluginHttp } {
  const calls: Call[] = [];
  const http = {
    request: async <T,>(method: string, url: string) => {
      const page = pages[calls.length];
      if (!page) throw new Error('unexpected extra http request');
      calls.push({ method, url });
      return page as unknown as T;
    },
  } as unknown as PluginHttp;
  return { calls, http };
}

describe('graphRequestAll', () => {
  it('follows @odata.nextLink across pages', async () => {
    const { http } = httpPages([
      { value: [1, 2], '@odata.nextLink': `${BASE}/next?$skiptoken=abc` },
      { value: [3] },
    ]);
    expect(await graphRequestAll<number>('/me/todo/lists/L/tasks', 'tok', 0, http)).toEqual([
      1, 2, 3,
    ]);
  });

  it('passes the absolute nextLink URL through unchanged', async () => {
    const { calls, http } = httpPages([
      { value: [1], '@odata.nextLink': `${BASE}/next?$skiptoken=abc` },
      { value: [2] },
    ]);
    await graphRequestAll<number>('/me/todo/lists/L/tasks', 'tok', 0, http);
    expect(calls[0].url).toBe(`${BASE}/me/todo/lists/L/tasks`);
    expect(calls[1].url).toBe(`${BASE}/next?$skiptoken=abc`);
    for (const c of calls) expect(c.method).toBe('GET');
  });

  it('stops once the limit is reached (no extra page fetched)', async () => {
    const { calls, http } = httpPages([
      { value: [1, 2, 3, 4, 5], '@odata.nextLink': `${BASE}/next` },
    ]);
    expect(await graphRequestAll<number>('/me/todo/lists/L/tasks', 'tok', 3, http)).toEqual([
      1, 2, 3,
    ]);
    expect(calls.length).toBe(1); // 未再追 nextLink
  });

  it('returns empty for an empty page', async () => {
    const { http } = httpPages([{ value: [] }]);
    expect(await graphRequestAll<number>('/me/todo/lists/L/tasks', 'tok', 0, http)).toEqual([]);
  });

  it('limit 0 means no limit (pulls all pages)', async () => {
    const { http } = httpPages([
      { value: [1], '@odata.nextLink': `${BASE}/next` },
      { value: [2] },
    ]);
    expect(await graphRequestAll<number>('/me/todo/lists/L/tasks', 'tok', 0, http)).toEqual([1, 2]);
  });

  it('propagates http.request errors', async () => {
    const http = {
      request: async () => {
        throw new Error('boom');
      },
    } as unknown as PluginHttp;
    await expect(graphRequestAll<number>('/x', 'tok', 0, http)).rejects.toThrow('boom');
  });

  it('dedups entries sharing the same id across pages (skiptoken overlap)', async () => {
    const { http } = httpPages([
      { value: [{ id: 'a' }, { id: 'b' }], '@odata.nextLink': `${BASE}/next` },
      { value: [{ id: 'b' }, { id: 'c' }] },
    ]);
    const res = await graphRequestAll<{ id: string }>(
      '/me/todo/lists/L/tasks',
      'tok',
      0,
      http,
    );
    expect(res.map((x) => x.id)).toEqual(['a', 'b', 'c']);
  });

  it('dedup does not drop items without a string id', async () => {
    const { http } = httpPages([{ value: [1, 2, 3] }]);
    expect(await graphRequestAll<number>('/x', 'tok', 0, http)).toEqual([1, 2, 3]);
  });
});

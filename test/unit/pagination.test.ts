import { describe, it, expect, vi } from 'vitest';
import { graphRequestAll } from '../../src/graph';

// graphRequestAll → graphRequest → PluginAPI.executeNodeScript。这里直接脚本化
// executeNodeScript 的返回值，端到端覆盖：分页循环跟随 nextLink、limit 截断、
// 绝对 nextLink 直连、以及桥层报错的透传。

const BASE = 'https://graph.microsoft.com/v1.0';

interface Page {
  value: unknown[];
  '@odata.nextLink'?: string;
}

interface Call {
  method: string;
  url: string;
}

function scriptPages(pages: Page[]): { calls: Call[]; fn: (req: unknown) => Promise<unknown> } {
  const calls: Call[] = [];
  const fn = async (req: { script: string; args: unknown[] }) => {
    const page = pages[calls.length];
    if (!page) throw new Error('unexpected extra executeNodeScript call');
    const [method, url] = req.args as [string, string];
    calls.push({ method, url });
    return { success: true, result: page };
  };
  return { calls, fn };
}

function stubApi(executeNodeScript: (req: unknown) => Promise<unknown>) {
  vi.stubGlobal('PluginAPI', {
    getCurrentLanguage: () => 'en',
    executeNodeScript,
  });
}

describe('graphRequestAll', () => {
  it('follows @odata.nextLink across pages', async () => {
    const { fn } = scriptPages([
      { value: [1, 2], '@odata.nextLink': `${BASE}/next?$skiptoken=abc` },
      { value: [3] },
    ]);
    stubApi(fn);
    expect(await graphRequestAll<number>('/me/todo/lists/L/tasks', 'tok')).toEqual([1, 2, 3]);
  });

  it('passes the absolute nextLink URL through unchanged', async () => {
    const { calls, fn } = scriptPages([
      { value: [1], '@odata.nextLink': `${BASE}/next?$skiptoken=abc` },
      { value: [2] },
    ]);
    stubApi(fn);
    await graphRequestAll<number>('/me/todo/lists/L/tasks', 'tok');
    expect(calls[0].url).toBe(`${BASE}/me/todo/lists/L/tasks`);
    expect(calls[1].url).toBe(`${BASE}/next?$skiptoken=abc`);
    for (const c of calls) expect(c.method).toBe('GET');
  });

  it('stops once the limit is reached (no extra page fetched)', async () => {
    const { calls, fn } = scriptPages([
      { value: [1, 2, 3, 4, 5], '@odata.nextLink': `${BASE}/next` },
    ]);
    stubApi(fn);
    expect(await graphRequestAll<number>('/me/todo/lists/L/tasks', 'tok', 3)).toEqual([1, 2, 3]);
    expect(calls.length).toBe(1); // 未再追 nextLink
  });

  it('returns empty for an empty page', async () => {
    const { fn } = scriptPages([{ value: [] }]);
    stubApi(fn);
    expect(await graphRequestAll<number>('/me/todo/lists/L/tasks', 'tok')).toEqual([]);
  });

  it('limit 0 means no limit (pulls all pages)', async () => {
    const { fn } = scriptPages([
      { value: [1], '@odata.nextLink': `${BASE}/next` },
      { value: [2] },
    ]);
    stubApi(fn);
    expect(await graphRequestAll<number>('/me/todo/lists/L/tasks', 'tok', 0)).toEqual([1, 2]);
  });

  it('throws when the bridge reports success:false', async () => {
    stubApi(async () => ({ success: false, error: 'boom' }));
    await expect(graphRequestAll<number>('/x', 'tok')).rejects.toThrow('boom');
  });
});

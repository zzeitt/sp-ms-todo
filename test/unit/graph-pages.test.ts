import { describe, expect, it, vi } from 'vitest';
import { graphRequestPagesUntil } from '../../src/graph';
import type { PluginHttp } from '../../src/types';

const U = (p: string): string => 'https://graph.microsoft.com/v1.0/' + p;
const page = (ids: string[], next?: string) => ({
  value: ids.map((id) => ({ id })),
  ...(next ? { '@odata.nextLink': next } : {}),
});

const httpOf = (pages: Record<string, unknown>) =>
  ({
    request: vi.fn(async (_method: string, url: string) => {
      const key = Object.keys(pages).find((p) => url.startsWith(p));
      if (!key) throw new Error('unexpected url ' + url);
      return pages[key];
    }),
  }) as unknown as PluginHttp;

describe('graphRequestPagesUntil (getByIds pagination)', () => {
  it('stops as soon as every requested id is covered', async () => {
    const http = httpOf({
      [U('list?page=1')]: page(['a', 'b'], U('list?page=2')),
      [U('list?page=2')]: page(['c', 'd'], U('list?page=3')),
      [U('list?page=3')]: page(['e']),
    });
    const items = await graphRequestPagesUntil<{ id: string }>(
      U('list?page=1'),
      'token',
      http,
      (all) => new Set(all.map((i) => i.id)).has('c'),
    );
    expect(items.map((i) => i.id)).toEqual(['a', 'b', 'c', 'd']);
    expect((http.request as unknown as { mock: { calls: unknown[][] } }).mock.calls.length).toBe(2);
  });

  it('walks to the end when the requested id is missing', async () => {
    const http = httpOf({
      [U('p1')]: page(['a'], U('p2')),
      [U('p2')]: page(['b']),
    });
    const items = await graphRequestPagesUntil<{ id: string }>(U('p1'), 't', http, (all) =>
      all.some((i) => i.id === 'zzz'),
    );
    expect(items.map((i) => i.id)).toEqual(['a', 'b']);
  });

  it('dedupes ids that repeat across pages (skiptoken overlap)', async () => {
    const http = httpOf({
      [U('p1')]: page(['a', 'b'], U('p2')),
      [U('p2')]: page(['b', 'c']),
    });
    const items = await graphRequestPagesUntil<{ id: string }>(U('p1'), 't', http, () => false);
    expect(items.map((i) => i.id)).toEqual(['a', 'b', 'c']);
  });

  it('stops at maxPages even when the target is never found', async () => {
    const http = httpOf({
      [U('p1')]: page(['a'], U('p2')),
      [U('p2')]: page(['b'], U('p3')),
      [U('p3')]: page(['c'], U('p4')),
    });
    const items = await graphRequestPagesUntil<{ id: string }>(U('p1'), 't', http, () => false, 2);
    expect(items.map((i) => i.id)).toEqual(['a', 'b']);
  });
});

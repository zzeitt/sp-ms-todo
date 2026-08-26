import { describe, it, expect } from 'vitest';
import { graphRequest } from '../../src/graph';
import type { PluginHttp } from '../../src/types';

// graphRequest 全程走宿主 PluginHttp（桌面 / 移动 / Web 三端一致）。这里校验请求
// 头（Authorization / Accept / Content-Type）、body 透传、以及缺通道/报错的处理。

const BASE = 'https://graph.microsoft.com/v1.0';

interface Call {
  method: string;
  url: string;
  body: unknown;
  options: unknown;
}

function makeHttp() {
  const calls: Call[] = [];
  const http = {
    request: async <T,>(method: string, url: string, body?: unknown, options?: unknown) => {
      calls.push({ method, url, body, options });
      return { value: 'ok' } as unknown as T;
    },
  } as unknown as PluginHttp;
  return { calls, http };
}

function headersOf(call: Call): Record<string, string> {
  return (call.options as { headers: Record<string, string> }).headers;
}

describe('graphRequest via PluginHttp', () => {
  it('sends GET with Authorization + Accept headers', async () => {
    const { calls, http } = makeHttp();
    await graphRequest('GET', '/me/todo/lists', 'tok', undefined, http);
    expect(calls.length).toBe(1);
    expect(calls[0].method).toBe('GET');
    expect(calls[0].url).toBe(`${BASE}/me/todo/lists`);
    expect(headersOf(calls[0]).Authorization).toBe('Bearer tok');
    expect(headersOf(calls[0]).Accept).toBe('application/json');
  });

  it('adds Content-Type and forwards the body for write requests', async () => {
    const { calls, http } = makeHttp();
    await graphRequest('POST', '/me/todo/lists', 'tok', { title: 'x' }, http);
    expect(headersOf(calls[0])['Content-Type']).toBe('application/json');
    expect(calls[0].body).toEqual({ title: 'x' });
  });

  it('passes absolute (nextLink) URLs through unchanged', async () => {
    const { calls, http } = makeHttp();
    await graphRequest('GET', `${BASE}/next?$skiptoken=abc`, 'tok', undefined, http);
    expect(calls[0].url).toBe(`${BASE}/next?$skiptoken=abc`);
  });

  it('throws when no http channel is provided', async () => {
    await expect(graphRequest('GET', '/x', 'tok')).rejects.toThrow();
  });

  it('wraps http.request errors with a Graph-specific message', async () => {
    const http = {
      request: async () => {
        throw new Error('boom');
      },
    } as unknown as PluginHttp;
    await expect(graphRequest('GET', '/x', 'tok', undefined, http)).rejects.toThrow('boom');
  });

  it('preserves the HTTP status on the thrown error so the host can detect 404/410', async () => {
    const http = {
      request: async () => {
        throw Object.assign(new Error('not found'), { status: 404 });
      },
    } as unknown as PluginHttp;
    await expect(
      graphRequest('GET', '/me/todo/lists/l1/tasks/t1', 'tok', undefined, http),
    ).rejects.toMatchObject({ status: 404 });
  });
});

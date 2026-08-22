import type { PluginHttp } from '../../../src/types';

// 在 Node 测试环境里用内置 fetch 实现宿主 PluginHttp，让集成测试走
// graphRequest → http.request 的真实代码路径，对 Graph 做一次真实 HTTPS 往返
// （Node 18+ 自带全局 fetch，测试环境为 node）。
export function makeFetchHttp(): PluginHttp {
  const request: PluginHttp['request'] = async (method, url, body, options) => {
    const headers: Record<string, string> = {
      ...(options?.headers || {}),
      'User-Agent': 'super-productivity-ms-todo-integration-test',
    };
    const res = await fetch(url, {
      method,
      headers,
      body: body === undefined || body === null ? undefined : JSON.stringify(body),
    });
    const text = await res.text();
    const data = text ? JSON.parse(text) : undefined;
    if (!res.ok) {
      throw new Error(`HTTP ${res.status}: ${text.slice(0, 500)}`);
    }
    return data as unknown;
  };

  return {
    get: (url, options) => request('GET', url, undefined, options),
    post: (url, body, options) => request('POST', url, body, options),
    put: (url, body, options) => request('PUT', url, body, options),
    patch: (url, body, options) => request('PATCH', url, body, options),
    delete: (url, options) => request('DELETE', url, undefined, options),
    request,
  };
}

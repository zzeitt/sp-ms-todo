// 用 PluginAPI.request 构造一个与 PluginHttp 同型的 HTTP 通道。
//
// 背景：issueProvider 的方法签名里才会注入 `http`（宿主每调用一次就新建一个
// PluginHttp helper）。侧边栏 / 顶部按钮的回调拿不到这个 `http`，之前只能靠
// `cacheContext` 缓存的最近一次 issueProvider 调用值，导致「缺少上下文」。
//
// 宿主其实把 `PluginAPI.request` 作为一等 API 暴露（底层与 issueProvider 的
// http 是同一个 PluginHttpService），只要 manifest 声明了 `http` 权限和
// `allowedHosts`（graph.microsoft.com / login.microsoftonline.com）就能直接发
// 请求。这里把它包装成 PluginHttp 形状，按钮即可不依赖缓存上下文地访问 Graph。
//
// 注意：`PluginAPI.request` 底层在桌面端走 `fetch(..., redirect:'error')`，移动端
// 写请求走 Capacitor 原生 HTTP，因此和 issueProvider 的 http 一样，桌面端写请求
// 仍可能受宿主 Electron CORS 预检影响（status=0）——这是宿主层面限制。

import type { PluginAPI as PluginApiType, PluginHttp, PluginHttpOptions } from './types';

declare const PluginAPI: PluginApiType;

/** 是否提供 PluginAPI.request（旧版 SP 可能没有，此时回退到缓存的 issueProvider http）。 */
export function hasRequestApi(): boolean {
  return typeof PluginAPI.request === 'function';
}

export function createRequestHttp(): PluginHttp {
  const req = <T = unknown>(
    method: string,
    url: string,
    body?: unknown,
    options?: PluginHttpOptions,
  ): Promise<T> =>
    PluginAPI.request(url, {
      method,
      body,
      params: options?.params,
      headers: options?.headers,
      timeout: options?.timeout,
      responseType: options?.responseType,
    }) as Promise<T>;

  return {
    get: (url, options) => req('GET', url, undefined, options),
    post: (url, body, options) => req('POST', url, body, options),
    put: (url, body, options) => req('PUT', url, body, options),
    patch: (url, body, options) => req('PATCH', url, body, options),
    delete: (url, options) => req('DELETE', url, undefined, options),
    request: (method, url, body, options) => req(method, url, body, options),
  };
}

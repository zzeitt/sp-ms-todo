// Microsoft Graph HTTP 客户端。
//
// graph.microsoft.com 对跨源请求返回完整的 CORS 头（Access-Control-Allow-Origin: *，
// 预检放行 authorization / content-type），所以渲染进程里的普通浏览器请求即可直达。
// 本插件全程走宿主注入的 PluginHttp（`http` 权限，底层是 Angular HttpClient→XHR，
// 桌面 Electron / 原生移动端 / Web 三端一致）——这与 GitHub / Azure DevOps / Trello /
// Linear / Gitea / ClickUp 等官方 issue 插件相同，它们都依赖 Graph 这类返回 CORS 的
// 公开 API。
//
// 不再使用 executeNodeScript（nodeExecution）：该权限是「高级权限」，SP 只在桌面
// Electron 上授权，移动端 / Web 端会直接把插件拒载（报 NODE_ONLY_DESKTOP）。而
// Graph 支持 CORS，PluginHttp 已足够，无需绕道 Node 子进程。

import type { PluginHttp } from './types';
import { t } from './i18n';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export type GraphMethod = 'GET' | 'POST' | 'PATCH' | 'DELETE';

/** 通过宿主 PluginHttp 执行一次 Graph 请求（桌面 / 移动 / Web 三端通用）。 */
export async function graphRequest<T>(
  method: GraphMethod,
  path: string,
  token: string,
  body?: unknown,
  http?: PluginHttp,
): Promise<T> {
  const url = path.startsWith('http') ? path : GRAPH_BASE + path;

  if (!http) {
    throw new Error(
      t({
        zh: '当前平台缺少 HTTP 通道，无法访问 Microsoft Graph',
        en: 'No HTTP channel is available on this platform, so Microsoft Graph cannot be reached.',
      }),
    );
  }
  const headers: Record<string, string> = {
    Authorization: 'Bearer ' + token,
    Accept: 'application/json',
  };
  if (body !== undefined && body !== null) {
    headers['Content-Type'] = 'application/json';
  }
  try {
    return await http.request<T>(method, url, body, { headers });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new Error(t({ zh: 'Graph 请求失败: ', en: 'Graph request failed: ' }) + msg);
  }
}

/** OData 分页列表响应（`value` + 可选下一页游标）。 */
export interface ODataList<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

/**
 * 遍历分页拉取条目，直到没有下一页或达到 `limit`。
 *
 * Graph 的集合接口（如 /me/todo/lists/{id}/tasks）默认每页只返回 50 条，其余
 * 结果放在响应里的 `@odata.nextLink`（带 $skiptoken 的完整 URL）。这里循环跟随
 * nextLink 直到末页；`limit` 为 0 时不限制，>0 时拉满 `limit` 条即停止，避免超大
 * 列表把渲染进程拖卡或占满本地存储。
 */
export async function graphRequestAll<T>(
  path: string,
  token: string,
  limit = 0,
  http?: PluginHttp,
): Promise<T[]> {
  let url: string | undefined = path;
  const items: T[] = [];
  while (url) {
    const res = (await graphRequest('GET', url, token, undefined, http)) as ODataList<T>;
    items.push(...res.value);
    if (limit > 0 && items.length >= limit) {
      return items.slice(0, limit);
    }
    url = res['@odata.nextLink'];
  }
  return items;
}


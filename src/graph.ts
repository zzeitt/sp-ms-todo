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

/** 从宿主 HTTP 错误里挖出可读信息（status + Graph 错误体），避免只显示 [object Object]。 */
function describeGraphError(e: unknown): string {
  const err = e as {
    status?: number;
    statusText?: string;
    message?: string;
    error?: {
      code?: string;
      message?: string;
      error_description?: string;
      error?: { code?: string; message?: string };
    };
  } | null | undefined;
  if (!err || typeof err !== 'object') {
    return e instanceof Error ? e.message : String(e);
  }
  const parts: string[] = [];
  if (typeof err.status === 'number') parts.push(`status=${err.status}`);
  if (err.statusText) parts.push(err.statusText);
  if (typeof err.message === 'string' && err.message) parts.push(err.message);
  const inner = err.error?.error ?? err.error;
  if (inner?.code) parts.push(`code=${inner.code}`);
  if (inner?.message) parts.push(inner.message);
  if (err.error?.error_description) parts.push(err.error.error_description);
  return parts.length ? parts.join(' — ') : String(e);
}

// Graph 在部分网络环境下较慢，SP PluginHttp 默认 30s 超时太容易误杀。
const GRAPH_TIMEOUT_MS = 60_000;
// 429/网络抖动在大量任务轮询时很常见；多给几次退避，避免一次限流直接丢掉本轮
// 所有远端更新（宿主 getByIds 的 catch 会丢弃整批）。
const GRAPH_RETRIES = 3;
const GRAPH_RETRY_DELAYS_MS = [1000, 3000, 7000];

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** 判断是否值得重试：超时 / 429 / 5xx / 408 / status=0（宿主网络层失败）/ 网络错误。 */
function isRetryableGraphError(e: unknown): boolean {
  const status = (e as { status?: number } | null | undefined)?.status;
  const msg = e instanceof Error ? e.message : String(e);
  return (
    (typeof status === 'number' && (status === 0 || status === 408 || status === 429 || status >= 500)) ||
    /timeout|timed out|network|ERR_HTTP2|ERR_CONNECTION|ERR_NETWORK|socket hang up/i.test(msg)
  );
}

/** 幂等写成（PATCH 同体重复执行结果相同）可安全重试；POST 不重试，避免重复创建。 */
const RETRYABLE_METHODS: ReadonlySet<GraphMethod> = new Set<GraphMethod>(['GET', 'DELETE', 'PATCH']);

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
  const options = { headers, timeout: GRAPH_TIMEOUT_MS };
  const request = (): Promise<T> => http.request<T>(method, url, body, options);

  try {
    return await request();
  } catch (e) {
    // GET / DELETE / PATCH 幂等，超时/网络错误时重试；POST 不重试，避免重复创建。
    if (RETRYABLE_METHODS.has(method) && isRetryableGraphError(e)) {
      for (let attempt = 1; attempt <= GRAPH_RETRIES; attempt++) {
        await sleep(GRAPH_RETRY_DELAYS_MS[Math.min(attempt - 1, GRAPH_RETRY_DELAYS_MS.length - 1)]);
        try {
          return await request();
        } catch (retryErr) {
          e = retryErr;
        }
      }
    }
    // 保留宿主 PluginHttp 提供的机器可读 status（404/410 = 远端已删除等），
    // 否则宿主 two-way-sync 只能用字符串去猜「该项已不存在」还是别的错误。
    const status = (e as { status?: unknown } | null | undefined)?.status;
    const err = new Error(
      t({ zh: 'Graph 请求失败: ', en: 'Graph request failed: ' }) +
        describeGraphError(e) +
        ` [${method} ${url}]`,
    );
    if (typeof status === 'number') {
      (err as { status?: number }).status = status;
    }
    throw err;
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
  const seenIds = new Set<string>();
  while (url) {
    const res = (await graphRequest('GET', url, token, undefined, http)) as ODataList<T>;
    for (const item of res.value) {
      // Graph 的 $skiptoken 分页在列表变化时可能跨页返回同一条目（快照重叠），
      // 这里按 id 去重，避免同步出重复任务/列表。
      const id = (item as { id?: unknown } | null | undefined)?.id;
      if (typeof id === 'string' && id) {
        if (seenIds.has(id)) {
          continue;
        }
        seenIds.add(id);
      }
      items.push(item);
    }
    if (limit > 0 && items.length >= limit) {
      return items.slice(0, limit);
    }
    url = res['@odata.nextLink'];
  }
  return items;
}


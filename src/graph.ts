// Microsoft Graph HTTP 客户端。
//
// graph.microsoft.com 不返回 CORS 头，渲染进程的 fetch（PluginHttp /
// PluginAPI.request）无法访问它。插件唯一可绕过 CORS 的通道是
// `executeNodeScript`（Electron 桌面端专属），它在 spawn 出来的 Node 子进程里
// 以完整网络权限执行脚本。
//
// 内嵌脚本【必须】用 `require('https')` 而不是 `fetch`：宿主只有在脚本里出现
// fs/path/os 之外的 require() 时才会判定为「复杂脚本」走 spawn 路径；纯 fetch
// 会被判为「可直接执行」，落到只有 fs/path/os、没有网络的 vm 沙箱里失败。

import type { PluginAPI as PluginApiType } from './types';
import { t } from './i18n';

declare const PluginAPI: PluginApiType;

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// 注意：spawn 包装只把脚本的返回值通过 `console.log(JSON.stringify({__result}))`
// 回传，并整体解析 stdout。所以脚本内【绝不能 console.log】（会污染结果通道），
// 需要调试信息时用 console.error（走 stderr）。
export const NODE_SCRIPT = `
const https = require('https');
const method = args[0];
const url = args[1];
const token = args[2];
const body = args[3];

const u = new URL(url);
const hasBody = body !== undefined && body !== null;
const payload = hasBody ? JSON.stringify(body) : null;

const result = await new Promise((resolve, reject) => {
  const req = https.request({
    hostname: u.hostname,
    port: u.port || 443,
    path: u.pathname + u.search,
    method: method,
    headers: {
      'Authorization': 'Bearer ' + token,
      'Accept': 'application/json',
      'User-Agent': 'super-productivity-ms-todo/0.1',
      ...(hasBody ? {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      } : {}),
    },
  }, (res) => {
    let raw = '';
    res.setEncoding('utf8');
    res.on('data', (c) => { raw += c; });
    res.on('end', () => {
      let parsed;
      if (raw) {
        try { parsed = JSON.parse(raw); } catch (e) { parsed = raw; }
      }
      if (res.statusCode >= 200 && res.statusCode < 300) {
        resolve(parsed);
      } else {
        reject(new Error('HTTP ' + res.statusCode + ': ' + raw.slice(0, 500)));
      }
    });
  });
  req.on('error', (e) => reject(new Error('network: ' + e.message)));
  if (payload !== null) req.write(payload);
  req.end();
});

return result;
`;

export type GraphMethod = 'GET' | 'POST' | 'PATCH' | 'DELETE';

/** 通过 nodeExecution 桥执行一次 Graph 请求。 */
export async function graphRequest<T>(
  method: GraphMethod,
  path: string,
  token: string,
  body?: unknown,
): Promise<T> {
  if (!PluginAPI.executeNodeScript) {
    throw new Error(
      t({
        zh: '此插件需要桌面端（Electron）的 Node 执行能力，Web 端不可用',
        en: 'This plugin needs the Node execution capability of the desktop (Electron) app; the web version is not supported.',
      }),
    );
  }
  const url = path.startsWith('http') ? path : GRAPH_BASE + path;
  const res = await PluginAPI.executeNodeScript({
    script: NODE_SCRIPT,
    args: [method, url, token, body],
  });
  if (!res.success) {
    const msg = typeof res.error === 'string' ? res.error : res.error?.message;
    throw new Error(t({ zh: 'Graph 请求失败: ', en: 'Graph request failed: ' }) + msg);
  }
  return res.result as T;
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
): Promise<T[]> {
  let url: string | undefined = path;
  const items: T[] = [];
  while (url) {
    const res = (await graphRequest('GET', url, token)) as ODataList<T>;
    items.push(...res.value);
    if (limit > 0 && items.length >= limit) {
      return items.slice(0, limit);
    }
    url = res['@odata.nextLink'];
  }
  return items;
}

/** 取一个有效的 Graph 访问令牌，未连接时抛错误。 */
export async function requireToken(): Promise<string> {
  const token = await PluginAPI.getOAuthToken();
  if (!token) {
    throw new Error(
      t({
        zh: '未连接 Microsoft 账户。请先在插件配置中点击「连接 Microsoft 账户」。',
        en: 'Not connected to a Microsoft account. Click "Connect Microsoft account" in the plugin settings first.',
      }),
    );
  }
  return token;
}

// Microsoft To Do 认证：OAuth 2.0 设备码流（Device Authorization Grant, RFC 8628）。
//
// 为什么用设备码流：SP 插件 OAuth 宿主在原生端生成的 redirect URI 是单斜杠
// scheme（com.super-productivity.app:/plugin-oauth-callback），Microsoft Entra
// 因缺少 host 拒绝登记该 URI，导致移动端无法用 oauthButton 登录。设备码流不依赖
// redirect URI —— App 主动轮询 token 端点，浏览器永不跳回 App —— 因而桌面 +
// Android + iOS 三端都能登录（Web 端因 token 端点无 CORS 仍不可用，与 SP 其它
// provider 一致）。
//
// 纯逻辑（设备码请求 / token 轮询 / refresh 换取 / 密钥存取 / 登录编排），依赖
// （HTTP 通道、密钥存储、对话框）全部注入，便于单元测试。底部的
// createSecretStore / createDeviceCodeDialog 才接宿主 PluginAPI 全局对象。

import type { PluginAPI as PluginApiType, PluginHttp } from './types';
import { t } from './i18n';

declare const PluginAPI: PluginApiType;

// --- 常量 ---

export const MS_DEVICE_CODE_URL =
  'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode';
export const MS_TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
export const SCOPES = ['Tasks.ReadWrite', 'offline_access'];

// Secret 存储键（本地、每设备、不进 sync/导出/备份）。
export const SECRET_REFRESH_TOKEN = 'ms-todo-refresh-token';
export const SECRET_ACCESS_TOKEN = 'ms-todo-access-token';
export const SECRET_ACCESS_EXPIRY = 'ms-todo-access-expiry';

// access token 提前多少毫秒视为「即将过期」而强制刷新。
const ACCESS_TOKEN_REFRESH_BUFFER_MS = 60_000;

// --- 类型 ---

export interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
}

export interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  error?: string;
  error_description?: string;
}

/** 密钥存储抽象（生产环境用 PluginAPI.getSecret/setSecret/deleteSecret 实现）。 */
export interface SecretStore {
  get(key: string): Promise<string | null>;
  set(key: string, value: string): Promise<void>;
  delete(key: string): Promise<void>;
}

/** 设备码展示对话框抽象。返回 true 表示用户已点「我已输入验证码」。 */
export interface DeviceCodeDialog {
  show(userCode: string, verificationUri: string): Promise<boolean>;
}

// --- 内部工具 ---

function toFormBody(params: Record<string, string>): string {
  return Object.entries(params)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&');
}

async function postForm<T>(
  http: PluginHttp,
  url: string,
  params: Record<string, string>,
): Promise<T> {
  return http.post<T>(url, toFormBody(params), {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    timeout: 60_000,
  });
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// --- 核心函数（依赖注入，可测试）---

/** 向 Graph 请求一个设备码。 */
export async function requestDeviceCode(
  clientId: string,
  http: PluginHttp,
): Promise<DeviceCodeResponse> {
  const res = await postForm<DeviceCodeResponse & { error?: string }>(
    http,
    MS_DEVICE_CODE_URL,
    { client_id: clientId, scope: SCOPES.join(' ') },
  );
  if (!res.device_code || !res.user_code || !res.verification_uri) {
    throw new Error(
      t({ zh: '获取设备码失败', en: 'Failed to obtain a device code.' }) +
        (res.error ? ` ${res.error}` : ''),
    );
  }
  return res;
}

/**
 * 轮询 token 端点，直到用户完成授权或失败。
 *
 * - `authorization_pending`：继续等；
 * - `slow_down`：按规范把下次轮询间隔再 +5s；
 * - 其它错误（access_denied / expired_token …）：直接抛错；
 * - 超过 `timeoutMs` 仍未完成：抛超时错误。
 */
export async function pollForToken(
  clientId: string,
  deviceCode: string,
  intervalSeconds: number,
  http: PluginHttp,
  timeoutMs = 900_000,
  sleepFn: (ms: number) => Promise<void> = sleep,
): Promise<TokenResponse> {
  const deadline = Date.now() + timeoutMs;
  let waitMs = Math.max(1, intervalSeconds) * 1000;
  for (;;) {
    const res = await postForm<TokenResponse>(http, MS_TOKEN_URL, {
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: clientId,
      device_code: deviceCode,
    });
    if (!res.error) {
      return res;
    }
    if (res.error === 'authorization_pending') {
      // 继续轮询
    } else if (res.error === 'slow_down') {
      waitMs += 5000;
    } else {
      throw new Error(
        t({ zh: '设备码登录失败: ', en: 'Device code sign-in failed: ' }) + res.error,
      );
    }
    if (Date.now() + waitMs > deadline) {
      throw new Error(t({ zh: '设备码登录超时', en: 'Device code sign-in timed out.' }));
    }
    await sleepFn(waitMs);
  }
}

/** 用 refresh token 换新的 access token（Microsoft 会轮换 refresh token）。 */
export async function refreshAccessToken(
  clientId: string,
  refreshToken: string,
  http: PluginHttp,
): Promise<TokenResponse> {
  const res = await postForm<TokenResponse>(http, MS_TOKEN_URL, {
    grant_type: 'refresh_token',
    client_id: clientId,
    refresh_token: refreshToken,
    scope: SCOPES.join(' '),
  });
  if (res.error || !res.access_token) {
    throw new Error(
      t({ zh: '刷新令牌失败: ', en: 'Failed to refresh the token: ' }) +
        (res.error ?? 'no access_token'),
    );
  }
  return res;
}

/** 把 token 写回密钥存储（access + 过期时间 + 可选的新 refresh token）。 */
export async function storeTokens(
  tokens: TokenResponse,
  secrets: SecretStore,
): Promise<void> {
  await secrets.set(SECRET_ACCESS_TOKEN, tokens.access_token);
  await secrets.set(
    SECRET_ACCESS_EXPIRY,
    String(Date.now() + tokens.expires_in * 1000),
  );
  if (tokens.refresh_token) {
    await secrets.set(SECRET_REFRESH_TOKEN, tokens.refresh_token);
  }
}

/** 取一个有效的 Graph access token：缓存有效则用缓存，否则用 refresh token 换。 */
export async function getAccessToken(
  clientId: string,
  http: PluginHttp,
  secrets: SecretStore,
): Promise<string> {
  const cached = await secrets.get(SECRET_ACCESS_TOKEN);
  const expiryStr = await secrets.get(SECRET_ACCESS_EXPIRY);
  if (cached && expiryStr) {
    const expiry = Number(expiryStr);
    if (Number.isFinite(expiry) && expiry - Date.now() > ACCESS_TOKEN_REFRESH_BUFFER_MS) {
      return cached;
    }
  }

  const refreshToken = await secrets.get(SECRET_REFRESH_TOKEN);
  if (!refreshToken) {
    throw new Error(
      t({
        zh:
          '未连接 Microsoft 账户。请打开插件配置，点击「测试连接」以完成设备码登录。',
        en:
          'Not connected to a Microsoft account. Open the plugin settings and click "Test connection" to sign in via device code.',
      }),
    );
  }

  const tokens = await refreshAccessToken(clientId, refreshToken, http);
  await storeTokens(tokens, secrets);
  return tokens.access_token;
}

/** 完整设备码登录编排：要码 → 弹窗 → 轮询 → 存 token。 */
export async function runDeviceCodeLogin(
  clientId: string,
  http: PluginHttp,
  dialog: DeviceCodeDialog,
  secrets: SecretStore,
): Promise<void> {
  const deviceCode = await requestDeviceCode(clientId, http);
  const confirmed = await dialog.show(deviceCode.user_code, deviceCode.verification_uri);
  if (!confirmed) {
    throw new Error(t({ zh: '已取消登录', en: 'Sign-in cancelled.' }));
  }
  const tokens = await pollForToken(
    clientId,
    deviceCode.device_code,
    deviceCode.interval,
    http,
  );
  await storeTokens(tokens, secrets);
}

// --- 默认依赖（接宿主 PluginAPI）---

export function createSecretStore(): SecretStore {
  return {
    get: (key) => PluginAPI.getSecret(key),
    set: (key, value) => PluginAPI.setSecret(key, value),
    delete: (key) => PluginAPI.deleteSecret(key),
  };
}

export function createDeviceCodeDialog(): DeviceCodeDialog {
  return {
    async show(userCode: string, verificationUri: string): Promise<boolean> {
      const okLabel = t({ zh: '我已输入验证码', en: 'I entered the code' });
      const result = await PluginAPI.openDialog({
        title: t({ zh: '连接 Microsoft 账户', en: 'Connect Microsoft account' }),
        content: t({
          zh: `请在浏览器打开 ${verificationUri}，输入代码：${userCode}`,
          en: `Open ${verificationUri} in a browser and enter the code: ${userCode}`,
        }),
        okBtnLabel: okLabel,
        cancelBtnLabel: t({ zh: '取消', en: 'Cancel' }),
      });
      return result === okLabel;
    },
  };
}

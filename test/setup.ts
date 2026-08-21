import { vi } from 'vitest';

// 纯逻辑里 i18n 的 t() 会读 PluginAPI.getCurrentLanguage()；分页单元测试会读
// PluginAPI.executeNodeScript。这里先给一个安全的默认实现，具体用例再按需覆盖。
//
// 注意：i18n.ts 用 `PluginAPI?.getCurrentLanguage?.()`，但 `PluginAPI` 是一个
// 未声明的全局标识符，可选链只会挡住 null/undefined 值、挡不住 ReferenceError，
// 因此测试里必须真的把它挂到 globalThis 上。
vi.stubGlobal('PluginAPI', {
  registerIssueProvider: () => {},
  startOAuthFlow: async () => ({ accessToken: '', refreshToken: '', expiresAt: 0 }),
  getOAuthToken: async () => 'test-token',
  clearOAuthToken: async () => {},
  getCurrentLanguage: () => 'en',
  translate: (key: string) => key,
  executeNodeScript: async () => ({ success: true, result: {} }),
  log: {
    err() {},
    error() {},
    log() {},
    info() {},
    warn() {},
    debug() {},
  },
});

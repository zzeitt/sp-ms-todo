// 插件级多语言（zh / en）。
//
// SP 不会自动翻译插件的 configFields 的 label / description：宿主在
// _mapPluginConfigField 里直接用 f.label / f.description，oauthButton 也是
// {{ btn.label }} 原文渲染。因此插件在运行时用 PluginAPI.getCurrentLanguage()
// 自行选择 zh / en 文案。
//
// 注意：configFields / issueDisplay 在 registerIssueProvider 时静态生成，切换
// SP 语言后需重载插件才会更新（官方机制同样要求插件侧自行解析）。
// SP 语言码：zh = 简体中文，en = 英文，其余一律回退英文。

declare const PluginAPI: {
  getCurrentLanguage(): string;
};

export interface L10n {
  zh: string;
  en: string;
}

/** 当前 SP 客户端语言是否为简体中文。 */
export function isZh(): boolean {
  const lang = (PluginAPI?.getCurrentLanguage?.() ?? 'en').toLowerCase();
  return lang.startsWith('zh');
}

/** 按当前语言取文案，非中文一律回退英文。 */
export function t(v: L10n): string {
  return isZh() ? v.zh : v.en;
}

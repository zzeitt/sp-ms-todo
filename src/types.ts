// 自包含的 @super-productivity/plugin-api 子集类型。
//
// 官方包是 monorepo 的 workspace file: 依赖，未发布到 npm；本插件独立构建，
// 故把这些接口按源码逐字 vendor 进来（字段名/签名与真实 PluginAPI 完全一致，
// 编译期只做结构校验，运行时由宿主注入真实的 PluginAPI 全局对象）。

export interface OAuthFlowConfig {
  authUrl: string;
  tokenUrl: string;
  clientId: string;
  clientSecret?: string;
  mobileClientId?: string;
  iosClientId?: string;
  webClientId?: string;
  scopes: string[];
  extraAuthParams?: Record<string, string>;
  redirectUri?: string;
}

export interface PluginSearchResult {
  id: string;
  title: string;
  url?: string;
  status?: string;
  assignee?: string;
  labels?: string[];
  start?: number;
  dueWithTime?: number;
  duration?: number;
  isAllDay?: boolean;
  description?: string;
  [key: string]: unknown;
}

export interface PluginIssueComment {
  author: string;
  body: string;
  created: number;
  [key: string]: unknown;
}

export interface PluginIssue {
  id: string;
  title: string;
  body?: string;
  url?: string;
  state?: string;
  lastUpdated?: number;
  assignee?: string;
  labels?: string[];
  comments?: PluginIssueComment[];
  [key: string]: unknown;
}

export interface PluginIssueField {
  field: string;
  label: string;
  type?: 'text' | 'markdown' | 'link' | 'date' | 'list';
  linkField?: string;
  hideEmpty?: boolean;
}

export type PluginSyncDirection = 'off' | 'pullOnly' | 'pushOnly' | 'both';

export interface PluginFieldMapping {
  taskField:
    | 'isDone'
    | 'title'
    | 'notes'
    | 'dueDay'
    | 'dueWithTime'
    | 'timeEstimate'
    | 'tagIds';
  issueField: string;
  defaultDirection: PluginSyncDirection;
  mutuallyExclusive?: string[];
  toIssueValue(
    taskValue: unknown,
    ctx: { issueId: string; issueNumber?: number },
  ): unknown;
  toTaskValue(
    issueValue: unknown,
    ctx: { issueId: string; issueNumber?: number },
  ): unknown;
}

export interface PluginHttpOptions {
  params?: Record<string, string>;
  headers?: Record<string, string>;
  timeout?: number;
  responseType?: 'json' | 'text';
}

// 宿主注入的全平台 HTTP 通道（`http` 权限）。底层是 Angular HttpClient→XHR，桌面
// Electron / 原生移动端 / Web 三端一致。Graph 返回完整 CORS 头，故此通道可直接
// 访问 graph.microsoft.com，无需 Node 子进程。
export interface PluginHttp {
  get<T = unknown>(url: string, options?: PluginHttpOptions): Promise<T>;
  post<T = unknown>(url: string, body: unknown, options?: PluginHttpOptions): Promise<T>;
  put<T = unknown>(url: string, body: unknown, options?: PluginHttpOptions): Promise<T>;
  patch<T = unknown>(url: string, body: unknown, options?: PluginHttpOptions): Promise<T>;
  delete<T = unknown>(url: string, options?: PluginHttpOptions): Promise<T>;
  request<T = unknown>(
    method: string,
    url: string,
    body?: unknown,
    options?: PluginHttpOptions,
  ): Promise<T>;
}

export interface PluginFormField {
  key: string;
  type:
    | 'input'
    | 'password'
    | 'textarea'
    | 'checkbox'
    | 'select'
    | 'multiSelect'
    | 'link'
    | 'oauthButton';
  label: string;
  required?: boolean;
  description?: string;
  options?: { label: string; value: string }[];
  url?: string;
  pattern?: string;
  advanced?: boolean;
  showIf?: string;
  oauthConfig?: OAuthFlowConfig;
  loadOptions?(
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<{ label: string; value: string }[]>;
}

export interface IssueProviderPluginDefinition {
  configFields: PluginFormField[];
  getHeaders(
    config: Record<string, unknown>,
  ): Record<string, string> | Promise<Record<string, string>>;
  searchIssues(
    searchTerm: string,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginSearchResult[]>;
  getById(
    issueId: string,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginIssue>;
  getIssueLink(issueId: string, config: Record<string, unknown>): string;
  testConnection?(config: Record<string, unknown>, http: PluginHttp): Promise<boolean>;
  getNewIssuesForBacklog?(
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginSearchResult[]>;
  issueDisplay: PluginIssueField[];
  fieldMappings?: PluginFieldMapping[];
  updateIssue?(
    id: string,
    changes: Record<string, unknown>,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<void>;
  createIssue?(
    title: string,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<{ issueId: string; issueNumber?: number; issueData: PluginIssue }>;
  extractSyncValues?(issue: PluginIssue): Record<string, unknown>;
  deleteIssue?(
    id: string,
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<void>;
  deletedStates?: string[];
}

export interface PluginAPI {
  registerIssueProvider(definition: IssueProviderPluginDefinition): void;
  startOAuthFlow(config: OAuthFlowConfig): Promise<{
    accessToken: string;
    refreshToken: string;
    expiresAt: number;
  }>;
  getOAuthToken(): Promise<string | null>;
  clearOAuthToken(): Promise<void>;
  getCurrentLanguage(): string;
  translate(key: string, params?: Record<string, string | number>): string;
  log: {
    err: (...args: unknown[]) => void;
    error: (...args: unknown[]) => void;
    log: (...args: unknown[]) => void;
    info: (...args: unknown[]) => void;
    warn: (...args: unknown[]) => void;
    debug: (...args: unknown[]) => void;
  };
}

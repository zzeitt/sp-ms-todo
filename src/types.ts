// 自包含的 @super-productivity/plugin-api 子集类型。
//
// 官方包是 monorepo 的 workspace file: 依赖，未发布到 npm；本插件独立构建，
// 故把这些接口按源码逐字 vendor 进来（字段名/签名与真实 PluginAPI 完全一致，
// 编译期只做结构校验，运行时由宿主注入真实的 PluginAPI 全局对象）。

export interface DialogCfg {
  title?: string;
  /** 富文本，宿主做白名单消毒后渲染。 */
  htmlContent?: string;
  /** 纯文本内容。 */
  content?: string;
  okBtnLabel?: string;
  cancelBtnLabel?: string;
  buttons?: { label: string; onClick?: () => void | Promise<void> }[];
}

export type DialogResult = string | undefined;

export interface SnackCfg {
  msg: string;
  type?: 'SUCCESS' | 'ERROR' | 'WARNING' | 'INFO';
  ico?: string;
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
    | 'link';
  label: string;
  required?: boolean;
  description?: string;
  options?: { label: string; value: string }[];
  url?: string;
  pattern?: string;
  advanced?: boolean;
  showIf?: string;
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
  /**
   * 批量版 getById（宿主刷新大量任务时使用）。缺失的 id 会被跳过、不当作已删除。
   * 实现后宿主一轮刷新只需 O(列表数) 次请求，而不是 O(任务数)。
   */
  getByIds?(
    issueIds: string[],
    config: Record<string, unknown>,
    http: PluginHttp,
  ): Promise<PluginIssue[]>;
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
  setSecret(key: string, value: string): Promise<void>;
  getSecret(key: string): Promise<string | null>;
  deleteSecret(key: string): Promise<void>;
  openDialog(dialogCfg: DialogCfg): Promise<DialogResult>;
  showSnack(snackCfg: SnackCfg): void;
  getCurrentLanguage(): string;
  translate(key: string, params?: Record<string, string | number>): string;
  registerSidePanelButton(button: SidePanelButtonCfg): void;
  registerHeaderButton(button: HeaderButtonCfg): void;
  registerWorkContextHeaderButton(button: WorkContextHeaderButtonCfg): void;
  registerMenuEntry(entry: MenuEntryCfg): void;
  showIndexHtmlAsView(): void;
  /** 宿主在插件加载完成后触发（对齐宿主 PluginRunner 的 onReady 回调）。 */
  onReady(handler: () => void): void;
  /** 注册「iframe 视图 ↔ plugin.js」桥接消息处理器（宿主经 PluginService.sendMessageToPlugin 送达）。 */
  onMessage(handler: (message: unknown) => unknown | Promise<unknown>): void;
  request(url: string, options?: PluginRequestOptions): Promise<unknown>;
  getSelectedTask(): Promise<SpTask | null>;
  getFocusedTask(): Promise<SpTask | null>;
  getTasks(): Promise<SpTask[]>;
  /** 已完成任务会被移出 getTasks()，只在这里可见（宿主 19.x 起提供）。 */
  getArchivedTasks?(): Promise<SpTask[]>;
  getAllProjects(): Promise<SpProject[]>;
  getAllTags(): Promise<SpTag[]>;
  getActiveWorkContext(): Promise<WorkContext | null>;
  updateTask(taskId: string, changes: Record<string, unknown>): Promise<unknown>;
  /**
   * 注册宿主 hook（taskUpdate / taskComplete 等）。用于在用户打勾后由插件
   * 自己兜底同步完成状态（宿主推送失败时不会重试）。
   */
  registerHook(
    hook: 'taskUpdate' | 'taskComplete',
    handler: (payload: PluginHookPayload) => void | Promise<void>,
  ): void;
  getConfig(): Promise<Record<string, unknown> | null>;
  log: {
    err: (...args: unknown[]) => void;
    error: (...args: unknown[]) => void;
    log: (...args: unknown[]) => void;
    info: (...args: unknown[]) => void;
    warn: (...args: unknown[]) => void;
    debug: (...args: unknown[]) => void;
  };
}

// --- 手动「强制同步」按钮所需的宿主类型 ---

/** 宿主 hook（taskUpdate）投递的载荷。 */
export interface PluginHookPayload {
  taskId?: string;
  task?: SpTask;
  changes?: Record<string, unknown>;
}

/** 宿主 PluginAPI.getTasks / getSelectedTask 返回的 SP 任务（精简字段）。 */
export interface SpTask {
  id: string;
  title: string;
  notes?: string;
  issueId?: string;
  issueProviderId?: string;
  issueType?: string;
  projectId?: string;
  tagIds?: string[];
  isDone?: boolean;
  dueDay?: string;
  dueWithTime?: number;
  /** 上次同步完成时记录的 issue 字段快照（state/title/body/dueDate/dueWithTime）。 */
  issueLastSyncedValues?: Record<string, unknown>;
  /** 上次同步时远端 issue 的 lastModified（毫秒）。 */
  issueLastUpdated?: number;
  /** 本地完成时间（毫秒）。 */
  doneOn?: number;
  modified?: number;
  [key: string]: unknown;
}

/** 宿主 PluginAPI.getAllProjects 返回的项目（精简字段）。 */
export interface SpProject {
  id: string;
  title: string;
  [key: string]: unknown;
}

/** 宿主 PluginAPI.getAllTags 返回的标签（精简字段）。 */
export interface SpTag {
  id: string;
  title: string;
  color?: string;
  [key: string]: unknown;
}

export interface SidePanelButtonCfg {
  label: string;
  icon?: string;
  onClick: () => void | Promise<void>;
  [key: string]: unknown;
}

/** 宿主 PluginAPI.request 的选项（与 PluginHttpOptions 对齐，外加 method/body）。 */
export interface PluginRequestOptions {
  method?: string;
  body?: unknown;
  params?: Record<string, string>;
  headers?: Record<string, string>;
  timeout?: number;
  responseType?: 'json' | 'text';
}

/** 顶部全局工具栏按钮（onClick 无参数，宿主不会顺带打开插件面板）。 */
export interface HeaderButtonCfg {
  label: string;
  icon?: string;
  color?: 'primary' | 'warn' | 'accent';
  onClick: () => void | Promise<void>;
  [key: string]: unknown;
}

/** 左侧菜单入口（打开插件视图 / index.html）。 */
export interface MenuEntryCfg {
  label: string;
  icon?: string;
  onClick: () => void | Promise<void>;
  [key: string]: unknown;
}

/** 工作上下文（项目 / 标签 / 今天）头部的按钮。 */
export interface WorkContextHeaderButtonCfg {
  label: string;
  icon?: string;
  showFor: Array<'PROJECT' | 'TAG' | 'TODAY'>;
  onClick: (workContext: WorkContext) => void | Promise<void>;
  [key: string]: unknown;
}

/** 当前活跃的工作上下文（项目 / 标签 / 今天视图）。 */
export interface WorkContext {
  id: string;
  type: 'PROJECT' | 'TAG' | 'TODAY';
  [key: string]: unknown;
}

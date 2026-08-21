import { describe, it, expect } from 'vitest';
import { graphRequestAll } from '../../src/graph';
import { makeFetchHttp } from './helpers/fetch-http';
import type { TodoTask } from '../../src/domain';

// 真实 Graph 集成测试：只有当提供了访问令牌与列表 ID 时才会运行（跳过时不失败）。
// 设置方式：
//   PowerShell: $env:MS_TODO_ACCESS_TOKEN='...'; $env:MS_TODO_LIST_ID='...'
//   bash:       export MS_TODO_ACCESS_TOKEN='...' MS_TODO_LIST_ID='...'
// 令牌可用 Entra 的 client-credentials 或一次性的委托令牌（需 Tasks.ReadWrite）。
const TOKEN = process.env.MS_TODO_ACCESS_TOKEN;
const LIST_ID = process.env.MS_TODO_LIST_ID;

describe.skipIf(!TOKEN || !LIST_ID)('Microsoft Graph integration', () => {
  it('lists tasks from the configured list over a real HTTPS round-trip', async () => {
    // 用 Node 内置 fetch 充当宿主 PluginHttp，走完整 graphRequest → http.request
    // 路径，对真实 Graph 发请求。
    const tasks = await graphRequestAll<TodoTask>(
      `/me/todo/lists/${LIST_ID}/tasks`,
      TOKEN!,
      10,
      makeFetchHttp(),
    );

    expect(Array.isArray(tasks)).toBe(true);
    if (tasks.length > 0) {
      expect(tasks[0]).toHaveProperty('id');
      expect(tasks[0]).toHaveProperty('title');
    }
  }, 90_000);
});

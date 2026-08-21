import { spawnSync } from 'node:child_process';
import { writeFileSync, unlinkSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

export interface BridgeResult {
  success: boolean;
  result?: unknown;
  error?: string;
}

// 模拟 SP 的 executeNodeScript 桥：把内嵌脚本跑在真实 Node 子进程里，并把脚本
// 的返回值经 stdout 的 JSON 回传。这样集成测试可以复用 graph.ts 里那段自包含的
// https 请求脚本（require + 顶层 await），对真实 Graph 发请求。
//
// SP 的宿主桥也走「spawn Node 子进程 → 脚本返回值经 console.log(JSON) 回传」，
// 这里尽量贴近其语义。仅用于集成测试，不在生产代码路径里。
export function runNodeScript(script: string, args: unknown[]): Promise<BridgeResult> {
  return new Promise((resolve) => {
    const driver = [
      'const script = ' + JSON.stringify(script) + ';',
      'const args = JSON.parse(process.argv[1]);',
      '(async () => {',
      "  const wrapped = '(async () => {\\n' + script + '\\n})()';",
      '  try {',
      '    const result = await eval(wrapped);',
      '    console.log(JSON.stringify({ success: true, result: result }));',
      '  } catch (e) {',
      "    console.log(JSON.stringify({ success: false, error: String((e && e.message) || e) }));",
      '  }',
      '})();',
    ].join('\n');

    const file = join(tmpdir(), `sp-ms-todo-bridge-${process.pid}.cjs`);
    writeFileSync(file, driver, 'utf8');
    try {
      const res = spawnSync(process.execPath, [file, JSON.stringify(args)], {
        encoding: 'utf8',
        timeout: 60_000,
      });
      const out = (res.stdout || '').trim();
      const lastLine = out.split(/\r?\n/).filter(Boolean).pop();
      try {
        resolve(JSON.parse(lastLine || '{}'));
      } catch {
        resolve({ success: false, error: out || (res.stderr || '') });
      }
    } finally {
      try {
        unlinkSync(file);
      } catch {
        // ignore
      }
    }
  });
}

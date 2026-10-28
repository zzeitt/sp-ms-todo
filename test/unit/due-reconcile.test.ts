import { describe, expect, it } from 'vitest';
import { selectDueDivergences } from '../../src/force-sync';
import type { SpTask } from '../../src/types';

const task = (over: Partial<SpTask> = {}): SpTask => ({
  id: 't1',
  title: 'task',
  issueType: 'plugin:ms-todo-provider',
  issueId: 'issue-1',
  issueProviderId: 'provider-1',
  ...over,
});

describe('selectDueDivergences（日期对账）', () => {
  it('本地计划日与基线不同 → 入选', () => {
    const t = task({ dueDay: '2026-10-27', issueLastSyncedValues: { dueDate: '2026-08-21' } });
    expect(selectDueDivergences([t]).map((x) => x.id)).toEqual(['t1']);
  });

  it('本地定时与基线不同 → 入选', () => {
    const t = task({
      dueWithTime: new Date(2026, 10, 15, 14, 30).getTime(),
      issueLastSyncedValues: { dueWithTime: null },
    });
    expect(selectDueDivergences([t]).map((x) => x.id)).toEqual(['t1']);
  });

  it('本地清空日期而基线有日期 → 入选（需要把清空推给远端）', () => {
    const t = task({ issueLastSyncedValues: { dueDate: '2026-08-21' } });
    expect(selectDueDivergences([t]).map((x) => x.id)).toEqual(['t1']);
  });

  it('一致（含都为空的 null/undefined 形式）→ 不入选', () => {
    const same = [
      task({ dueDay: '2026-10-27', issueLastSyncedValues: { dueDate: '2026-10-27' } }),
      task({ id: 't2', issueLastSyncedValues: { dueDate: null } }),
      task({ id: 't3', issueLastSyncedValues: undefined }),
      task({
        id: 't4',
        dueWithTime: 123,
        issueLastSyncedValues: { dueWithTime: 123 },
      }),
    ];
    expect(selectDueDivergences(same)).toEqual([]);
  });

  it('忽略其它 provider / 未链接任务', () => {
    const others = [
      task({ issueType: 'plugin:other', dueDay: '2026-10-27', issueLastSyncedValues: {} }),
      task({ issueId: undefined, dueDay: '2026-10-27' }),
    ];
    expect(selectDueDivergences(others)).toEqual([]);
  });
});

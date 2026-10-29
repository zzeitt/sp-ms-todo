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

  it('已完成任务一律不入选（SP 打勾会清空本地日期，不能拿去清远端）', () => {
    const t = task({ isDone: true, issueLastSyncedValues: { dueDate: '2026-08-21' } });
    expect(selectDueDivergences([t])).toEqual([]);
  });

  it('基线带时刻、本地同一日期只剩纯日期 → 旧版本映射缺口，不入选', () => {
    const t = task({
      dueDay: '2026-11-07',
      issueLastSyncedValues: {
        dueDate: '2026-11-07',
        dueWithTime: new Date(2026, 10, 7, 9, 33).getTime(),
      },
    });
    expect(selectDueDivergences([t])).toEqual([]);
  });

  it('基线只有旧版 UTC 午夜时刻、本地同一日期 → 不入选', () => {
    const t = task({
      dueDay: '2026-08-15',
      issueLastSyncedValues: { dueWithTime: new Date(2026, 7, 15, 0, 0).getTime() },
    });
    expect(selectDueDivergences([t])).toEqual([]);
  });

  it('两边时刻相同、基线多出派生 dueDate → 不入选', () => {
    const ts = new Date(2026, 9, 29, 20, 0).getTime();
    const t = task({
      dueWithTime: ts,
      issueLastSyncedValues: { dueDate: '2026-10-29', dueWithTime: ts },
    });
    expect(selectDueDivergences([t])).toEqual([]);
  });

  it('本地把定时推迟一天 → 入选（真实改动）', () => {
    const t = task({
      dueWithTime: new Date(2026, 9, 29, 20, 0).getTime(),
      issueLastSyncedValues: { dueWithTime: new Date(2026, 9, 28, 20, 0).getTime() },
    });
    expect(selectDueDivergences([t]).map((x) => x.id)).toEqual(['t1']);
  });

  it('忽略其它 provider / 未链接任务', () => {
    const others = [
      task({ issueType: 'plugin:other', dueDay: '2026-10-27', issueLastSyncedValues: {} }),
      task({ issueId: undefined, dueDay: '2026-10-27' }),
    ];
    expect(selectDueDivergences(others)).toEqual([]);
  });
});

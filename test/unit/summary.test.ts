import { describe, expect, it } from 'vitest';
import { summarizeSyncState } from '../../src/force-sync';
import type { SpTask } from '../../src/types';

const task = (over: Partial<SpTask> = {}): SpTask => ({
  id: 't',
  title: 'task',
  issueType: 'plugin:ms-todo-provider',
  issueId: 'issue-1',
  issueProviderId: 'provider-1',
  ...over,
});

describe('summarizeSyncState（概览统计）', () => {
  it('统计已链接数、完成分叉、日期分叉与最近同步时间', () => {
    const done = task({
      id: 'done',
      isDone: true,
      issueId: 'i-done',
      issueLastUpdated: 1_000,
      issueLastSyncedValues: { state: 'notStarted' },
    });
    const due = task({
      id: 'due',
      issueId: 'i-due',
      issueLastUpdated: 3_000,
      dueDay: '2026-10-30',
      issueLastSyncedValues: { dueDate: '2026-10-29' },
    });
    const fresh = task({ id: 'fresh', issueId: 'i-fresh', issueLastUpdated: 2_000 });
    const unlinked = task({ id: 'local', issueId: undefined });

    const summary = summarizeSyncState([done, due, fresh, unlinked]);
    expect(summary).toEqual({
      linked: 3,
      completionDiverged: 1,
      dueDiverged: 1,
      lastSyncAt: 3_000,
    });
  });

  it('没有已链接任务时 lastSyncAt 为 null', () => {
    expect(summarizeSyncState([task({ issueId: undefined })])).toEqual({
      linked: 0,
      completionDiverged: 0,
      dueDiverged: 0,
      lastSyncAt: null,
    });
  });

  it('忽略非本插件、无 issueId 与非法时间戳', () => {
    const other = task({ issueType: 'plugin:other', issueLastUpdated: 9_999 });
    const bad = task({ id: 'bad', issueId: 'i-bad', issueLastUpdated: Number.NaN });
    const summary = summarizeSyncState([other, bad]);
    expect(summary.linked).toBe(1);
    expect(summary.lastSyncAt).toBeNull();
  });
});

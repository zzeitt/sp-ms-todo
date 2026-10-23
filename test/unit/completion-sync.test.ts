import { describe, expect, it } from 'vitest';
import {
  decideCompletionSync,
  filterIssuesByIds,
  selectCompletionDivergences,
  shouldPushCompletionOnUpdate,
} from '../../src/force-sync';
import type { SpTask } from '../../src/types';

const task = (over: Partial<SpTask> = {}): SpTask => ({
  id: 't1',
  title: 'task',
  issueType: 'plugin:ms-todo-provider',
  issueId: 'issue-1',
  issueProviderId: 'provider-1',
  isDone: true,
  ...over,
});

describe('decideCompletionSync', () => {
  it('local done + remote completed -> only sync the baseline (no write to MS)', () => {
    expect(
      decideCompletionSync({ isDone: true, baselineState: 'notStarted', remoteState: 'completed' })
        .action,
    ).toBe('baseline-only');
  });

  it('local done + remote not completed + unchanged since baseline -> push', () => {
    expect(
      decideCompletionSync({ isDone: true, baselineState: 'notStarted', remoteState: 'notStarted' })
        .action,
    ).toBe('push');
  });

  it('local done + no baseline -> push (explicit completion intent wins)', () => {
    expect(decideCompletionSync({ isDone: true, remoteState: 'notStarted' }).action).toBe('push');
  });

  it('local done + remote changed away from baseline -> still push the completion', () => {
    const d = decideCompletionSync({
      isDone: true,
      baselineState: 'notStarted',
      remoteState: 'inProgress',
    });
    expect(d.action).toBe('push');
    expect(d.reason).toMatch(/remote changed/i);
  });

  it('local done + remote state unknown (fetch failed) -> do nothing', () => {
    expect(decideCompletionSync({ isDone: true, baselineState: 'notStarted' }).action).toBe('none');
    expect(decideCompletionSync({ isDone: true, remoteState: undefined }).action).toBe('none');
  });

  it('local open + remote completed -> pull (mark the local task done)', () => {
    expect(
      decideCompletionSync({ isDone: false, baselineState: 'notStarted', remoteState: 'completed' })
        .action,
    ).toBe('pull');
  });

  it('local open + remote completed -> pull even when the baseline already says completed', () => {
    expect(
      decideCompletionSync({ isDone: false, baselineState: 'completed', remoteState: 'completed' })
        .action,
    ).toBe('pull');
  });

  it('local open + remote open -> nothing to do', () => {
    expect(
      decideCompletionSync({ isDone: false, baselineState: 'notStarted', remoteState: 'notStarted' })
        .action,
    ).toBe('none');
  });
});

describe('shouldPushCompletionOnUpdate (taskUpdate hook guard)', () => {
  it('fires for our linked task completed locally', () => {
    expect(
      shouldPushCompletionOnUpdate(task({ issueLastSyncedValues: { state: 'notStarted' } }), {
        isDone: true,
      }),
    ).toBe(true);
  });

  it('ignores reopen (isDone=false)', () => {
    expect(shouldPushCompletionOnUpdate(task({ isDone: false }), { isDone: false })).toBe(false);
  });

  it('ignores pull-originated updates that carry sync bookkeeping', () => {
    expect(
      shouldPushCompletionOnUpdate(task(), {
        isDone: true,
        issueLastSyncedValues: { state: 'completed' },
      }),
    ).toBe(false);
    expect(shouldPushCompletionOnUpdate(task(), { isDone: true, issueWasUpdated: true })).toBe(
      false,
    );
  });

  it('ignores tasks of other providers / unlinked tasks / already-synced completion', () => {
    expect(shouldPushCompletionOnUpdate(task({ issueType: 'plugin:other' }), { isDone: true })).toBe(
      false,
    );
    expect(shouldPushCompletionOnUpdate(task({ issueId: undefined }), { isDone: true })).toBe(false);
    expect(
      shouldPushCompletionOnUpdate(task({ issueLastSyncedValues: { state: 'completed' } }), {
        isDone: true,
      }),
    ).toBe(false);
  });
});

describe('selectCompletionDivergences', () => {
  const all: SpTask[] = [
    task({ id: 'push', isDone: true, issueLastSyncedValues: { state: 'notStarted' } }),
    task({ id: 'pull', isDone: false, issueLastSyncedValues: { state: 'completed' } }),
    task({ id: 'same-done', isDone: true, issueLastSyncedValues: { state: 'completed' } }),
    task({ id: 'same-open', isDone: false, issueLastSyncedValues: { state: 'notStarted' } }),
    task({ id: 'other', issueType: 'plugin:other', isDone: true, issueLastSyncedValues: { state: 'notStarted' } }),
    task({ id: 'unlinked', issueId: undefined, isDone: true }),
    task({ id: 'no-baseline', isDone: true, issueLastSyncedValues: undefined }),
  ];

  it('selects only the diverged, linked tasks of this provider', () => {
    expect(selectCompletionDivergences(all).map((t) => t.id).sort()).toEqual([
      'no-baseline',
      'pull',
      'push',
    ]);
  });
});

describe('filterIssuesByIds (getByIds batching)', () => {
  const issues = [{ id: 'a' }, { id: 'b' }, { id: 'c' }];
  it('returns only the requested issues, keeping provider order', () => {
    expect(filterIssuesByIds(issues, ['c', 'a']).map((i) => i.id)).toEqual(['a', 'c']);
  });
  it('returns an empty array when nothing matches', () => {
    expect(filterIssuesByIds(issues, ['zzz'])).toEqual([]);
  });
});

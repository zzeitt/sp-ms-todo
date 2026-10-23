import { describe, it, expect } from 'vitest';
import {
  buildCreateTaskBody,
  buildProjectToListMap,
  classifyLinkedTasks,
  findIssueProviderId,
  isMsTodoTask,
  isPushableTask,
  isUnsynced,
  majorityList,
  PLUGIN_ISSUE_TYPE,
  resolvePushListId,
  selectLinkedMsTasks,
  selectTasksToPush,
  selectTasksToSync,
} from '../../src/force-sync';
import type { SpTask } from '../../src/types';

function makeTask(overrides: Partial<SpTask> = {}): SpTask {
  return {
    id: 'local-1',
    title: 'Buy milk',
    issueType: PLUGIN_ISSUE_TYPE,
    ...overrides,
  };
}

describe('isMsTodoTask', () => {
  it('matches this provider by issueType', () => {
    expect(isMsTodoTask(makeTask())).toBe(true);
  });

  it('rejects other issue types', () => {
    expect(isMsTodoTask(makeTask({ issueType: 'GITHUB' }))).toBe(false);
    expect(isMsTodoTask(makeTask({ issueType: undefined }))).toBe(false);
  });

  it('rejects null / undefined', () => {
    expect(isMsTodoTask(null)).toBe(false);
    expect(isMsTodoTask(undefined)).toBe(false);
  });
});

describe('isUnsynced', () => {
  it('is unsynced when there is no issueId', () => {
    expect(isUnsynced(makeTask({ issueId: undefined }))).toBe(true);
  });

  it('is synced when issueId is present', () => {
    expect(isUnsynced(makeTask({ issueId: 'graph-1' }))).toBe(false);
  });
});

describe('isPushableTask', () => {
  it('allows an unsynced provider task', () => {
    expect(isPushableTask(makeTask({ issueId: undefined }))).toBe(true);
  });

  it('allows a failed auto-create task (no issueType, no issueId)', () => {
    expect(isPushableTask(makeTask({ issueType: undefined, issueId: undefined }))).toBe(true);
  });

  it('rejects a synced task (has issueId)', () => {
    expect(isPushableTask(makeTask({ issueId: 'g1' }))).toBe(false);
  });

  it('rejects another provider task', () => {
    expect(isPushableTask(makeTask({ issueType: 'GITHUB', issueId: undefined }))).toBe(false);
  });

  it('rejects null / undefined', () => {
    expect(isPushableTask(null)).toBe(false);
    expect(isPushableTask(undefined)).toBe(false);
  });
});

describe('selectTasksToPush', () => {
  const unsynced = makeTask({ id: 'a', issueId: undefined });
  const failedAutoCreate = makeTask({ id: 'b', issueType: undefined, issueId: undefined });
  const synced = makeTask({ id: 'c', issueId: 'graph-c' });
  const other = makeTask({ id: 'd', issueType: 'GITHUB', issueId: undefined });

  it('picks the selected unsynced provider task', () => {
    expect(selectTasksToPush(unsynced).map((t) => t.id)).toEqual(['a']);
  });

  it('picks a failed auto-create task (no issueType)', () => {
    expect(selectTasksToPush(failedAutoCreate).map((t) => t.id)).toEqual(['b']);
  });

  it('does not push a synced selected task', () => {
    expect(selectTasksToPush(synced)).toEqual([]);
  });

  it('does not push a non-provider selected task', () => {
    expect(selectTasksToPush(other)).toEqual([]);
  });

  it('returns empty when nothing is selected', () => {
    expect(selectTasksToPush(null)).toEqual([]);
    expect(selectTasksToPush(undefined)).toEqual([]);
  });
});

describe('selectTasksToSync', () => {
  const unsyncedA = makeTask({ id: 'a', issueId: undefined, projectId: 'p1' });
  const unsyncedB = makeTask({ id: 'b', issueType: undefined, issueId: undefined, projectId: 'p2' });
  const synced = makeTask({ id: 'c', issueId: 'graph-c', projectId: 'p1' });
  const other = makeTask({ id: 'd', issueType: 'GITHUB', issueId: undefined, projectId: 'p1' });

  it('returns all pushable tasks when not scoped to a project', () => {
    expect(selectTasksToSync([unsyncedA, unsyncedB, synced, other]).map((t) => t.id)).toEqual([
      'a',
      'b',
    ]);
  });

  it('scopes to the given project', () => {
    expect(
      selectTasksToSync([unsyncedA, unsyncedB, synced, other], 'p1').map((t) => t.id),
    ).toEqual(['a']);
  });

  it('returns empty when nothing is pushable', () => {
    expect(selectTasksToSync([synced, other])).toEqual([]);
  });

  it('returns empty for an empty list', () => {
    expect(selectTasksToSync([])).toEqual([]);
  });
});

describe('findIssueProviderId', () => {
  const linked = makeTask({
    id: 'a',
    issueId: 'g1',
    issueProviderId: 'prov-1',
    projectId: 'p1',
  });
  const otherProject = makeTask({
    id: 'b',
    issueId: 'g2',
    issueProviderId: 'prov-2',
    projectId: 'p2',
  });

  it('finds the provider id from a linked task in the same project', () => {
    expect(findIssueProviderId([linked], 'p1')).toBe('prov-1');
  });

  it('skips linked tasks in other projects when scoped', () => {
    expect(findIssueProviderId([otherProject], 'p1')).toBeUndefined();
  });

  it('returns the first provider id when not scoped to a project', () => {
    expect(findIssueProviderId([otherProject, linked])).toBe('prov-2');
  });

  it('ignores unlinked / non-provider tasks', () => {
    const unlinked = makeTask({ id: 'c', issueId: undefined, issueProviderId: 'prov-x' });
    const other = makeTask({ id: 'd', issueType: 'GITHUB', issueId: 'g3', issueProviderId: 'prov-y' });
    expect(findIssueProviderId([unlinked, other])).toBeUndefined();
  });
});

describe('selectLinkedMsTasks', () => {
  const linked = makeTask({ id: 'a', issueId: 'g1' });
  const unsynced = makeTask({ id: 'b', issueId: undefined });
  const other = makeTask({ id: 'c', issueType: 'GITHUB', issueId: 'g2' });

  it('returns linked provider tasks only', () => {
    const picked = selectLinkedMsTasks([linked, unsynced, other]);
    expect(picked.map((t) => t.id)).toEqual(['a']);
  });

  it('returns empty when nothing is linked', () => {
    expect(selectLinkedMsTasks([])).toEqual([]);
    expect(selectLinkedMsTasks([unsynced])).toEqual([]);
  });
});

describe('buildCreateTaskBody', () => {
  it('builds a title-only body', () => {
    expect(buildCreateTaskBody(makeTask({ title: 'A' }))).toEqual({ title: 'A' });
  });

  it('maps isDone to completed status', () => {
    expect(buildCreateTaskBody(makeTask({ isDone: true })).status).toBe('completed');
  });

  it('maps notes to body.content', () => {
    expect(buildCreateTaskBody(makeTask({ notes: 'do it' })).body).toEqual({
      content: 'do it',
      contentType: 'text',
    });
  });

  it('maps dueDay to UTC-midnight dueDateTime', () => {
    expect(buildCreateTaskBody(makeTask({ dueDay: '2026-08-26' })).dueDateTime).toEqual({
      dateTime: '2026-08-26T00:00:00.0000000',
      timeZone: 'UTC',
    });
  });

  it('prefers dueWithTime (full UTC wall-clock) over dueDay', () => {
    const ms = Date.UTC(2026, 7, 26, 6, 30, 0);
    expect(
      buildCreateTaskBody(makeTask({ dueDay: '2026-08-26', dueWithTime: ms })).dueDateTime,
    ).toEqual({ dateTime: '2026-08-26T06:30:00.0000000', timeZone: 'UTC' });
  });
});

describe('buildProjectToListMap', () => {
  const msTaskToList = new Map<string, string>([
    ['g1', 'list-a'],
    ['g2', 'list-b'],
  ]);

  it('maps each project to the list of its linked task', () => {
    const linked = [
      makeTask({ id: 'a', issueId: 'g1', projectId: 'p1' }),
      makeTask({ id: 'b', issueId: 'g2', projectId: 'p2' }),
    ];
    expect(buildProjectToListMap(linked, msTaskToList)).toEqual(
      new Map([['p1', 'list-a'], ['p2', 'list-b']]),
    );
  });

  it('takes the first list when a project has several linked tasks', () => {
    const linked = [
      makeTask({ id: 'a', issueId: 'g1', projectId: 'p1' }),
      makeTask({ id: 'b', issueId: 'g2', projectId: 'p1' }),
    ];
    expect(buildProjectToListMap(linked, msTaskToList).get('p1')).toBe('list-a');
  });

  it('skips linked tasks whose issueId is not in the index', () => {
    const linked = [makeTask({ id: 'a', issueId: 'g-unknown', projectId: 'p1' })];
    expect(buildProjectToListMap(linked, msTaskToList).has('p1')).toBe(false);
  });

  it('skips tasks without projectId', () => {
    const linked = [makeTask({ id: 'a', issueId: 'g1', projectId: undefined })];
    expect(buildProjectToListMap(linked, msTaskToList).size).toBe(0);
  });
});

describe('resolvePushListId', () => {
  const projectToList = new Map<string, string>([['p1', 'list-a']]);
  const task = makeTask({ projectId: 'p1' });

  it('prefers the explicit target list', () => {
    expect(resolvePushListId(task, 'list-explicit', projectToList, 'list-fallback')).toBe(
      'list-explicit',
    );
  });

  it('falls back to the task project bound list', () => {
    expect(resolvePushListId(task, undefined, projectToList, 'list-fallback')).toBe('list-a');
  });

  it('falls back to the fallback list when the project is unknown', () => {
    expect(resolvePushListId(task, undefined, new Map(), 'list-fallback')).toBe('list-fallback');
  });

  it('returns undefined when nothing resolves', () => {
    expect(resolvePushListId(task, undefined, new Map(), undefined)).toBeUndefined();
  });
});

describe('majorityList', () => {
  it('returns undefined for empty / all-empty input', () => {
    expect(majorityList([])).toBeUndefined();
    expect(majorityList([undefined, undefined])).toBeUndefined();
  });

  it('returns the only value', () => {
    expect(majorityList(['list-a'])).toBe('list-a');
  });

  it('returns the most frequent value', () => {
    expect(majorityList(['a', 'b', 'a'])).toBe('a');
  });

  it('breaks ties by first-inserted order', () => {
    expect(majorityList(['a', 'b'])).toBe('a');
  });
});

describe('classifyLinkedTasks', () => {
  const index = new Map<string, string>([
    ['g1', 'list-a'],
    ['g2', 'list-a'],
    ['g3', 'list-b'],
  ]);

  it('flags a task whose issueId is missing from the index as dead', () => {
    const linked = [makeTask({ id: 't', issueId: 'g-gone', projectId: 'p1' })];
    const health = classifyLinkedTasks(linked, index);
    expect(health.dead.map((d) => d.issueId)).toEqual(['g-gone']);
    expect(health.mismatched).toEqual([]);
  });

  it('flags a task in a list different from its project majority as mismatched', () => {
    const linked = [
      makeTask({ id: 'a', issueId: 'g1', projectId: 'p1' }),
      makeTask({ id: 'b', issueId: 'g2', projectId: 'p1' }),
      makeTask({ id: 'c', issueId: 'g3', projectId: 'p1' }),
    ];
    const health = classifyLinkedTasks(linked, index);
    expect(health.dead).toEqual([]);
    expect(health.mismatched.map((m) => m.issueId)).toEqual(['g3']);
    expect(health.mismatched[0].listId).toBe('list-b');
    expect(health.mismatched[0].expectedListId).toBe('list-a');
  });

  it('does not flag tasks that all live in the same list', () => {
    const linked = [
      makeTask({ id: 'a', issueId: 'g1', projectId: 'p1' }),
      makeTask({ id: 'b', issueId: 'g2', projectId: 'p1' }),
    ];
    const health = classifyLinkedTasks(linked, index);
    expect(health.dead).toEqual([]);
    expect(health.mismatched).toEqual([]);
    expect(health.checked).toBe(2);
  });

  it('does not flag a project with a single linked task', () => {
    const linked = [makeTask({ id: 'a', issueId: 'g3', projectId: 'p1' })];
    const health = classifyLinkedTasks(linked, index);
    expect(health.dead).toEqual([]);
    expect(health.mismatched).toEqual([]);
  });

  it('groups tasks without projectId together', () => {
    const linked = [
      makeTask({ id: 'a', issueId: 'g1', projectId: undefined }),
      makeTask({ id: 'b', issueId: 'g3', projectId: undefined }),
    ];
    const health = classifyLinkedTasks(linked, index);
    expect(health.mismatched.map((m) => m.issueId)).toEqual(['g3']);
  });
});


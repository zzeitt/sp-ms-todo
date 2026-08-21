import { describe, it, expect } from 'vitest';
import {
  mapTaskToIssue,
  mapTaskToSearchResult,
  extractDueDate,
  buildDueDateTime,
  extractSyncValues,
  TODO_WEB_URL,
} from '../../src/domain';
import type { TodoTask } from '../../src/domain';

const baseTask: TodoTask = {
  id: 'task-1',
  title: 'Buy milk',
  status: 'notStarted',
};

describe('extractDueDate', () => {
  it('extracts the date part', () => {
    expect(
      extractDueDate({ dateTime: '2026-08-21T10:30:00.0000000', timeZone: 'UTC' }),
    ).toBe('2026-08-21');
  });

  it('returns undefined for missing / null / empty', () => {
    expect(extractDueDate(undefined)).toBeUndefined();
    expect(extractDueDate(null)).toBeUndefined();
    expect(extractDueDate({ dateTime: '', timeZone: 'UTC' })).toBeUndefined();
  });
});

describe('buildDueDateTime', () => {
  it('builds UTC midnight', () => {
    expect(buildDueDateTime('2026-08-21')).toEqual({
      dateTime: '2026-08-21T00:00:00.0000000',
      timeZone: 'UTC',
    });
  });

  it('returns null for empty / null / undefined', () => {
    expect(buildDueDateTime('')).toBeNull();
    expect(buildDueDateTime(null)).toBeNull();
    expect(buildDueDateTime(undefined)).toBeNull();
  });
});

describe('mapTaskToIssue', () => {
  it('maps core fields with sensible defaults', () => {
    const issue = mapTaskToIssue(baseTask);
    expect(issue.id).toBe('task-1');
    expect(issue.title).toBe('Buy milk');
    expect(issue.state).toBe('notStarted');
    expect(issue.url).toBe(TODO_WEB_URL);
    expect(issue.body).toBe('');
  });

  it('maps body content, dueDate and importance', () => {
    const issue = mapTaskToIssue({
      ...baseTask,
      body: { content: 'notes here', contentType: 'text' },
      dueDateTime: { dateTime: '2026-08-21T00:00:00.0000000', timeZone: 'UTC' },
      importance: 'high',
    });
    expect(issue.body).toBe('notes here');
    expect(issue.dueDate).toBe('2026-08-21');
    expect(issue.importance).toBe('high');
  });

  it('maps lastUpdated as epoch millis', () => {
    const issue = mapTaskToIssue({
      ...baseTask,
      lastModifiedDateTime: '2026-08-21T00:00:00Z',
    });
    expect(issue.lastUpdated).toBe(new Date('2026-08-21T00:00:00Z').getTime());
  });

  it('leaves lastUpdated undefined when absent', () => {
    expect(mapTaskToIssue(baseTask).lastUpdated).toBeUndefined();
  });
});

describe('mapTaskToSearchResult', () => {
  it('spreads issue fields plus description and labels', () => {
    const sr = mapTaskToSearchResult({
      ...baseTask,
      status: 'completed',
      body: { content: 'the body', contentType: 'text' },
      dueDateTime: { dateTime: '2026-08-21T00:00:00.0000000', timeZone: 'UTC' },
      lastModifiedDateTime: '2026-08-21T00:00:00Z',
      importance: 'high',
    });

    // 回归点：SP 导入引擎读的是 PluginIssue 的同名字段
    // （state / body / dueDate / lastUpdated），而非 status / description。
    expect(sr.state).toBe('completed');
    expect(sr.body).toBe('the body');
    expect(sr.dueDate).toBe('2026-08-21');
    expect(sr.lastUpdated).toBe(new Date('2026-08-21T00:00:00Z').getTime());
    expect(sr.description).toBe('the body');
    expect(sr.labels).toEqual(['high']);
  });

  it('omits labels for normal importance', () => {
    expect(mapTaskToSearchResult({ ...baseTask, importance: 'normal' }).labels).toEqual([]);
    expect(mapTaskToSearchResult(baseTask).labels).toEqual([]);
  });
});

describe('extractSyncValues', () => {
  it('extracts sync fields with defaults', () => {
    expect(extractSyncValues({ id: 'x', title: 'T', state: 'completed' })).toEqual({
      state: 'completed',
      title: 'T',
      body: '',
      dueDate: null,
    });
  });
});

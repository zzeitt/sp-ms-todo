import { describe, it, expect } from 'vitest';
import {
  mapTaskToIssue,
  mapTaskToSearchResult,
  extractDueDate,
  buildDueDateTime,
  buildDueDateTimeFromTs,
  extractDueWithTime,
  formatDueDisplay,
  extractSyncValues,
  TODO_WEB_URL,
} from '../../src/domain';
import type { TodoTask } from '../../src/domain';

const baseTask: TodoTask = {
  id: 'task-1',
  title: 'Buy milk',
  status: 'notStarted',
};

const pad2 = (n: number): string => String(n).padStart(2, '0');

/** 本机时区某天午夜（测试在任何时区都成立）。 */
const localMidnight = (y: number, m: number, d: number): number =>
  new Date(y, m - 1, d, 0, 0, 0, 0).getTime();

/** MS 表示“纯日期”的形态：当地午夜的 UTC 即时（中国即前一天 16:00Z）。 */
const dateOnlyDue = (y: number, m: number, d: number) => {
  const t = new Date(localMidnight(y, m, d));
  return {
    dateTime:
      `${t.getUTCFullYear()}-${pad2(t.getUTCMonth() + 1)}-${pad2(t.getUTCDate())}` +
      `T${pad2(t.getUTCHours())}:${pad2(t.getUTCMinutes())}:${pad2(t.getUTCSeconds())}.0000000`,
    timeZone: 'UTC',
  };
};

describe('extractDueDate', () => {
  it('extracts the local date of a date-only reminder', () => {
    expect(extractDueDate(dateOnlyDue(2026, 8, 21))).toBe('2026-08-21');
  });

  it('returns undefined for missing / null / empty', () => {
    expect(extractDueDate(undefined)).toBeUndefined();
    expect(extractDueDate(null)).toBeUndefined();
    expect(extractDueDate({ dateTime: '', timeZone: 'UTC' })).toBeUndefined();
  });
});

describe('buildDueDateTime', () => {
  it('writes date-only as <date>T00:00:00 in the device timezone', () => {
    const due = buildDueDateTime('2026-08-21')!;
    expect(due.dateTime).toBe('2026-08-21T00:00:00.0000000');
    expect(due.timeZone).toBe(Intl.DateTimeFormat().resolvedOptions().timeZone);
  });

  it('returns null for empty / null / undefined', () => {
    expect(buildDueDateTime('')).toBeNull();
    expect(buildDueDateTime(null)).toBeNull();
    expect(buildDueDateTime(undefined)).toBeNull();
  });
});

describe('buildDueDateTimeFromTs', () => {
  it('builds a UTC wall-clock from a timestamp', () => {
    const ms = Date.UTC(2026, 7, 21, 6, 30, 45);
    expect(buildDueDateTimeFromTs(ms)).toEqual({
      dateTime: '2026-08-21T06:30:45.0000000',
      timeZone: 'UTC',
    });
  });

  it('returns null for non-number / null / undefined', () => {
    expect(buildDueDateTimeFromTs(null)).toBeNull();
    expect(buildDueDateTimeFromTs(undefined)).toBeNull();
    expect(buildDueDateTimeFromTs(Number.NaN)).toBeNull();
  });
});

describe('extractDueWithTime', () => {
  it('returns undefined for local midnight (date-only reminder)', () => {
    expect(extractDueWithTime(dateOnlyDue(2026, 8, 21))).toBeUndefined();
    expect(
      extractDueWithTime({
        dateTime: '2026-08-21T00:00:00.0000000',
        timeZone: 'China Standard Time',
      }),
    ).toBeUndefined();
  });

  it('parses a UTC-tagged time as UTC', () => {
    expect(
      extractDueWithTime({ dateTime: '2026-08-21T06:30:00.0000000', timeZone: 'UTC' }),
    ).toBe(Date.UTC(2026, 7, 21, 6, 30, 0));
  });

  it('parses a non-UTC time as device-local wall-clock', () => {
    const due = { dateTime: '2026-08-21T14:00:00.0000000', timeZone: 'China Standard Time' };
    expect(extractDueWithTime(due)).toBe(new Date(2026, 7, 21, 14, 0, 0).getTime());
  });

  it('returns undefined for missing / null / empty / malformed', () => {
    expect(extractDueWithTime(undefined)).toBeUndefined();
    expect(extractDueWithTime(null)).toBeUndefined();
    expect(extractDueWithTime({ dateTime: '', timeZone: 'UTC' })).toBeUndefined();
    expect(extractDueWithTime({ dateTime: 'not-a-date', timeZone: 'UTC' })).toBeUndefined();
  });
});

describe('formatDueDisplay', () => {
  it('shows date only for a date-only reminder', () => {
    expect(formatDueDisplay(dateOnlyDue(2026, 8, 21))).toBe('2026-08-21');
  });

  it('shows date + local time for a timed reminder', () => {
    expect(
      formatDueDisplay({
        dateTime: '2026-08-21T14:05:00.0000000',
        timeZone: 'China Standard Time',
      }),
    ).toBe('2026-08-21 14:05');
  });

  it('returns undefined when absent', () => {
    expect(formatDueDisplay(undefined)).toBeUndefined();
    expect(formatDueDisplay(null)).toBeUndefined();
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
      dueDateTime: dateOnlyDue(2026, 8, 21),
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

  it('maps reminderDateTime (isReminderOn) to dueWithTime, dueDateTime to dueDate', () => {
    const issue = mapTaskToIssue({
      ...baseTask,
      dueDateTime: dateOnlyDue(2026, 8, 21),
      reminderDateTime: {
        dateTime: '2026-08-21T14:00:00.0000000',
        timeZone: 'China Standard Time',
      },
      isReminderOn: true,
    });
    expect(issue.dueDate).toBe('2026-08-21');
    expect(issue.dueWithTime).toBe(new Date(2026, 7, 21, 14, 0, 0).getTime());
    expect(issue.dueDisplay).toBe('2026-08-21 14:00');
  });

  it('keeps legacy timed dueDateTime values as dueWithTime (backward compatible)', () => {
    const issue = mapTaskToIssue({
      ...baseTask,
      dueDateTime: { dateTime: '2026-08-21T14:00:00.0000000', timeZone: 'China Standard Time' },
    });
    expect(issue.dueDate).toBe('2026-08-21');
    expect(issue.dueWithTime).toBe(new Date(2026, 7, 21, 14, 0, 0).getTime());
    expect(issue.dueDisplay).toBe('2026-08-21 14:00');
  });

  it('emits dueDate (not dueWithTime) for a date-only reminder', () => {
    const issue = mapTaskToIssue({
      ...baseTask,
      dueDateTime: dateOnlyDue(2026, 8, 21),
    });
    expect(issue.dueDate).toBe('2026-08-21');
    expect(issue.dueWithTime).toBeUndefined();
    expect(issue.dueDisplay).toBe('2026-08-21');
  });
});

describe('mapTaskToSearchResult', () => {
  it('spreads issue fields plus description and labels', () => {
    const sr = mapTaskToSearchResult({
      ...baseTask,
      status: 'completed',
      body: { content: 'the body', contentType: 'text' },
      dueDateTime: dateOnlyDue(2026, 8, 21),
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
      dueWithTime: null,
    });
  });

  it('passes through a numeric dueWithTime', () => {
    const ms = Date.UTC(2026, 7, 21, 14, 0, 0);
    expect(extractSyncValues({ id: 'x', title: 'T', dueWithTime: ms })).toEqual({
      state: undefined,
      title: 'T',
      body: '',
      dueDate: null,
      dueWithTime: ms,
    });
  });
});

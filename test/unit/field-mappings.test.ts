import { describe, it, expect } from 'vitest';
import { FIELD_MAPPINGS } from '../../src/domain';

function mappingFor(taskField: string) {
  const m = FIELD_MAPPINGS.find((x) => x.taskField === taskField);
  if (!m) throw new Error(`no field mapping for ${taskField}`);
  return m;
}

const ctx = { issueId: 'i', issueNumber: 1 };

describe('FIELD_MAPPINGS', () => {
  it('declares all five bidirectional mappings', () => {
    expect(FIELD_MAPPINGS.map((m) => m.taskField).sort()).toEqual([
      'dueDay',
      'dueWithTime',
      'isDone',
      'notes',
      'title',
    ]);
    for (const m of FIELD_MAPPINGS) {
      expect(m.defaultDirection).toBe('both');
    }
  });

  it('isDone <-> state', () => {
    const m = mappingFor('isDone');
    expect(m.toIssueValue(true, ctx)).toBe('completed');
    expect(m.toIssueValue(false, ctx)).toBe('notStarted');
    expect(m.toTaskValue('completed', ctx)).toBe(true);
    expect(m.toTaskValue('notStarted', ctx)).toBe(false);
  });

  it('title round-trips', () => {
    const m = mappingFor('title');
    expect(m.toIssueValue('hello', ctx)).toBe('hello');
    expect(m.toTaskValue('hello', ctx)).toBe('hello');
  });

  it('notes <-> body round-trips', () => {
    const m = mappingFor('notes');
    expect(m.toIssueValue('note', ctx)).toBe('note');
    expect(m.toTaskValue('note', ctx)).toBe('note');
  });

  it('dueDay <-> dueDate', () => {
    const m = mappingFor('dueDay');
    expect(m.toIssueValue('2026-08-21', ctx)).toBe('2026-08-21');
    expect(m.toIssueValue('', ctx)).toBeNull();
    expect(m.toTaskValue('2026-08-21', ctx)).toBe('2026-08-21');
    expect(m.toTaskValue('', ctx)).toBeUndefined();
  });

  it('dueWithTime <-> dueWithTime (millis)', () => {
    const m = mappingFor('dueWithTime');
    const ms = Date.UTC(2026, 7, 21, 14, 0, 0);
    expect(m.toIssueValue(ms, ctx)).toBe(ms);
    expect(m.toIssueValue(undefined, ctx)).toBeNull();
    expect(m.toIssueValue('not-a-number', ctx)).toBeNull();
    expect(m.toTaskValue(ms, ctx)).toBe(ms);
    expect(m.toTaskValue(null, ctx)).toBeUndefined();
    expect(m.toTaskValue('not-a-number', ctx)).toBeUndefined();
  });
});

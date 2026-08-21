import { describe, it, expect } from 'vitest';
import { parseConfig, DEFAULT_MAX_TASKS } from '../../src/domain';

describe('parseConfig', () => {
  it('parses a valid listId', () => {
    expect(parseConfig({ listId: 'abc-123' }).listId).toBe('abc-123');
  });

  it('trims whitespace from listId', () => {
    expect(parseConfig({ listId: '  abc  ' }).listId).toBe('abc');
  });

  it('throws when listId is missing', () => {
    expect(() => parseConfig({})).toThrow();
  });

  it('throws when listId is blank', () => {
    expect(() => parseConfig({ listId: '   ' })).toThrow();
  });

  it('defaults maxTasks when absent', () => {
    expect(parseConfig({ listId: 'x' }).maxTasks).toBe(DEFAULT_MAX_TASKS);
  });

  it('parses numeric maxTasks from a string', () => {
    expect(parseConfig({ listId: 'x', maxTasks: '100' }).maxTasks).toBe(100);
  });

  it('treats "0" as unlimited', () => {
    expect(parseConfig({ listId: 'x', maxTasks: '0' }).maxTasks).toBe(0);
  });

  it('falls back to default on invalid maxTasks', () => {
    expect(parseConfig({ listId: 'x', maxTasks: 'abc' }).maxTasks).toBe(DEFAULT_MAX_TASKS);
    expect(parseConfig({ listId: 'x', maxTasks: '-5' }).maxTasks).toBe(DEFAULT_MAX_TASKS);
    expect(parseConfig({ listId: 'x', maxTasks: '3.7' }).maxTasks).toBe(3);
  });
});

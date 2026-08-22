import { describe, it, expect } from 'vitest';
import { parseConfig, getClientId, DEFAULT_MAX_TASKS } from '../../src/domain';

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

describe('getClientId', () => {
  it('prefers the config clientId', () => {
    expect(getClientId({ clientId: 'config-id' }, 'build-id')).toBe('config-id');
  });

  it('trims whitespace from the config clientId', () => {
    expect(getClientId({ clientId: '  config-id  ' }, 'build-id')).toBe('config-id');
  });

  it('falls back to the build-time default when config is empty', () => {
    expect(getClientId({}, 'build-id')).toBe('build-id');
    expect(getClientId({ clientId: '   ' }, 'build-id')).toBe('build-id');
  });

  it('throws when neither config nor default provides a clientId', () => {
    expect(() => getClientId({}, '')).toThrow();
    expect(() => getClientId({ clientId: '' })).toThrow();
  });
});

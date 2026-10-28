import { describe, expect, it } from 'vitest';
import {
  buildDueDateTime,
  buildDueDateTimeFromTs,
  extractDueDate,
  extractDueWithTime,
  parseGraphDue,
} from '../../src/domain';

const pad = (n: number): string => String(n).padStart(2, '0');

/** 把一个绝对时刻表示成 Graph 的 UTC 墙钟形态（模拟 MS 返回纯日期时的表示）。 */
const asUtcWall = (ms: number): { dateTime: string; timeZone: string } => {
  const d = new Date(ms);
  return {
    dateTime:
      `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}` +
      `T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}.0000000`,
    timeZone: 'UTC',
  };
};

/** 本机时区的某天午夜（测试与 CI 在任何时区都应成立）。 */
const localMidnight = (y: number, m: number, d: number): number =>
  new Date(y, m - 1, d, 0, 0, 0, 0).getTime();

describe('仅日期（dueDay ↔ Graph dueDateTime）', () => {
  it('写回纯日期时用 <日期>T00:00:00 + 设备时区名（MS 会归一化为当地午夜）', () => {
    const due = buildDueDateTime('2026-11-15');
    expect(due).not.toBeNull();
    expect(due!.dateTime).toBe('2026-11-15T00:00:00.0000000');
    expect(due!.timeZone).toBe(Intl.DateTimeFormat().resolvedOptions().timeZone);
  });

  it('写回再读回仍是纯日期，且日期不变', () => {
    const due = buildDueDateTime('2026-11-15')!;
    expect(extractDueDate(due)).toBe('2026-11-15');
    expect(extractDueWithTime(due)).toBeUndefined();
    expect(parseGraphDue(due)?.isDateOnly).toBe(true);
    expect(parseGraphDue(due)?.localDay).toBe('2026-11-15');
  });

  it('MS 返回“UTC 即时形式的当地午夜”时识别为纯日期（+8 时区即 16:00Z）', () => {
    const due = asUtcWall(localMidnight(2026, 11, 15));
    expect(extractDueDate(due)).toBe('2026-11-15');
    expect(extractDueWithTime(due)).toBeUndefined();
  });

  it('带时刻的提醒仍按绝对时刻解析，且不会被当成纯日期', () => {
    const ms = new Date(2026, 10, 15, 14, 30, 0, 0).getTime();
    const due = buildDueDateTimeFromTs(ms)!;
    expect(extractDueDate(due)).toBeUndefined();
    expect(extractDueWithTime(due)).toBe(ms);
    expect(parseGraphDue(due)?.isDateOnly).toBe(false);
  });

  it('兼容旧版写回的 UTC 午夜（T00:00:00Z），仍按纯日期读取', () => {
    const legacy = { dateTime: '2026-11-15T00:00:00.0000000', timeZone: 'UTC' };
    expect(extractDueDate(legacy)).toBe('2026-11-15');
    expect(extractDueWithTime(legacy)).toBeUndefined();
    expect(parseGraphDue(legacy)?.isDateOnly).toBe(true);
  });

  it('非 UTC 时区的墙钟 00:00 视为当地午夜（仅日期）', () => {
    const due = { dateTime: '2026-11-15T00:00:00.0000000', timeZone: 'China Standard Time' };
    expect(extractDueDate(due)).toBe('2026-11-15');
    expect(extractDueWithTime(due)).toBeUndefined();
  });

  it('空值 / 非法值返回 undefined，不抛错', () => {
    expect(extractDueDate(null)).toBeUndefined();
    expect(extractDueWithTime(undefined)).toBeUndefined();
    expect(extractDueDate({ dateTime: 'not-a-date', timeZone: 'UTC' })).toBeUndefined();
    expect(buildDueDateTime('')).toBeNull();
    expect(buildDueDateTime('2026-13-40')).toBeNull();
  });
});

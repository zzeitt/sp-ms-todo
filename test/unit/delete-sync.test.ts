import { describe, expect, it } from 'vitest';
import { DELETED_STATE, isDeleteSyncEnabled } from '../../src/force-sync';

describe('isDeleteSyncEnabled（To Do 删除 → SP 删除）', () => {
  it('默认开启：配置缺失时同步删除', () => {
    expect(isDeleteSyncEnabled(undefined)).toBe(true);
    expect(isDeleteSyncEnabled({})).toBe(true);
  });

  it('反向开关未勾选时仍然开启', () => {
    expect(isDeleteSyncEnabled({ keepRemoteDeletedTasks: false })).toBe(true);
  });

  it('勾选“保留远端已删除任务”后关闭', () => {
    expect(isDeleteSyncEnabled({ keepRemoteDeletedTasks: true })).toBe(false);
  });

  it('兼容字符串形式的开关值', () => {
    expect(isDeleteSyncEnabled({ keepRemoteDeletedTasks: 'true' })).toBe(false);
    expect(isDeleteSyncEnabled({ keepRemoteDeletedTasks: 'on' })).toBe(false);
  });

  it('占位 issue 使用宿主识别的 deleted 状态', () => {
    expect(DELETED_STATE).toBe('deleted');
  });
});

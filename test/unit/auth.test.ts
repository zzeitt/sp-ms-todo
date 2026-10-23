import { describe, it, expect, vi } from 'vitest';
import type { PluginHttp } from '../../src/types';
import {
  requestDeviceCode,
  pollForToken,
  refreshAccessToken,
  storeTokens,
  getAccessToken,
  runDeviceCodeLogin,
  withHostDbRetry,
  SECRET_ACCESS_TOKEN,
  SECRET_ACCESS_EXPIRY,
  SECRET_REFRESH_TOKEN,
  type SecretStore,
} from '../../src/auth';

const noSleep = async (): Promise<void> => {};

function mockHttp(): PluginHttp {
  return {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
    request: vi.fn(),
  } as unknown as PluginHttp;
}

function mockSecrets(initial: Record<string, string> = {}): SecretStore & {
  data: Map<string, string>;
} {
  const data = new Map(Object.entries(initial));
  return {
    data,
    get: vi.fn(async (k: string) => data.get(k) ?? null),
    set: vi.fn(async (k: string, v: string) => {
      data.set(k, v);
    }),
    delete: vi.fn(async (k: string) => {
      data.delete(k);
    }),
  };
}

describe('requestDeviceCode', () => {
  it('posts client_id + scopes and returns the device code', async () => {
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>).mockResolvedValue({
      device_code: 'dc',
      user_code: 'ABC-123',
      verification_uri: 'https://microsoft.com/devicelogin',
      expires_in: 900,
      interval: 5,
    });

    const res = await requestDeviceCode('cid', http);

    expect(res.device_code).toBe('dc');
    expect(res.user_code).toBe('ABC-123');
    expect(http.post).toHaveBeenCalledTimes(1);
    const [url, body] = (http.post as ReturnType<typeof vi.fn>).mock.calls[0];
    expect(url).toContain('/devicecode');
    expect(body).toContain('client_id=cid');
    expect(body).toContain('scope=Tasks.ReadWrite%20offline_access');
  });

  it('throws when the response is missing required fields', async () => {
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>).mockResolvedValue({ error: 'nope' });

    await expect(requestDeviceCode('cid', http)).rejects.toThrow();
  });
});

describe('pollForToken', () => {
  it('keeps polling on authorization_pending then succeeds', async () => {
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>)
      .mockResolvedValueOnce({ error: 'authorization_pending' })
      .mockResolvedValueOnce({ access_token: 'at', refresh_token: 'rt', expires_in: 3600 });

    const tokens = await pollForToken('cid', 'dc', 5, http, 900_000, noSleep);

    expect(tokens.access_token).toBe('at');
    expect(tokens.refresh_token).toBe('rt');
    expect(http.post).toHaveBeenCalledTimes(2);
  });

  it('keeps polling on slow_down then succeeds', async () => {
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>)
      .mockResolvedValueOnce({ error: 'slow_down' })
      .mockResolvedValueOnce({ access_token: 'at', expires_in: 3600 });

    const tokens = await pollForToken('cid', 'dc', 5, http, 900_000, noSleep);

    expect(tokens.access_token).toBe('at');
  });

  it('throws on a terminal error', async () => {
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>).mockResolvedValue({
      error: 'access_denied',
      error_description: 'denied',
    });

    await expect(pollForToken('cid', 'dc', 5, http, 900_000, noSleep)).rejects.toThrow();
  });

  it('throws when the poll budget is exhausted', async () => {
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>).mockResolvedValue({
      error: 'authorization_pending',
    });

    await expect(pollForToken('cid', 'dc', 5, http, 0, noSleep)).rejects.toThrow();
  });
});

describe('refreshAccessToken', () => {
  it('exchanges a refresh token for new tokens', async () => {
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>).mockResolvedValue({
      access_token: 'new-at',
      refresh_token: 'new-rt',
      expires_in: 3600,
    });

    const tokens = await refreshAccessToken('cid', 'rt', http);

    expect(tokens.access_token).toBe('new-at');
    const [url, body] = (http.post as ReturnType<typeof vi.fn>).mock.calls[0];
    expect(url).toContain('/token');
    expect(body).toContain('grant_type=refresh_token');
    expect(body).toContain('refresh_token=rt');
  });

  it('throws on an error response', async () => {
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>).mockResolvedValue({
      error: 'invalid_grant',
    });

    await expect(refreshAccessToken('cid', 'rt', http)).rejects.toThrow();
  });
});

describe('storeTokens', () => {
  it('stores access token, expiry and refresh token', async () => {
    const secrets = mockSecrets();
    await storeTokens({ access_token: 'at', refresh_token: 'rt', expires_in: 3600 }, secrets);

    expect(secrets.data.get(SECRET_ACCESS_TOKEN)).toBe('at');
    expect(secrets.data.get(SECRET_REFRESH_TOKEN)).toBe('rt');
    const expiry = secrets.data.get(SECRET_ACCESS_EXPIRY);
    expect(expiry).toBeTruthy();
    expect(Number(expiry)).toBeGreaterThan(Date.now());
  });

  it('does not overwrite refresh token when absent', async () => {
    const secrets = mockSecrets({ [SECRET_REFRESH_TOKEN]: 'keep-me' });
    await storeTokens({ access_token: 'at', expires_in: 3600 }, secrets);

    expect(secrets.data.get(SECRET_REFRESH_TOKEN)).toBe('keep-me');
  });
});

describe('getAccessToken', () => {
  it('returns a cached access token when still valid', async () => {
    const secrets = mockSecrets({
      [SECRET_ACCESS_TOKEN]: 'cached-at',
      [SECRET_ACCESS_EXPIRY]: String(Date.now() + 3600_000),
    });
    const http = mockHttp();

    const token = await getAccessToken('cid', http, secrets);

    expect(token).toBe('cached-at');
    expect(http.post).not.toHaveBeenCalled();
  });

  it('refreshes when there is no valid cache', async () => {
    const secrets = mockSecrets({ [SECRET_REFRESH_TOKEN]: 'rt' });
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>).mockResolvedValue({
      access_token: 'new-at',
      refresh_token: 'new-rt',
      expires_in: 3600,
    });

    const token = await getAccessToken('cid', http, secrets);

    expect(token).toBe('new-at');
    expect(secrets.data.get(SECRET_ACCESS_TOKEN)).toBe('new-at');
    expect(secrets.data.get(SECRET_REFRESH_TOKEN)).toBe('new-rt');
  });

  it('throws when not connected', async () => {
    const secrets = mockSecrets();
    const http = mockHttp();

    await expect(getAccessToken('cid', http, secrets)).rejects.toThrow();
  });
});

describe('runDeviceCodeLogin', () => {
  it('runs the full device-code login and stores tokens', async () => {
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>)
      .mockResolvedValueOnce({
        device_code: 'dc',
        user_code: 'ABC-123',
        verification_uri: 'https://microsoft.com/devicelogin',
        expires_in: 900,
        interval: 5,
      })
      .mockResolvedValueOnce({ access_token: 'at', refresh_token: 'rt', expires_in: 3600 });
    const dialog = { show: vi.fn(async () => true) };
    const secrets = mockSecrets();

    await runDeviceCodeLogin('cid', http, dialog, secrets);

    expect(dialog.show).toHaveBeenCalledWith('ABC-123', 'https://microsoft.com/devicelogin');
    expect(secrets.data.get(SECRET_REFRESH_TOKEN)).toBe('rt');
  });

  it('throws when the dialog is cancelled', async () => {
    const http = mockHttp();
    (http.post as ReturnType<typeof vi.fn>).mockResolvedValue({
      device_code: 'dc',
      user_code: 'ABC-123',
      verification_uri: 'https://microsoft.com/devicelogin',
      expires_in: 900,
      interval: 5,
    });
    const dialog = { show: vi.fn(async () => false) };
    const secrets = mockSecrets();

    await expect(runDeviceCodeLogin('cid', http, dialog, secrets)).rejects.toThrow();
    expect(secrets.data.get(SECRET_REFRESH_TOKEN)).toBeUndefined();
  });
});

describe('withHostDbRetry', () => {
  const closingError = new Error(
    "Failed to execute 'transaction' on 'IDBDatabase': The database connection is closing.",
  );

  it('returns the value on first success without retrying', async () => {
    const op = vi.fn(async () => 'ok');

    await expect(withHostDbRetry(op, 2, 0)).resolves.toBe('ok');
    expect(op).toHaveBeenCalledTimes(1);
  });

  it('retries a transient failure then succeeds', async () => {
    const op = vi
      .fn<() => Promise<string>>()
      .mockRejectedValueOnce(closingError)
      .mockResolvedValueOnce('ok');

    await expect(withHostDbRetry(op, 2, 0)).resolves.toBe('ok');
    expect(op).toHaveBeenCalledTimes(2);
  });

  it('rethrows the last error after exhausting retries', async () => {
    const op = vi.fn(async () => {
      throw closingError;
    });

    await expect(withHostDbRetry(op, 2, 0)).rejects.toBe(closingError);
    expect(op).toHaveBeenCalledTimes(3);
  });
});

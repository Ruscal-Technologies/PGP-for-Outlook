import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

// session-cache.js keeps its cache/listener list in module-level state with no
// reset export, so each test gets a fresh module instance via resetModules()
// + a dynamic re-import — otherwise onSessionCleared() listeners registered in
// one test would still fire in the next.
let sessionCache;

beforeEach(async () => {
  vi.resetModules();
  sessionCache = await import('../web/js/pgp/session-cache.js');
});

afterEach(() => {
  vi.useRealTimers();
});

describe('cacheSessionKey / getSessionKey / hasSessionKey', () => {
  it('round-trips a cached key and its display metadata', () => {
    const { cacheSessionKey, getSessionKey, hasSessionKey, getSessionEmail, getSessionShortId } = sessionCache;

    expect(hasSessionKey()).toBe(false);
    expect(getSessionKey()).toBeNull();

    const fakeKey = { isPrivate: () => true };
    cacheSessionKey(fakeKey, 'alice@example.com', 'ABCD1234');

    expect(hasSessionKey()).toBe(true);
    expect(getSessionKey()).toBe(fakeKey);
    expect(getSessionEmail()).toBe('alice@example.com');
    expect(getSessionShortId()).toBe('ABCD1234');
  });

  it('defaults shortId to an empty string when omitted', () => {
    const { cacheSessionKey, getSessionShortId } = sessionCache;
    cacheSessionKey({}, 'bob@example.com');
    expect(getSessionShortId()).toBe('');
  });
});

describe('clearSessionKey', () => {
  it('clears cached state and fires registered listeners', () => {
    const { cacheSessionKey, clearSessionKey, hasSessionKey, getSessionEmail, onSessionCleared } = sessionCache;

    cacheSessionKey({}, 'alice@example.com');
    const listener = vi.fn();
    onSessionCleared(listener);

    clearSessionKey();

    expect(hasSessionKey()).toBe(false);
    expect(getSessionEmail()).toBeNull();
    expect(listener).toHaveBeenCalledTimes(1);
  });

  it('does not let a throwing listener prevent the cache from clearing', () => {
    const { cacheSessionKey, clearSessionKey, hasSessionKey, onSessionCleared } = sessionCache;

    cacheSessionKey({}, 'alice@example.com');
    onSessionCleared(() => { throw new Error('boom'); });

    expect(() => clearSessionKey()).not.toThrow();
    expect(hasSessionKey()).toBe(false);
  });
});

describe('inactivity timeout', () => {
  it('auto-clears the cache after CACHE_TIMEOUT_MS of inactivity', () => {
    vi.useFakeTimers();
    const { cacheSessionKey, hasSessionKey, CACHE_TIMEOUT_MS } = sessionCache;

    cacheSessionKey({}, 'alice@example.com');
    expect(hasSessionKey()).toBe(true);

    vi.advanceTimersByTime(CACHE_TIMEOUT_MS - 1);
    expect(hasSessionKey()).toBe(true);

    vi.advanceTimersByTime(1);
    expect(hasSessionKey()).toBe(false);
  });

  it('resets the timer on every getSessionKey() call, delaying expiry', () => {
    vi.useFakeTimers();
    const { cacheSessionKey, getSessionKey, hasSessionKey, CACHE_TIMEOUT_MS } = sessionCache;

    cacheSessionKey({}, 'alice@example.com');

    // Advance almost to expiry, then "use" the key — this should push the
    // expiry back out by a full CACHE_TIMEOUT_MS rather than letting it lapse.
    vi.advanceTimersByTime(CACHE_TIMEOUT_MS - 1);
    getSessionKey();
    vi.advanceTimersByTime(CACHE_TIMEOUT_MS - 1);

    expect(hasSessionKey()).toBe(true);
  });
});

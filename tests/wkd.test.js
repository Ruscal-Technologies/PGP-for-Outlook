import { describe, it, expect, vi, afterEach } from 'vitest';
import WKD, { encodeZBase32 } from '../web/js/wkd.js';

// Z-Base32 test vectors below are hand-derived by tracing encodeZBase32's own
// bit-shifting logic (see web/js/wkd.js) rather than pulled from an external
// spec — this is a regression guard against future changes to that logic, not
// a conformance test against the RFC 6189 reference implementation.
describe('encodeZBase32', () => {
  it('returns an empty string for empty input', () => {
    expect(encodeZBase32(new Uint8Array([]))).toBe('');
  });

  it('encodes a single zero byte', () => {
    expect(encodeZBase32(new Uint8Array([0x00]))).toBe('yy');
  });

  it('encodes a single 0xff byte', () => {
    expect(encodeZBase32(new Uint8Array([0xff]))).toBe('9h');
  });

  it('is deterministic for the same input', () => {
    const input = new Uint8Array([1, 2, 3, 4, 5]);
    expect(encodeZBase32(input)).toBe(encodeZBase32(input));
  });
});

// Computes the same SHA-1 + Z-Base32 hash WKD.lookup() computes internally,
// using the real (Node-native) WebCrypto API — this independently derives the
// expected URL hash segment instead of hardcoding a magic string, so the test
// stays correct even if the underlying digest algorithm choice ever changes
// to something we'd want to catch.
async function expectedWkdHash(localPart) {
  const encoded = new TextEncoder().encode(localPart.toLowerCase());
  const digest = new Uint8Array(await globalThis.crypto.subtle.digest('SHA-1', encoded));
  return encodeZBase32(digest);
}

afterEach(() => {
  delete global.fetch;
});

describe('WKD.lookup', () => {
  it('queries the advanced (openpgpkey subdomain) URL first and returns the raw bytes on success', async () => {
    const expectedHash = await expectedWkdHash('alice');
    const responseBytes = new Uint8Array([1, 2, 3]);
    global.fetch = vi.fn().mockResolvedValue({ status: 200, arrayBuffer: async () => responseBytes.buffer });

    const wkd = new WKD();
    const result = await wkd.lookup({ email: 'alice@example.com' });

    expect(global.fetch).toHaveBeenCalledTimes(1);
    expect(global.fetch.mock.calls[0][0]).toBe(
      `https://openpgpkey.example.com/.well-known/openpgpkey/example.com/hu/${expectedHash}?l=alice`
    );
    expect(Array.from(result)).toEqual([1, 2, 3]);
  });

  it('falls back to the direct URL when the advanced lookup fails', async () => {
    const expectedHash = await expectedWkdHash('bob');
    const responseBytes = new Uint8Array([9, 9]);
    global.fetch = vi.fn()
      .mockResolvedValueOnce({ status: 404, statusText: 'Not Found' })
      .mockResolvedValueOnce({ status: 200, arrayBuffer: async () => responseBytes.buffer });

    const wkd = new WKD();
    const result = await wkd.lookup({ email: 'bob@example.com' });

    expect(global.fetch).toHaveBeenCalledTimes(2);
    expect(global.fetch.mock.calls[1][0]).toBe(
      `https://example.com/.well-known/openpgpkey/hu/${expectedHash}?l=bob`
    );
    expect(Array.from(result)).toEqual([9, 9]);
  });

  it('throws when both the advanced and direct lookups fail', async () => {
    global.fetch = vi.fn().mockResolvedValue({ status: 404, statusText: 'Not Found' });

    const wkd = new WKD();
    await expect(wkd.lookup({ email: 'nobody@example.com' })).rejects.toThrow(/Direct WKD lookup failed/);
    expect(global.fetch).toHaveBeenCalledTimes(2);
  });

  it('lowercases the local part only for hashing — the query parameter keeps the original case', async () => {
    const expectedHash = await expectedWkdHash('Alice'); // hashing lowercases internally
    global.fetch = vi.fn().mockResolvedValue({ status: 200, arrayBuffer: async () => new Uint8Array([1]).buffer });

    const wkd = new WKD();
    await wkd.lookup({ email: 'Alice@example.com' });

    expect(global.fetch.mock.calls[0][0]).toBe(
      `https://openpgpkey.example.com/.well-known/openpgpkey/example.com/hu/${expectedHash}?l=Alice`
    );
  });

  it('throws when no email is provided', async () => {
    // These two validation checks run before any fetch call, but the WKD
    // constructor itself checks globalThis.fetch and falls back to
    // require('node-fetch') when it's missing — stub it so construction
    // doesn't fail on an unrelated missing dependency in this test env.
    global.fetch = vi.fn();
    const wkd = new WKD();
    await expect(wkd.lookup({})).rejects.toThrow(/must provide an email/);
  });

  it('throws for an email missing an @ sign', async () => {
    global.fetch = vi.fn();
    const wkd = new WKD();
    await expect(wkd.lookup({ email: 'not-an-email' })).rejects.toThrow(/Invalid e-mail address/);
  });
});

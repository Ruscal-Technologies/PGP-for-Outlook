import { describe, it, expect } from 'vitest';
import { encodeZBase32 } from '../web/js/wkd.js';

// Z-Base32 test vectors below are hand-derived by tracing encodeZBase32's own
// bit-shifting logic (see web/js/wkd.js) rather than pulled from an external
// spec — this is a regression guard against future changes to that logic, not
// a conformance test against the RFC 6189 reference implementation. Deep
// `WKD.lookup()` behavior (hashing, URL building, fetch/parse) is out of scope
// for this pass — see CLAUDE.md's test-suite section for the noted gap.
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

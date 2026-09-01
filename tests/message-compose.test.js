// @vitest-environment jsdom
//
// stripPgpArmorBlock() walks/mutates/re-serializes a real DOM (parses HTML,
// sets textContent on specific nodes, reads innerHTML back) -- a hand-rolled
// fake risks subtly diverging from real browser HTML parsing/serialization
// semantics for exactly the function whose correctness matters most in this
// feature, so this file runs under jsdom (real DOM) rather than this repo's
// usual small hand-rolled stubs (see tests/quoted-content.test.js for why
// those suffice elsewhere: no mutate-and-reserialize step there).
import { describe, it, expect, beforeEach, vi } from 'vitest';

let stripPgpArmorBlock;

beforeEach(async () => {
  // MessageCompose.js calls Office.onReady(...) at module load time; the
  // no-op stub (matching tests/message-read-popout.test.js's convention)
  // means that callback body never actually runs under test -- irrelevant
  // here, since stripPgpArmorBlock and the module-level BroadcastChannel
  // handoff listener (tested separately below) don't depend on it.
  global.Office = { onReady: () => {} };
  ({ stripPgpArmorBlock } = await import('../web/MessageCompose.js'));
});

const ARMOR = '-----BEGIN PGP MESSAGE-----\nVersion: Test\n\nabc123==\n-----END PGP MESSAGE-----';

describe('stripPgpArmorBlock', () => {
  it('removes an armor block contained in a single text node', () => {
    const html = `<div>before-text ${ARMOR} after-text</div>`;
    const { found, before, after } = stripPgpArmorBlock(html);

    expect(found).toBe(true);
    expect(before).toContain('before-text');
    expect(before).not.toContain('BEGIN PGP MESSAGE');
    expect(after).toContain('after-text');
    expect(after).not.toContain('END PGP MESSAGE');
  });

  it('removes an armor block split across <br>-joined lines', () => {
    const lines = ARMOR.split('\n').join('<br>');
    const html = `<div>before-text<br>${lines}<br>after-text</div>`;
    const { found, before, after } = stripPgpArmorBlock(html);

    expect(found).toBe(true);
    expect(before).toContain('before-text');
    expect(after).toContain('after-text');
    expect(before + after).not.toContain('BEGIN PGP MESSAGE');
  });

  it('removes an armor block inside a <pre> (matches this add-in\'s own setBodyAsync wrapping)', () => {
    const html = `<html><body><div>before-text</div><pre style="white-space:pre-wrap;">${ARMOR}</pre><div>after-text</div></body></html>`;
    const { found, before, after } = stripPgpArmorBlock(html);

    expect(found).toBe(true);
    expect(before).toContain('before-text');
    expect(after).toContain('after-text');
    expect(before + after).not.toContain('BEGIN PGP MESSAGE');
  });

  it('removes an armor block nested inside a blockquote, preserving unrelated sibling content', () => {
    const html = `<div><p>Reply text</p><blockquote><p>unrelated quote line</p><div>${ARMOR}</div></blockquote></div>`;
    const { found, before, after } = stripPgpArmorBlock(html);

    expect(found).toBe(true);
    expect(before).toContain('Reply text');
    expect(before).toContain('unrelated quote line');
    expect(before + after).not.toContain('BEGIN PGP MESSAGE');
  });

  it('returns found:false when no armor block is present, without mutating anything', () => {
    const html = '<div>just a normal reply, nothing encrypted here</div>';
    const result = stripPgpArmorBlock(html);
    expect(result).toEqual({ found: false });
  });

  it('returns found:false when BEGIN is present but END never appears', () => {
    const html = '<div>-----BEGIN PGP MESSAGE-----\nabc123 (truncated, no end marker)</div>';
    const result = stripPgpArmorBlock(html);
    expect(result).toEqual({ found: false });
  });
});

describe('reply handoff (BroadcastChannel)', () => {
  it('acks a matching handoff broadcast and splices the decrypted content into the body, replacing the armor', async () => {
    let savedBody = null;
    global.Office = {
      onReady: () => {},
      CoercionType: { Html: 'html', Text: 'text' },
      AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
      context: {
        mailbox: {
          item: {
            body: {
              getAsync: (_coercionType, cb) => cb({
                status: 'succeeded',
                value: `<div>Reply header info</div><div>${ARMOR}</div>`,
              }),
              setAsync: (html, _options, cb) => { savedBody = html; cb({ status: 'succeeded' }); },
            },
          },
        },
      },
    };
    document.body.innerHTML = '<div id="status-bar" class="pgp-hidden"></div>';

    vi.resetModules();
    await import('../web/MessageCompose.js');

    const sender = new BroadcastChannel('pgp_reply_handoff');
    const acked = new Promise((resolve) => {
      sender.onmessage = (event) => {
        if (event.data?.type === 'pgp-reply-handoff-ack') resolve(event.data.token);
      };
    });
    sender.postMessage({ type: 'pgp-reply-handoff', token: 'test-token-1', text: 'the decrypted message', isHtml: false });

    await expect(acked).resolves.toBe('test-token-1');
    // applyReplyHandoff's body-write happens after the ack in the same async
    // flow -- give its pending promises a turn to settle before asserting.
    await new Promise((resolve) => setTimeout(resolve, 0));

    sender.close();
    expect(savedBody).toContain('Reply header info');
    expect(savedBody).toContain('the decrypted message');
    expect(savedBody).not.toContain('BEGIN PGP MESSAGE');
  });

  it('leaves the body untouched and warns when no armor block is found in the native quote', async () => {
    let setAsyncCalled = false;
    global.Office = {
      onReady: () => {},
      CoercionType: { Html: 'html', Text: 'text' },
      AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
      context: {
        mailbox: {
          item: {
            body: {
              getAsync: (_coercionType, cb) => cb({ status: 'succeeded', value: '<div>no armor here</div>' }),
              setAsync: (_html, _options, cb) => { setAsyncCalled = true; cb({ status: 'succeeded' }); },
            },
          },
        },
      },
    };
    const statusEl = document.createElement('div');
    statusEl.id = 'status-bar';
    statusEl.className = 'pgp-hidden';
    document.body.innerHTML = '';
    document.body.appendChild(statusEl);

    vi.resetModules();
    await import('../web/MessageCompose.js');

    const sender = new BroadcastChannel('pgp_reply_handoff');
    const acked = new Promise((resolve) => {
      sender.onmessage = (event) => {
        if (event.data?.type === 'pgp-reply-handoff-ack') resolve(event.data.token);
      };
    });
    sender.postMessage({ type: 'pgp-reply-handoff', token: 'test-token-2', text: 'decrypted text', isHtml: false });

    await expect(acked).resolves.toBe('test-token-2');
    await new Promise((resolve) => setTimeout(resolve, 0));

    sender.close();
    expect(setAsyncCalled).toBe(false);
    expect(statusEl.textContent).toContain('Could not find the encrypted message');
  });
});

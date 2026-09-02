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

  it('still splits correctly when the input already contains the base splice marker text literally', () => {
    // The internal marker is only an implementation detail, but the input is
    // attacker-influenceable PGP message content -- a literal collision must
    // not corrupt the split (e.g. drop content, or leak the marker itself).
    const html = `<div>before __PGP_ARMOR_SPLICE__ text ${ARMOR} after __PGP_ARMOR_SPLICE__ text</div>`;
    const { found, before, after } = stripPgpArmorBlock(html);

    expect(found).toBe(true);
    expect(before).toContain('before __PGP_ARMOR_SPLICE__ text');
    expect(after).toContain('after __PGP_ARMOR_SPLICE__ text');
    expect(before + after).not.toContain('BEGIN PGP MESSAGE');
  });
});

describe('reply handoff (BroadcastChannel)', () => {
  const CONVERSATION_ID = 'conversation-test-1';

  // Waits for `promise` to settle, or for `ms` to pass with nothing
  // happening -- used to positively assert "no ack arrives" without waiting
  // out a full real timeout.
  function raceTimeout(promise, ms) {
    return Promise.race([
      promise.then((v) => ({ settled: true, value: v })),
      new Promise((resolve) => setTimeout(() => resolve({ settled: false }), ms)),
    ]);
  }

  function makeOfficeStub({ bodyHtml, composeType }) {
    let savedBody = null;
    let setAsyncCalled = false;
    const office = {
      onReady: () => {},
      CoercionType: { Html: 'html', Text: 'text' },
      AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
      MailboxEnums: { ComposeType: { Reply: 'reply', NewMail: 'newMail', Forward: 'forward' } },
      context: {
        mailbox: {
          item: {
            conversationId: CONVERSATION_ID,
            getComposeTypeAsync: (cb) => cb({ status: 'succeeded', value: { composeType } }),
            body: {
              getAsync: (_coercionType, cb) => cb({ status: 'succeeded', value: bodyHtml }),
              setAsync: (html, _options, cb) => { savedBody = html; setAsyncCalled = true; cb({ status: 'succeeded' }); },
            },
          },
        },
      },
    };
    return { office, getSavedBody: () => savedBody, wasSetAsyncCalled: () => setAsyncCalled };
  }

  beforeEach(() => {
    document.body.innerHTML = '<div id="status-bar" class="pgp-hidden"></div>';
  });

  it('acks a matching handoff broadcast and splices the decrypted content into the body, replacing the armor', async () => {
    const { office, getSavedBody } = makeOfficeStub({
      bodyHtml: `<div>Reply header info</div><div>${ARMOR}</div>`,
      composeType: 'reply',
    });
    global.Office = office;

    vi.resetModules();
    const { setupReplyHandoffListener } = await import('../web/MessageCompose.js');
    await setupReplyHandoffListener(true); // has110=true -> exercises the getComposeTypeAsync gate

    const { getReplyHandoffChannelName } = await import('../web/js/pgp/reply-handoff-channel.js');
    const sender = new BroadcastChannel(getReplyHandoffChannelName(CONVERSATION_ID));
    const acked = new Promise((resolve) => {
      sender.onmessage = (event) => {
        if (event.data?.type === 'pgp-reply-handoff-ack') resolve(event.data.token);
      };
    });
    sender.postMessage({ type: 'pgp-reply-handoff', token: 'test-token-1', text: 'the decrypted message', isHtml: false });

    await expect(acked).resolves.toBe('test-token-1');
    sender.close();

    const savedBody = getSavedBody();
    expect(savedBody).toContain('Reply header info');
    expect(savedBody).toContain('the decrypted message');
    expect(savedBody).not.toContain('BEGIN PGP MESSAGE');
  });

  it('does not ack, and does not write the body, when no armor block is found (so MessageRead.js\'s fallback can still trigger)', async () => {
    const { office, wasSetAsyncCalled } = makeOfficeStub({
      bodyHtml: '<div>no armor here</div>',
      composeType: 'reply',
    });
    global.Office = office;

    vi.resetModules();
    const { setupReplyHandoffListener } = await import('../web/MessageCompose.js');
    await setupReplyHandoffListener(true);

    const { getReplyHandoffChannelName } = await import('../web/js/pgp/reply-handoff-channel.js');
    const sender = new BroadcastChannel(getReplyHandoffChannelName(CONVERSATION_ID));
    const acked = new Promise((resolve) => {
      sender.onmessage = (event) => {
        if (event.data?.type === 'pgp-reply-handoff-ack') resolve(event.data.token);
      };
    });
    sender.postMessage({ type: 'pgp-reply-handoff', token: 'test-token-2', text: 'decrypted text', isHtml: false });

    const result = await raceTimeout(acked, 300);
    sender.close();

    expect(result.settled).toBe(false); // no ack -- read pane's timeout fallback must still be able to fire
    expect(wasSetAsyncCalled()).toBe(false);
    expect(document.getElementById('status-bar').textContent).toContain('Could not find the encrypted message');
  });

  it('does not ack when the body write itself fails', async () => {
    const office = {
      onReady: () => {},
      CoercionType: { Html: 'html', Text: 'text' },
      AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
      MailboxEnums: { ComposeType: { Reply: 'reply', NewMail: 'newMail', Forward: 'forward' } },
      context: {
        mailbox: {
          item: {
            conversationId: CONVERSATION_ID,
            getComposeTypeAsync: (cb) => cb({ status: 'succeeded', value: { composeType: 'reply' } }),
            body: {
              getAsync: (_coercionType, cb) => cb({ status: 'succeeded', value: `<div>${ARMOR}</div>` }),
              setAsync: (_html, _options, cb) => cb({ status: 'failed', error: { message: 'simulated setAsync failure' } }),
            },
          },
        },
      },
    };
    global.Office = office;

    vi.resetModules();
    const { setupReplyHandoffListener } = await import('../web/MessageCompose.js');
    await setupReplyHandoffListener(true);

    const { getReplyHandoffChannelName } = await import('../web/js/pgp/reply-handoff-channel.js');
    const sender = new BroadcastChannel(getReplyHandoffChannelName(CONVERSATION_ID));
    const acked = new Promise((resolve) => {
      sender.onmessage = (event) => {
        if (event.data?.type === 'pgp-reply-handoff-ack') resolve(event.data.token);
      };
    });
    sender.postMessage({ type: 'pgp-reply-handoff', token: 'test-token-3', text: 'decrypted text', isHtml: false });

    const result = await raceTimeout(acked, 300);
    sender.close();

    expect(result.settled).toBe(false);
    expect(document.getElementById('status-bar').textContent).toContain('Could not automatically insert');
  });

  it('never sets up a listener at all for a non-reply compose window (newMail/forward)', async () => {
    const { office, wasSetAsyncCalled } = makeOfficeStub({
      bodyHtml: `<div>${ARMOR}</div>`,
      composeType: 'newMail',
    });
    global.Office = office;

    vi.resetModules();
    const { setupReplyHandoffListener } = await import('../web/MessageCompose.js');
    await setupReplyHandoffListener(true);

    const { getReplyHandoffChannelName } = await import('../web/js/pgp/reply-handoff-channel.js');
    const sender = new BroadcastChannel(getReplyHandoffChannelName(CONVERSATION_ID));
    const acked = new Promise((resolve) => {
      sender.onmessage = (event) => {
        if (event.data?.type === 'pgp-reply-handoff-ack') resolve(event.data.token);
      };
    });
    sender.postMessage({ type: 'pgp-reply-handoff', token: 'test-token-4', text: 'decrypted text', isHtml: false });

    const result = await raceTimeout(acked, 300);
    sender.close();

    expect(result.settled).toBe(false);
    expect(wasSetAsyncCalled()).toBe(false);
  });
});

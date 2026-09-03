// @vitest-environment jsdom
//
// armReplyHandoffListener() lives in web/js/pgp/reply-handoff-runtime-core.js
// and is shared by MessageCompose.js's own listener (tests/message-compose.test.js
// already covers the ack/timeout/warning behavior end-to-end) and
// web/ReplyHandoffPane.js (the dedicated minimal pane, #22). This file only
// covers the piece message-compose.test.js doesn't: the `onSettled` callback,
// which the pane uses to decide when to close itself. jsdom is needed because
// a successful splice still goes through stripPgpArmorBlock's real DOM walk.
import { describe, it, expect } from 'vitest';
import { armReplyHandoffListener } from '../web/js/pgp/reply-handoff-runtime-core.js';
import { HANDOFF_PENDING_MARKER } from '../web/js/pgp/reply-handoff-channel.js';

const ARMOR = '-----BEGIN PGP MESSAGE-----\nVersion: Test\n\nabc123==\n-----END PGP MESSAGE-----';

function makeOfficeStub({ bodyHtml, composeType, conversationId }) {
  let savedBody = null;
  return {
    office: {
      onReady: () => {},
      CoercionType: { Html: 'html', Text: 'text' },
      AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
      MailboxEnums: { ComposeType: { Reply: 'reply', ReplyAll: 'replyAll', NewMail: 'newMail', Forward: 'forward' } },
      context: {
        mailbox: {
          item: {
            conversationId,
            getComposeTypeAsync: (cb) => cb({ status: 'succeeded', value: { composeType } }),
            body: {
              getAsync: (_coercionType, cb) => cb({ status: 'succeeded', value: bodyHtml }),
              setAsync: (html, _options, cb) => { savedBody = html; cb({ status: 'succeeded' }); },
            },
          },
        },
      },
    },
    getSavedBody: () => savedBody,
  };
}

describe('armReplyHandoffListener — onSettled', () => {
  it('calls onSettled with success:true once the handoff acks', async () => {
    const conversationId = 'conv-settled-success';
    const { office } = makeOfficeStub({
      bodyHtml: `<div>${ARMOR}</div>`,
      composeType: 'reply',
      conversationId,
    });
    global.Office = office;

    const { getReplyHandoffChannelName } = await import('../web/js/pgp/reply-handoff-channel.js');
    const settled = new Promise((resolve) => {
      armReplyHandoffListener({ has110: true, has114: false, onSettled: resolve });
    });

    // Give armReplyHandoffListener's internal await a tick to actually arm
    // the channel before the sender posts to it.
    await new Promise((r) => setTimeout(r, 10));
    const sender = new BroadcastChannel(getReplyHandoffChannelName(conversationId));
    sender.postMessage({ type: 'pgp-reply-handoff', token: 'tok-1', text: 'decrypted', isHtml: false });

    const result = await settled;
    sender.close();

    expect(result).toEqual({ success: true, message: 'Decrypted message inserted into this reply.' });
  });

  it('strips the HANDOFF_PENDING_MARKER (prepended by MessageRead.js\'s displayReplyForm/displayReplyAllForm call) as part of a successful splice', async () => {
    const conversationId = 'conv-settled-marker';
    const { office, getSavedBody } = makeOfficeStub({
      bodyHtml: `<div>${HANDOFF_PENDING_MARKER}</div><div>${ARMOR}</div>`,
      composeType: 'reply',
      conversationId,
    });
    global.Office = office;

    const { getReplyHandoffChannelName } = await import('../web/js/pgp/reply-handoff-channel.js');
    const settled = new Promise((resolve) => {
      armReplyHandoffListener({ has110: true, has114: false, onSettled: resolve });
    });

    await new Promise((r) => setTimeout(r, 10));
    const sender = new BroadcastChannel(getReplyHandoffChannelName(conversationId));
    sender.postMessage({ type: 'pgp-reply-handoff', token: 'tok-marker', text: 'decrypted', isHtml: false });

    const result = await settled;
    sender.close();

    expect(result.success).toBe(true);
    expect(getSavedBody()).not.toContain(HANDOFF_PENDING_MARKER);
    expect(getSavedBody()).toContain('decrypted');
  });

  it('calls onSettled with success:false immediately when the compose window is not a reply', async () => {
    global.Office = makeOfficeStub({ bodyHtml: '', composeType: 'newMail', conversationId: 'conv-x' }).office;

    const result = await new Promise((resolve) => {
      armReplyHandoffListener({ has110: true, has114: false, onSettled: resolve });
    });

    expect(result.success).toBe(false);
    expect(result.message).toMatch(/isn't a reply/i);
  });

  it('calls onSettled with success:false immediately when there is no conversationId (or inReplyTo)', async () => {
    global.Office = makeOfficeStub({ bodyHtml: '', composeType: 'reply', conversationId: undefined }).office;

    const result = await new Promise((resolve) => {
      armReplyHandoffListener({ has110: true, has114: false, onSettled: resolve });
    });

    expect(result.success).toBe(false);
    expect(result.message).toMatch(/no conversation/i);
  });

  it('calls onSettled with success:false immediately when BroadcastChannel is unavailable', async () => {
    global.Office = makeOfficeStub({ bodyHtml: '', composeType: 'reply', conversationId: 'conv-y' }).office;
    const realBroadcastChannel = global.BroadcastChannel;
    // @ts-ignore -- deliberately simulate an environment without it (e.g. the
    // classic-Windows JS-only runtime, confirmed to lack it live).
    delete global.BroadcastChannel;

    const result = await new Promise((resolve) => {
      armReplyHandoffListener({ has110: true, has114: false, onSettled: resolve });
    });

    global.BroadcastChannel = realBroadcastChannel;

    expect(result.success).toBe(false);
    expect(result.message).toMatch(/broadcastchannel/i);
  });
});

import { describe, it, expect, vi } from 'vitest';

// Minimal stub matching tests/message-read-popout.test.js's conventions —
// only what openNativeReplyWithHandoff / openReplyComposeForm touch.
function installStubs({ conversationId, internetMessageId }) {
  const statusEl = { className: '', textContent: '', classList: { remove: vi.fn(), add: vi.fn() }, appendChild: vi.fn() };
  global.document = {
    getElementById: (id) => (id === 'status-bar' ? statusEl : null),
    createElement: () => ({}),
    createTextNode: () => ({}),
  };
  const displayReplyForm = vi.fn();
  const displayReplyAllForm = vi.fn();
  const displayNewMessageFormAsync = vi.fn((_formData, cb) => cb({ status: 'succeeded' }));
  global.Office = {
    onReady: () => {},
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    context: {
      mailbox: {
        item: { conversationId, internetMessageId, displayReplyForm, displayReplyAllForm },
        displayNewMessageFormAsync,
      },
    },
  };
  return { statusEl, displayReplyForm, displayReplyAllForm, displayNewMessageFormAsync };
}

let openNativeReplyWithHandoff;

describe('openNativeReplyWithHandoff — missing conversationId', () => {
  it('skips straight to the existing displayNewMessageForm fallback, never opening a native reply, when neither conversationId nor internetMessageId is available', async () => {
    const { statusEl, displayReplyForm, displayReplyAllForm, displayNewMessageFormAsync } =
      installStubs({ conversationId: undefined, internetMessageId: undefined });
    ({ openNativeReplyWithHandoff } = await import('../web/MessageRead.js'));

    openNativeReplyWithHandoff(false, ['a@example.com'], [], 'Re: hi', '<p>quoted</p>');

    // Never opens the native reply we couldn't safely hand plaintext to.
    expect(displayReplyForm).not.toHaveBeenCalled();
    expect(displayReplyAllForm).not.toHaveBeenCalled();
    // Falls back to the existing, already-working path instead.
    expect(displayNewMessageFormAsync).toHaveBeenCalledTimes(1);
    expect(statusEl.textContent).toContain('backup reply window');
  });

  it('still attempts the native reply + handoff, using internetMessageId as the scoping ID, when conversationId is missing but internetMessageId is available', async () => {
    const internetMessageId = '<abc123@example.com>';
    const { displayReplyForm, displayReplyAllForm } =
      installStubs({ conversationId: undefined, internetMessageId });
    ({ openNativeReplyWithHandoff } = await import('../web/MessageRead.js'));
    const { getReplyHandoffChannelName } = await import('../web/js/pgp/reply-handoff-channel.js');

    openNativeReplyWithHandoff(true, ['a@example.com'], [], 'Re: hi', '<p>quoted</p>');

    // Opens the native reply (Reply All, per the `true` argument) instead of
    // falling back immediately -- internetMessageId was enough to scope the
    // channel safely.
    expect(displayReplyAllForm).toHaveBeenCalledTimes(1);
    expect(displayReplyForm).not.toHaveBeenCalled();

    // Confirm it's actually broadcasting on the internetMessageId-derived
    // channel (not the shared base channel), and ack it so the retry
    // interval/timeout inside openNativeReplyWithHandoff clean themselves up
    // promptly instead of running for the full REPLY_HANDOFF_TIMEOUT_MS.
    const channelName = getReplyHandoffChannelName(internetMessageId);
    const probe = new BroadcastChannel(channelName);
    const handoff = await new Promise((resolve) => {
      probe.onmessage = (event) => {
        if (event.data?.type === 'pgp-reply-handoff') resolve(event.data);
      };
    });
    probe.postMessage({ type: 'pgp-reply-handoff-ack', token: handoff.token });
    // Give openNativeReplyWithHandoff's own onmessage handler a turn to
    // process the ack (closes its channel, clears its timers) before the
    // test ends, so nothing fires asynchronously after teardown.
    await new Promise((resolve) => setTimeout(resolve, 20));
    probe.close();

    expect(handoff.type).toBe('pgp-reply-handoff');
  });
});

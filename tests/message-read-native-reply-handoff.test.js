import { describe, it, expect, vi } from 'vitest';

// Minimal stub matching tests/message-read-popout.test.js's conventions —
// only what openNativeReplyWithHandoff / openReplyComposeForm touch.
function installStubs({ conversationId, internetMessageId }) {
  const statusEl = { className: '', textContent: '', classList: { remove: vi.fn(), add: vi.fn() }, appendChild: vi.fn() };
  const replyBtn = { disabled: false };
  const replyAllBtn = { disabled: false };
  global.document = {
    getElementById: (id) => {
      if (id === 'status-bar') return statusEl;
      if (id === 'btn-reply-encrypted') return replyBtn;
      if (id === 'btn-reply-all-encrypted') return replyAllBtn;
      return null;
    },
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
  return { statusEl, replyBtn, replyAllBtn, displayReplyForm, displayReplyAllForm, displayNewMessageFormAsync };
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
    // No native reply was ever opened here, so the warning must not claim a
    // second window exists to close (see issue #16) -- this is the ONLY
    // window that opened.
    expect(statusEl.textContent).not.toContain('backup reply window');
    expect(statusEl.textContent).not.toContain('other blank reply window');
    expect(statusEl.textContent).toContain('this reply window was opened');
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
    // formData is a required parameter of both APIs (issue #20) -- calling
    // with zero arguments throws synchronously on every real invocation.
    expect(displayReplyAllForm).toHaveBeenCalledWith('');

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

describe('openNativeReplyWithHandoff — fallback warning accuracy (issue #16)', () => {
  it('warns without mentioning a second window when displayReplyForm/displayReplyAllForm itself throws', async () => {
    const { statusEl, displayReplyForm } = installStubs({ conversationId: 'conv-1', internetMessageId: undefined });
    displayReplyForm.mockImplementation(() => { throw new Error('boom'); });
    ({ openNativeReplyWithHandoff } = await import('../web/MessageRead.js'));

    openNativeReplyWithHandoff(false, ['a@example.com'], [], 'Re: hi', '<p>quoted</p>');

    expect(statusEl.textContent).not.toContain('backup reply window');
    expect(statusEl.textContent).not.toContain('other blank reply window');
    expect(statusEl.textContent).toContain('this reply window was opened');
  });

  it('warns and mentions the other open window when BroadcastChannel construction fails after the native reply already opened', async () => {
    const { statusEl } = installStubs({ conversationId: 'conv-2', internetMessageId: undefined });
    const originalBroadcastChannel = global.BroadcastChannel;
    global.BroadcastChannel = class {
      constructor() { throw new Error('no BroadcastChannel'); }
    };
    ({ openNativeReplyWithHandoff } = await import('../web/MessageRead.js'));

    try {
      openNativeReplyWithHandoff(false, ['a@example.com'], [], 'Re: hi', '<p>quoted</p>');
    } finally {
      global.BroadcastChannel = originalBroadcastChannel;
    }

    expect(statusEl.textContent).toContain('backup reply window');
    expect(statusEl.textContent).toContain('other blank reply window');
  });

  it('warns with the timeout-specific message and mentions the other open window when no ack ever arrives', async () => {
    vi.useFakeTimers();
    const { statusEl } = installStubs({ conversationId: 'conv-3', internetMessageId: undefined });
    ({ openNativeReplyWithHandoff } = await import('../web/MessageRead.js'));

    openNativeReplyWithHandoff(false, ['a@example.com'], [], 'Re: hi', '<p>quoted</p>');
    await vi.advanceTimersByTimeAsync(15000);
    vi.useRealTimers();

    expect(statusEl.textContent).toContain("didn't finish in time");
    expect(statusEl.textContent).toContain('other blank reply window');
  });
});

describe('openNativeReplyWithHandoff — reply buttons disabled while in flight (issue #17)', () => {
  it('disables the reply buttons once the native reply opens, and re-enables them once the ack arrives', async () => {
    const { replyBtn, replyAllBtn } = installStubs({ conversationId: 'conv-4', internetMessageId: undefined });
    ({ openNativeReplyWithHandoff } = await import('../web/MessageRead.js'));
    const { getReplyHandoffChannelName } = await import('../web/js/pgp/reply-handoff-channel.js');

    openNativeReplyWithHandoff(false, ['a@example.com'], [], 'Re: hi', '<p>quoted</p>');

    expect(replyBtn.disabled).toBe(true);
    expect(replyAllBtn.disabled).toBe(true);

    const probe = new BroadcastChannel(getReplyHandoffChannelName('conv-4'));
    const handoff = await new Promise((resolve) => {
      probe.onmessage = (event) => {
        if (event.data?.type === 'pgp-reply-handoff') resolve(event.data);
      };
    });
    probe.postMessage({ type: 'pgp-reply-handoff-ack', token: handoff.token });
    await new Promise((resolve) => setTimeout(resolve, 20));
    probe.close();

    expect(replyBtn.disabled).toBe(false);
    expect(replyAllBtn.disabled).toBe(false);
  });

  it('re-enables the reply buttons after a timed-out handoff falls back', async () => {
    vi.useFakeTimers();
    const { replyBtn, replyAllBtn } = installStubs({ conversationId: 'conv-5', internetMessageId: undefined });
    ({ openNativeReplyWithHandoff } = await import('../web/MessageRead.js'));

    openNativeReplyWithHandoff(false, ['a@example.com'], [], 'Re: hi', '<p>quoted</p>');
    expect(replyBtn.disabled).toBe(true);

    await vi.advanceTimersByTimeAsync(15000);
    vi.useRealTimers();

    expect(replyBtn.disabled).toBe(false);
    expect(replyAllBtn.disabled).toBe(false);
  });

  it('never disables the buttons when the handoff is skipped before a native reply opens (no scoping ID)', async () => {
    const { replyBtn, replyAllBtn } = installStubs({ conversationId: undefined, internetMessageId: undefined });
    ({ openNativeReplyWithHandoff } = await import('../web/MessageRead.js'));

    openNativeReplyWithHandoff(false, ['a@example.com'], [], 'Re: hi', '<p>quoted</p>');

    expect(replyBtn.disabled).toBe(false);
    expect(replyAllBtn.disabled).toBe(false);
  });
});

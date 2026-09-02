import { describe, it, expect, vi } from 'vitest';

// Minimal stub matching tests/message-read-popout.test.js's conventions —
// only what openNativeReplyWithHandoff / openReplyComposeForm touch.
function installStubs({ conversationId }) {
  const statusEl = { className: '', textContent: '', classList: { remove: vi.fn(), add: vi.fn() } };
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
        item: { conversationId, displayReplyForm, displayReplyAllForm },
        displayNewMessageFormAsync,
      },
    },
  };
  return { statusEl, displayReplyForm, displayReplyAllForm, displayNewMessageFormAsync };
}

let openNativeReplyWithHandoff;

describe('openNativeReplyWithHandoff — missing conversationId', () => {
  it('skips straight to the existing displayNewMessageForm fallback, never opening a native reply, when conversationId is missing', async () => {
    const { statusEl, displayReplyForm, displayReplyAllForm, displayNewMessageFormAsync } =
      installStubs({ conversationId: undefined });
    ({ openNativeReplyWithHandoff } = await import('../web/MessageRead.js'));

    openNativeReplyWithHandoff(false, ['a@example.com'], [], 'Re: hi', '<p>quoted</p>');

    // Never opens the native reply we couldn't safely hand plaintext to.
    expect(displayReplyForm).not.toHaveBeenCalled();
    expect(displayReplyAllForm).not.toHaveBeenCalled();
    // Falls back to the existing, already-working path instead.
    expect(displayNewMessageFormAsync).toHaveBeenCalledTimes(1);
    expect(statusEl.textContent).toContain('backup reply window');
  });
});

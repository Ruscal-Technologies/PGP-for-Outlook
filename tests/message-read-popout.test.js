import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

// MessageRead.js calls Office.onReady(...) at module load time (unlike the
// pgp/* modules, which only touch Office lazily inside function bodies), so
// it needs a minimal Office/document/window stub in place before import —
// none of which need to be a real DOM/Office runtime, just enough that the
// pure decision logic under test doesn't throw when it calls
// showStatus()/openDecryptedPopup() internally. BroadcastChannel is a real
// Node global in this runtime, so openDecryptedPopupDialog's handshake setup
// can run for real without needing a fake.
function installStubs() {
  const statusEl = { className: '', textContent: '', classList: { remove: vi.fn(), add: vi.fn() } };
  global.document = {
    getElementById: (id) => (id === 'status-bar' ? statusEl : null),
    createElement: () => ({}),
    createTextNode: () => ({}),
  };
  global.Office = { onReady: () => {} };
  // A truthy fake window, so openDecryptedPopup's success branch runs
  // instead of its own "popup blocked" fallback overwriting statusEl.
  const fakeChildWindow = {
    document: { open: vi.fn(), write: vi.fn(), close: vi.fn() },
    focus: vi.fn(),
  };
  const windowOpen = vi.fn(() => fakeChildWindow);
  global.window = { open: windowOpen, location: { href: 'https://example.test/MessageRead.html' } };
  return { statusEl, windowOpen, fakeChildWindow };
}

/**
 * Extends installStubs() with an Office.context.ui.displayDialogAsync that
 * succeeds synchronously, capturing the DialogEventReceived handler so tests
 * can simulate the dialog closing. This is exactly the "displayDialogAsync
 * succeeds, then the dialog closes" sequence the critical regression lived
 * in (onPopoutDialogClosed skipping cleanup on the ordinary 12006 path).
 */
function installDialogStubs() {
  const stubs = installStubs();
  // One fresh fake dialog object + captured handler per displayDialogAsync
  // call, in order -- mirroring how the real API hands back a distinct
  // dialog object each time it's called. This distinctness is exactly what
  // the "_popoutDialog !== dialog" stale-session guard checks, so reusing a
  // single fake object here would silently defeat the test.
  const dialogEventHandlers = [];
  global.Office.AsyncResultStatus = { Succeeded: 'succeeded', Failed: 'failed' };
  global.Office.EventType = { DialogMessageReceived: 'dialogMessageReceived', DialogEventReceived: 'dialogEventReceived' };
  global.Office.context = {
    ui: {
      displayDialogAsync: vi.fn((_url, _opts, callback) => {
        const dialog = { close: vi.fn(), addEventHandler: vi.fn((type, handler) => {
          if (type === 'dialogEventReceived') dialogEventHandlers.push(handler);
        }) };
        callback({ status: 'succeeded', value: dialog });
      }),
    },
  };
  return {
    ...stubs,
    // Defaults to the most recently opened dialog; pass an explicit index to
    // target an earlier (now-superseded) one instead.
    closeDialog: (arg, sessionIndex = dialogEventHandlers.length - 1) => dialogEventHandlers[sessionIndex](arg),
  };
}

let handleDialogOpenFailure;
let openDecryptedPopupDialog;

beforeEach(async () => {
  // Office.onReady is the only thing MessageRead.js touches at module load
  // time; document/window are only ever read lazily inside function bodies,
  // so each test installs its own fresh stubs right before calling in.
  global.Office = { onReady: () => {} };
  vi.resetModules();
  ({ handleDialogOpenFailure, openDecryptedPopupDialog } = await import('../web/MessageRead.js'));
});

afterEach(() => {
  vi.useRealTimers();
});

describe('handleDialogOpenFailure', () => {
  it('surfaces "already open" for code 12007 and does NOT fall back to the legacy popup', () => {
    const { statusEl, windowOpen } = installStubs();
    handleDialogOpenFailure({ code: 12007 }, 'decrypted text', false, 'Subject');

    expect(statusEl.textContent).toBe('A pop-out window is already open.');
    // The whole point of this branch: opening a second window underneath an
    // already-open dialog would be confusing, so window.open() must not fire.
    expect(windowOpen).not.toHaveBeenCalled();
  });

  it('falls back to the legacy popup for any other error code, with the original text/subject intact', () => {
    const { statusEl, windowOpen, fakeChildWindow } = installStubs();
    handleDialogOpenFailure({ code: 12002 }, 'decrypted text', false, 'Subject');

    expect(statusEl.textContent).toBe('Could not open the pop-out window as a dialog; opening a regular window instead.');
    expect(windowOpen).toHaveBeenCalledTimes(1);
    // Not just "a window opened" — confirm the original arguments actually
    // flowed all the way through to the fallback's rendered document.
    const html = fakeChildWindow.document.write.mock.calls[0][0];
    expect(html).toContain('decrypted text');
    expect(html).toContain('PGP Decrypted : Subject');
  });
});

describe('onPopoutDialogClosed (via openDecryptedPopupDialog)', () => {
  it('an ordinary close (12006) before the handshake completes cancels the pending timeout -- no stray fallback later', () => {
    vi.useFakeTimers();
    const { statusEl, windowOpen, closeDialog } = installDialogStubs();

    openDecryptedPopupDialog('plaintext secret', false, 'Subject');
    // Simulate the user closing the dialog before it ever signals readiness
    // over BroadcastChannel -- this is the exact sequence the regression
    // broke: cleanup was skipped on this path, so the handshake timer kept
    // running and fired an unprompted plaintext window ~10s later.
    closeDialog({ error: 12006 });

    vi.advanceTimersByTime(11000); // past the 10s handshake backstop

    expect(statusEl.textContent).toBe('');
    expect(windowOpen).not.toHaveBeenCalled();
  });

  it('a non-12006 close triggers the legacy-popup fallback immediately', () => {
    const { statusEl, windowOpen, closeDialog } = installDialogStubs();

    openDecryptedPopupDialog('plaintext secret', false, 'Subject');
    closeDialog({ error: 12002 }); // e.g. the dialog failed to load

    expect(statusEl.textContent).toBe('The pop-out window closed unexpectedly. Opening a regular window instead.');
    expect(windowOpen).toHaveBeenCalledTimes(1);
  });

  it('a stale close event for a superseded dialog session is ignored', () => {
    // The second session's own handshake timer is left pending (its dialog
    // is never closed in this test) -- fake timers keep that a no-op instead
    // of a real 10s Node timer outliving the test.
    vi.useFakeTimers();
    const { statusEl, windowOpen, closeDialog } = installDialogStubs();

    openDecryptedPopupDialog('plaintext secret', false, 'Subject');
    // A second attempt supersedes the first dialog's tracked reference...
    openDecryptedPopupDialog('plaintext secret 2', false, 'Subject 2');
    // ...so a straggling close event from the FIRST dialog (index 0) must be
    // a no-op, not act on the second (current) session's state.
    closeDialog({ error: 12002 }, 0);

    expect(statusEl.textContent).toBe('');
    expect(windowOpen).not.toHaveBeenCalled();
  });
});

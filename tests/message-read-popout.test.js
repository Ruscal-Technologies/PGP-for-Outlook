import { describe, it, expect, beforeEach, vi } from 'vitest';

// MessageRead.js calls Office.onReady(...) at module load time (unlike the
// pgp/* modules, which only touch Office lazily inside function bodies), so
// it needs a minimal Office/document/window stub in place before import —
// none of which need to be a real DOM/Office runtime, just enough that the
// pure decision logic under test (handleDialogOpenFailure) doesn't throw
// when it calls showStatus()/openDecryptedPopup() internally.
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
  global.window = { open: windowOpen };
  return { statusEl, windowOpen };
}

let handleDialogOpenFailure;

beforeEach(async () => {
  // Office.onReady is the only thing MessageRead.js touches at module load
  // time; document/window are only ever read lazily inside function bodies,
  // so each test installs its own fresh stubs right before calling in.
  global.Office = { onReady: () => {} };
  vi.resetModules();
  ({ handleDialogOpenFailure } = await import('../web/MessageRead.js'));
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

  it('falls back to the legacy popup for any other error code', () => {
    const { statusEl, windowOpen } = installStubs();
    handleDialogOpenFailure({ code: 12002 }, 'decrypted text', false, 'Subject');

    expect(statusEl.textContent).toBe('Could not open the pop-out window as a dialog; opening a regular window instead.');
    expect(windowOpen).toHaveBeenCalledTimes(1);
  });
});

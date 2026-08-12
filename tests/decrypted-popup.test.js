import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

// DecryptedPopup.js's module-load-time init() reads location.search and,
// only if a token is present, goes on to construct a BroadcastChannel and
// start a timer. Omitting the token here makes init() take its early-return
// branch (just calls showError) so import doesn't leave a live channel or
// timer behind — renderPayload, the function under test, is unaffected by
// that early return since it's a separate exported function.
function installStubs() {
  const els = {
    'popout-error': { textContent: '', classList: { remove: vi.fn(), add: vi.fn() } },
    'popout-text': { textContent: '', classList: { remove: vi.fn(), add: vi.fn() } },
    'popout-html-wrapper': { classList: { remove: vi.fn(), add: vi.fn() } },
    'popout-html-frame': { srcdoc: '' },
  };
  global.document = { getElementById: (id) => els[id], title: '' };
  global.location = { search: '' };
  return els;
}

let renderPayload;

beforeEach(async () => {
  installStubs();
  vi.resetModules();
  ({ renderPayload } = await import('../web/DecryptedPopup.js'));
});

afterEach(() => {
  vi.useRealTimers();
});

describe('renderPayload', () => {
  it('renders HTML payloads into the sandboxed iframe via srcdoc', () => {
    const els = installStubs();
    renderPayload({ text: '<b>hello</b>', isHtml: true, title: 'PGP Decrypted : Test' });

    expect(els['popout-html-frame'].srcdoc).toBe('<b>hello</b>');
    expect(els['popout-html-wrapper'].classList.remove).toHaveBeenCalledWith('pgp-hidden');
    expect(els['popout-text'].classList.remove).not.toHaveBeenCalled();
    expect(global.document.title).toBe('PGP Decrypted : Test');
  });

  it('renders plaintext payloads via textContent, not the HTML iframe', () => {
    const els = installStubs();
    renderPayload({ text: 'hello world', isHtml: false, title: 'PGP Decrypted : Test' });

    expect(els['popout-text'].textContent).toBe('hello world');
    expect(els['popout-text'].classList.remove).toHaveBeenCalledWith('pgp-hidden');
    expect(els['popout-html-wrapper'].classList.remove).not.toHaveBeenCalled();
  });

  it('falls back to a default title when none is provided', () => {
    renderPayload({ text: 'hello', isHtml: false, title: '' });
    expect(global.document.title).toBe('PGP Decrypted');
  });
});

/**
 * init() itself is an unexported IIFE that runs at module-load time, so
 * these tests can't call it directly -- they set up location.search/Office
 * stubs, import the module (which runs init() as a side effect), and drive
 * the rest through a real BroadcastChannel acting as the "parent", exactly
 * as MessageRead.js's openDecryptedPopupDialog does in production. Node's
 * BroadcastChannel is a real global in this runtime (confirmed separately),
 * so no fake is needed for the channel itself.
 */
describe('init() (the module-load-time handshake)', () => {
  function installOfficeStub() {
    const messageParent = vi.fn();
    global.Office = {
      onReady: () => Promise.resolve(),
      context: { ui: { messageParent } },
    };
    return { messageParent };
  }

  async function loadWithToken(token) {
    const els = installStubs();
    global.location = { search: `?token=${token}` };
    vi.resetModules();
    await import('../web/DecryptedPopup.js');
    return els;
  }

  it('completes the handshake and renders a valid payload delivered over BroadcastChannel', async () => {
    // The parent channel must exist and be subscribed BEFORE init() runs
    // (i.e. before the module import below) -- init()'s "dialog-listening"
    // broadcast fires synchronously during that import, and BroadcastChannel
    // does not queue messages for listeners that subscribe afterward.
    const parentChannel = new BroadcastChannel('pgp_popout_happy-path');
    const listening = new Promise((resolve) => {
      parentChannel.onmessage = (event) => {
        if (event.data?.type === 'dialog-listening') resolve();
      };
    });

    const els = await loadWithToken('happy-path');
    await listening;
    parentChannel.postMessage({ type: 'payload', text: '<b>secret</b>', isHtml: true, title: 'PGP Decrypted : Test' });
    await new Promise((resolve) => setTimeout(resolve, 10)); // let the message be delivered

    expect(els['popout-html-frame'].srcdoc).toBe('<b>secret</b>');
    expect(els['popout-html-wrapper'].classList.remove).toHaveBeenCalledWith('pgp-hidden');
    expect(els['popout-error'].textContent).toBe('');
    parentChannel.close();
  });

  it('rejects a payload whose text field is not a string, instead of rendering "undefined"', async () => {
    const parentChannel = new BroadcastChannel('pgp_popout_bad-payload');
    const listening = new Promise((resolve) => {
      parentChannel.onmessage = (event) => {
        if (event.data?.type === 'dialog-listening') resolve();
      };
    });

    const els = await loadWithToken('bad-payload');
    await listening;
    parentChannel.postMessage({ type: 'payload', text: 12345, isHtml: false, title: 'x' });
    await new Promise((resolve) => setTimeout(resolve, 10));

    expect(els['popout-error'].textContent).toBe('The decrypted content could not be displayed. Please try again.');
    expect(els['popout-text'].textContent).toBe(''); // never touched
    parentChannel.close();
  });

  it('shows a timeout error and notifies the parent if no payload arrives in time', async () => {
    vi.useFakeTimers();
    const { messageParent } = installOfficeStub();
    const els = await loadWithToken('timeout-case');

    await vi.advanceTimersByTimeAsync(8001); // past DecryptedPopup.js's own 8s budget

    expect(els['popout-error'].textContent).toBe('The decrypted content did not arrive in time. Please try again.');
    expect(messageParent).toHaveBeenCalledWith(JSON.stringify({ type: 'popout-error', reason: 'timeout' }));
  });

  it('shows an error and notifies the parent if BroadcastChannel is unavailable', async () => {
    const { messageParent } = installOfficeStub();
    const realBroadcastChannel = global.BroadcastChannel;
    global.BroadcastChannel = function () {
      throw new Error('BroadcastChannel disabled by policy');
    };
    try {
      const els = await loadWithToken('no-broadcast-channel');
      expect(els['popout-error'].textContent).toBe('This pop-out window could not connect to receive the decrypted content.');
      await Promise.resolve(); // flush notifyParent's Office.onReady().then(...) microtask
      expect(messageParent).toHaveBeenCalledWith(JSON.stringify({ type: 'popout-error', reason: 'broadcast-channel-unavailable' }));
    } finally {
      global.BroadcastChannel = realBroadcastChannel;
    }
  });

  it('shows an error without touching BroadcastChannel when no token is present', async () => {
    const els = await loadWithToken(''); // ?token= with an empty value
    expect(els['popout-error'].textContent).toBe('This pop-out window was opened without a valid token.');
  });
});

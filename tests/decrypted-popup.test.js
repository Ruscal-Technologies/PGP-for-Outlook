import { describe, it, expect, beforeEach, vi } from 'vitest';

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

import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { join } from 'path';

// Regression test for a bug that has now shipped twice: btn-reply-encrypted/
// btn-reply-all-encrypted (old ids) no longer matched their own displayed
// labels after a relabel, and the id/label mismatch caused the replyAll
// boolean passed to handleReplyEncrypted() to get silently swapped -- once
// when "fixed" for #20 by matching the boolean to the (misleading) id name
// instead of the label. The ids have since been renamed so each one now
// describes its own button (btn-reply-all-encrypted / btn-reply-sender-encrypted)
// -- this test checks the source text directly (rather than importing the
// module, which requires the full Office.onReady stub story other
// MessageRead.js tests carry) so a future rename or copy-paste error in
// either file fails a test instead of shipping silently a third time.
const htmlSource = readFileSync(join(__dirname, '../web/MessageRead.html'), 'utf8');
const jsSource = readFileSync(join(__dirname, '../web/MessageRead.js'), 'utf8');

describe('Reply / Reply All button id, label, and wiring agreement', () => {
  it('btn-reply-all-encrypted is labeled "Reply All" in the HTML', () => {
    const match = htmlSource.match(/id="btn-reply-all-encrypted"[^>]*>\s*([^<]*)</);
    expect(match).not.toBeNull();
    expect(match[1]).toContain('Reply All');
  });

  it('btn-reply-sender-encrypted is labeled "Reply Sender" in the HTML', () => {
    const match = htmlSource.match(/id="btn-reply-sender-encrypted"[^>]*>\s*([^<]*)</);
    expect(match).not.toBeNull();
    expect(match[1]).toContain('Reply Sender');
  });

  it('wires btn-reply-all-encrypted\'s click to handleReplyEncrypted(true)', () => {
    expect(jsSource).toContain(
      "el('btn-reply-all-encrypted').addEventListener('click', () => handleReplyEncrypted(true));"
    );
  });

  it('wires btn-reply-sender-encrypted\'s click to handleReplyEncrypted(false)', () => {
    expect(jsSource).toContain(
      "el('btn-reply-sender-encrypted').addEventListener('click', () => handleReplyEncrypted(false));"
    );
  });
});

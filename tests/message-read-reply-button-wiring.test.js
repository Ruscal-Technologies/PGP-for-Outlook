import { describe, it, expect } from 'vitest';

// MessageRead.js calls Office.onReady(...) at module load time, so it needs
// a minimal Office stub in place before import -- the no-op stub means the
// wiring callback body never actually runs under test, which is fine here:
// this file tests REPLY_BUTTON_WIRING as a plain exported data structure,
// not the live event-listener wiring itself (see its own module-level
// comment in MessageRead.js for why that data structure exists).
global.Office = { onReady: () => {} };

describe('REPLY_BUTTON_WIRING', () => {
  it('maps each button id to the replyAll value matching its ACTUAL DISPLAYED LABEL, not its legacy id name', async () => {
    // Regression test for a bug that has now shipped twice: the element ids
    // (btn-reply-encrypted / btn-reply-all-encrypted) are historical names
    // that no longer match what the buttons say on screen -- btn-reply-encrypted
    // is labeled "Reply All" and btn-reply-all-encrypted is labeled "Reply
    // Sender" (web/MessageRead.html). A prior fix (#20/#21) "corrected" the
    // wiring by matching booleans to id names instead of labels, silently
    // swapping Reply All and Reply Sender's actual behavior. This test pins
    // the mapping to what the user sees, not what the ids say.
    const { REPLY_BUTTON_WIRING } = await import('../web/MessageRead.js');

    expect(REPLY_BUTTON_WIRING).toEqual([
      { id: 'btn-reply-encrypted', replyAll: true },      // labeled "Reply All"
      { id: 'btn-reply-all-encrypted', replyAll: false }, // labeled "Reply Sender"
    ]);
  });
});

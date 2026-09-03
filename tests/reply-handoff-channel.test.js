import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, it, expect } from 'vitest';
import { getReplyHandoffChannelName, HANDOFF_PENDING_MARKER } from '../web/js/pgp/reply-handoff-channel.js';

describe('getReplyHandoffChannelName', () => {
  it('falls back to the base name when no conversationId is given', () => {
    expect(getReplyHandoffChannelName()).toBe('pgp_reply_handoff');
    expect(getReplyHandoffChannelName('')).toBe('pgp_reply_handoff');
  });

  it('derives a distinct, deterministic name per conversationId, with no possibility of collision', () => {
    const a = getReplyHandoffChannelName('conversation-A');
    const b = getReplyHandoffChannelName('conversation-B');

    expect(a).toBe(`pgp_reply_handoff_${encodeURIComponent('conversation-A')}`);
    expect(b).toBe(`pgp_reply_handoff_${encodeURIComponent('conversation-B')}`);
    expect(a).not.toBe(b);
    expect(getReplyHandoffChannelName('conversation-A')).toBe(a);
  });

  it('URI-encodes characters a real conversationId could contain', () => {
    const id = 'AAQkAD=/weird+id&chars';
    expect(getReplyHandoffChannelName(id)).toBe(`pgp_reply_handoff_${encodeURIComponent(id)}`);
  });
});

describe('HANDOFF_PENDING_MARKER', () => {
  it('has no quotes or HTML-reserved characters (must survive plain-text/HTML round-tripping unescaped)', () => {
    expect(HANDOFF_PENDING_MARKER).not.toMatch(/["'&<>]/);
  });

  it('matches the literal duplicate in Functions/ReplyHandoffRuntime.classic.js exactly', () => {
    // That file can't import this module (no ES modules in its restricted
    // runtime, see HANDOFF_PENDING_MARKER's own docblock) -- if the two
    // ever drift, the classic-Windows notification gate silently never
    // fires. This assertion is the tripwire.
    const classicFilePath = fileURLToPath(new URL('../web/Functions/ReplyHandoffRuntime.classic.js', import.meta.url));
    const classicSource = readFileSync(classicFilePath, 'utf8');
    expect(classicSource).toContain(`'${HANDOFF_PENDING_MARKER}'`);
  });
});

import { describe, it, expect } from 'vitest';
import { getReplyHandoffChannelName } from '../web/js/pgp/reply-handoff-channel.js';

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
